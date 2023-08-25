public class Script : ScriptBase
{
    public override async Task<HttpResponseMessage> ExecuteAsync()
    {
        // Check if the operation ID matches what is specified in the OpenAPI definition of the connector
        if (this.Context.OperationId == "CreateExcelSession")
        {
            return await this.HandleForwardAndTransformOperation().ConfigureAwait(false);
        }
        if (this.Context.OperationId == "ConvertDataForXMultiple")
        {
            return await this.HandleConvertXMultiple().ConfigureAwait(false);
        }
        if(this.Context.OperationId == "GetAddressRangeBatch")
        {
            return await this.HandleGetAddressRangeBatch().ConfigureAwait(false);
        }
        if(this.Context.OperationId == "ValidateDataRegex")
        {
            return await this.HandleValidateDataRegex().ConfigureAwait(false);
        }
        if(this.Context.OperationId == "FindDuplicates")
        {
            return await this.HandleDuplicates().ConfigureAwait(false);
        }
        if(this.Context.OperationId == "ChunkArray")
        {
            return await this.HandleChunkArray().ConfigureAwait(false);
        }

        // Handle an invalid operation ID
        HttpResponseMessage response = new HttpResponseMessage(HttpStatusCode.BadRequest);
        response.Content = CreateJsonContent($"Unknown operation ID '{this.Context.OperationId}'");
        return response;
    }

    private async Task<HttpResponseMessage> HandleForwardAndTransformOperation()
    {
        // Use the context to forward/send an HTTP request
        HttpResponseMessage response = await this.Context.SendAsync(this.Context.Request, this.CancellationToken).ConfigureAwait(continueOnCapturedContext: false);

        // Do the transformation if the response was successful, otherwise return error responses as-is
        if (response.IsSuccessStatusCode)
        {
            string location = response.Headers.TryGetValues("location", out var values) ? values.FirstOrDefault() : "";
        
            if(location != "https://graph.microsoft.com/") {
                var newResult = new JObject{ ["location"] = location, };            
                response.Content = CreateJsonContent(newResult.ToString());
            }
        }
        return response;
    }

    private async Task<HttpResponseMessage> HandleConvertXMultiple()
    {
        // Use the context to forward/send an HTTP request
        var tableName = this.Context.Request.Headers.TryGetValues("x-TableName", out var tblValues) ? tblValues.FirstOrDefault() : "";
        var columnInput = this.Context.Request.Headers.TryGetValues("x-Columns", out var colValues) ? colValues.FirstOrDefault() : "";
        var trackerColumn = this.Context.Request.Headers.TryGetValues("x-TrackerColumn", out var trColumn) ? trColumn.FirstOrDefault(): "";
        var trackerValue = this.Context.Request.Headers.TryGetValues("x-TrackerValue", out var trValue) ? trValue.FirstOrDefault(): "";
        var guidColumn = this.Context.Request.Headers.TryGetValues("x-GuidColumn", out var guidCol) ? guidCol.FirstOrDefault(): "";
        var primaryColumn = this.Context.Request.Headers.TryGetValues("x-PrimaryColumn", out var prCol) ? prCol.FirstOrDefault(): "";
        var mergeColumns = this.Context.Request.Headers.TryGetValues("x-MergeColumns", out var mergeCol) ? mergeCol.FirstOrDefault(): "";
        var mergeColumnName = this.Context.Request.Headers.TryGetValues("x-MergeColumnName", out var mergeColName) ? mergeColName.FirstOrDefault(): "";
        var columnDelimiter = this.Context.Request.Headers.TryGetValues("x-ColumnDelimiter", out var colDelim) ? colDelim.FirstOrDefault(): "|";
        var addColumnNames = this.Context.Request.Headers.TryGetValues("x-AdditionalColumnNames", out var addColName) ? addColName.FirstOrDefault(): "";
        var addColumnValues = this.Context.Request.Headers.TryGetValues("x-AdditionalColumnValues", out var addColValue) ? addColValue.FirstOrDefault(): "";

        var ttlInput = this.Context.Request.Headers.TryGetValues("x-TTL", out var ttlValue) ? ttlValue.FirstOrDefault(): "-1";
        int ttl = int.Parse(ttlInput);

        var dataInput = await this.Context.Request.Content.ReadAsStringAsync().ConfigureAwait(false);
        var columns = new List<string>(columnInput.Split(','));        
        var batchData = JsonConvert.DeserializeObject<List<List<string>>>(dataInput);        
        var output = new List<Dictionary<string, string>>(batchData.Count);

        foreach (var row in batchData)
        {
            int additionalCol = 1;
            bool useGuid = !string.IsNullOrEmpty(guidColumn) && !string.IsNullOrEmpty(primaryColumn);
            //if GuidColumn and PrimaryColumn is provided
            if(useGuid) additionalCol++;

            //if ttl is provided
            if(ttl > 0) additionalCol++;
            var primaryNameValue = string.Empty;

            //if additional columns is provided
            string[] additionalCols = addColumnNames.Split(',');
            string[] additionalVals = addColumnValues.Split(columnDelimiter.ToCharArray()[0]);
            if(!string.IsNullOrEmpty(addColumnNames)) {
                additionalCol += additionalCols.Length;
            }

            var item = new Dictionary<string, string>(columns.Count + additionalCol);
            item["@odata.type"] = $"Microsoft.Dynamics.CRM.{ tableName }";   

            if(trackerColumn.Length > 0) item[trackerColumn] = trackerValue;
            if(ttl > 0) item["ttlinseconds"] = ttl.ToString();
            
            var columnEnumerator = columns.GetEnumerator();
            var rowEnumerator = row.GetEnumerator();
            while (columnEnumerator.MoveNext() && rowEnumerator.MoveNext())
            {
                item[columnEnumerator.Current] = rowEnumerator.Current;
                if(useGuid && columnEnumerator.Current == primaryColumn )
                {
                    byte[] hashCode;
                    using (MD5 md5 = MD5.Create())
                    {
                        hashCode = md5.ComputeHash(Encoding.UTF8.GetBytes(rowEnumerator.Current));
                    }
                    item[guidColumn] = (new Guid(hashCode)).ToString();
                }
            }
            if(mergeColumns.Length > 0)
            {
                string[] mColumns = mergeColumns.Split(',');
                StringBuilder mergedValue = new StringBuilder();
                foreach(string column in mColumns)
                {
                    string colValue = (string)item[column];
                    mergedValue.Append(colValue).Append(columnDelimiter);
                }
                if(mergedValue.Length > 0)
                {
                    mergedValue.Length -= columnDelimiter.Length;
                    item[mergeColumnName] = mergedValue.ToString();
                }
            }
            if(additionalCols.Length > 0)
            {
                for(int i = 0; i < additionalCols.Length; i++)
                {
                    item[additionalCols[i]] = additionalVals[i];
                }
            }
            output.Add(item);
        }        
        var result = JsonConvert.SerializeObject(output, Newtonsoft.Json.Formatting.None);            
        HttpResponseMessage response = new HttpResponseMessage(HttpStatusCode.OK);
        response.Content = CreateJsonContent(result.ToString());
        return response;
    }

    private async Task<HttpResponseMessage> HandleGetAddressRangeBatch()
    {
        // Use the context to forward/send an HTTP request
        var address = this.Context.Request.Headers.TryGetValues("x-Address", out var addressValue) ? addressValue.FirstOrDefault() : "";
        var batchSize = int.Parse(this.Context.Request.Headers.TryGetValues("x-BatchSize", out var bSize) ? bSize.FirstOrDefault() : "100");
        var firstRow = int.Parse(this.Context.Request.Headers.TryGetValues("x-StartRow", out var sRow) ? sRow.FirstOrDefault(): "1");
        if(firstRow <= 1) firstRow = 1;

        List<string> batches = new List<string>();
        string[] rangeParts = address.Split(':');
        string startCell = rangeParts[0];
        string endCell = rangeParts[1];

        Match startCellNum = Regex.Match(startCell, @"\d+");
        int startRow = int.Parse(startCellNum.Value) + (firstRow - 1);

        Match endRowNum = Regex.Match(endCell, @"\d+");
        int endRow = int.Parse(endRowNum.Value);
        int numRows = endRow - startRow + 1;
        int numBatches = (numRows + batchSize - 1) / batchSize;
    
        for (int i = 0; i < numBatches; i++)
        {
            int startBatchRow = startRow + i * batchSize;
            int endBatchRow = Math.Min(startBatchRow + batchSize - 1, endRow);
            string batchRange = $"{startCell.Replace(startCellNum.Value, "")}{startBatchRow}:{endCell.Replace(endRowNum.Value, "")}{endBatchRow}";
            batches.Add(batchRange);
        }
        
        var result = new {
            Batch = batches, TotalRows = numRows
        };

        HttpResponseMessage response = new HttpResponseMessage(HttpStatusCode.OK);
        response.Content = CreateJsonContent(JsonConvert.SerializeObject(result).ToString());
        return response;
    }

    private async Task<HttpResponseMessage> HandleValidateDataRegex()
    {
        var regexInput = this.Context.Request.Headers.TryGetValues("Regex", out var regexValue) ? regexValue.FirstOrDefault() : "";
        var colIdxInput = this.Context.Request.Headers.TryGetValues("ColumnIndex", out var colIndexValue) ? colIndexValue.FirstOrDefault(): "-1";
        int colIdx = int.Parse(colIdxInput);

        var firstCellAddress = this.Context.Request.Headers.TryGetValues("FirstCellAddress", out var firstCellAddressValue) ? firstCellAddressValue.FirstOrDefault(): "";
                
        string dataInput = await this.Context.Request.Content.ReadAsStringAsync().ConfigureAwait(false);       
        JArray dataArray = JArray.Parse(dataInput);
        Regex regex = new Regex(regexInput);
        List<CellData> results = new List<CellData>();

        for(int i = 0; i < dataArray.Count; i++)
        {
            if(colIdx == -1)
            {
                for(int j = 0; j < dataArray[i].Count(); j++)
                {
                    string dataItem = dataArray[i][j].ToString();
                    if(!regex.IsMatch(dataItem)) results.Add(new CellData(firstCellAddress, j, i, dataItem));
                }                
            }
            else
            {
                string dataItem = dataArray[i][colIdx].ToString();
                if(!regex.IsMatch(dataItem)) results.Add(new CellData(firstCellAddress, colIdx, i, dataItem));
            }
        } 
        HttpResponseMessage response = new HttpResponseMessage(HttpStatusCode.OK);
        response.Content = CreateJsonContent(JsonConvert.SerializeObject(results).ToString());
        return response;
    }

    private async Task<HttpResponseMessage> HandleDuplicates()
    {
        var colIdxInput = this.Context.Request.Headers.TryGetValues("ColumnIndex", out var colIndexValue) ? colIndexValue.FirstOrDefault(): "-1";
        var caseSensitiveInput = this.Context.Request.Headers.TryGetValues("CaseSensitive", out var caseSensitiveValue)? caseSensitiveValue.FirstOrDefault(): "true";

        int colIdx = int.Parse(colIdxInput);
        bool caseSensitive = bool.Parse(caseSensitiveInput);

        var firstCellAddress = this.Context.Request.Headers.TryGetValues("FirstCellAddress", out var firstCellAddressValue) ? firstCellAddressValue.FirstOrDefault(): "";
        string dataInput = await this.Context.Request.Content.ReadAsStringAsync().ConfigureAwait(false);
        JArray dataArray = JArray.Parse(dataInput);
        var dataset = new Dictionary<string, CellData>();
        var duplicates = new Dictionary<string, List<CellData>>();

        for(int i = 0; i < dataArray.Count; i++)
        {
            if(colIdx > -1)
            {
                string dataItem = dataArray[i][colIdx].ToString();
                if(!caseSensitive) dataItem = dataItem.ToLower();
                if(dataset.Keys.Contains(dataItem))
                {
                    if(!duplicates.Keys.Contains(dataItem))
                    {
                        CellData org = dataset[dataItem];
                        duplicates.Add(dataItem, new List<CellData>());
                        duplicates[dataItem].Add(new CellData(firstCellAddress, org.Col, org.Row, dataItem));
                    }
                    duplicates[dataItem].Add(new CellData(firstCellAddress, colIdx, i, dataItem));
                }
                else
                {
                    dataset.Add(dataItem, new CellData(firstCellAddress, colIdx, i, dataItem));
                }                
            }
        } 

        List<ResultSet> results = new List<ResultSet>();
        foreach(var dup in duplicates)
        {
            ResultSet rs = new ResultSet(dup.Key);
            foreach(CellData cell in dup.Value)
            {
                rs.Cells.Add(cell.Cell);
            }
            results.Add(rs);
        }

        HttpResponseMessage response = new HttpResponseMessage(HttpStatusCode.OK);
        response.Content = CreateJsonContent(JsonConvert.SerializeObject(results).ToString());
        return response;
    }

    private async Task<HttpResponseMessage> HandleChunkArray()
    {
        var chunkSizeInput = this.Context.Request.Headers.TryGetValues("ChunkSize", out var chunkSizeValue) ? chunkSizeValue.FirstOrDefault(): "200";
        int chunkSize = int.Parse(chunkSizeInput);
        
        var dataInput = await this.Context.Request.Content.ReadAsStringAsync().ConfigureAwait(false);
        JArray dataArray = JArray.Parse(dataInput);
        List<List<string>> inputList = dataArray.ToObject<List<List<string>>>();
        List<List<List<string>>> result = new List<List<List<string>>>();
        for(int i = 0; i < inputList.Count; i += chunkSize)
        {
            List<List<string>> chunk = inputList.GetRange(i, Math.Min(chunkSize, inputList.Count - i));
            result.Add(chunk);
        }

        HttpResponseMessage response = new HttpResponseMessage(HttpStatusCode.OK);
        response.Content = CreateJsonContent(JArray.FromObject(result).ToString());
        return response;
    }
}

internal class ResultSet
{
    public ResultSet(string value)
    {
        this.Value = value;
        this.Cells = new List<string>();
    }
    public string Value { get; set; }
    public List<string> Cells { get; set; }
}

internal class CellData
{
    private string start;
    private int col; 
    private int row;
    private string data;
    public CellData(string startAddress, int col, int row, string data)
    {
        this.start = startAddress;
        this.col = col;
        this.row = row;
        this.Data = data;
    }
    
    [JsonIgnore]
    public int Col { get { return this.col; }}

    [JsonIgnore]
    public int Row { get { return this.row; }}

    public string Cell { 
        get {
            Match colMatch = Regex.Match(this.start, @"[0-9]+$");
            int startRow = int.Parse(colMatch.Value);
            string startCol = this.start.Replace(startRow.ToString(), "");

            return $"{ this.CalculateEndColumn(startCol, this.col) }{this.row + startRow}";  
        }
    }
    public string Data { get; set; }

    private string CalculateEndColumn(string startingColumn, int totalColumns)
    {            
        return this.GetColumnName(this.GetColumnValue(startingColumn) + totalColumns);
    }

    private int GetColumnValue(string column)
    {
        int value = 0;
        foreach (char c in column) { value = value * 26 + (c - 'A' + 1); }
        return value;
    }

    private string GetColumnName(int value)
    {
        string columnName = "";
        while (value > 0)
        {
            int remainder = (value - 1) % 26;
            char columnChar = (char)('A' + remainder);
            columnName = columnChar + columnName;
            value = (value - 1) / 26;
        }
        return columnName;
    }
}
