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
            //var responseString = await response.Content.ReadAsStringAsync().ConfigureAwait(continueOnCapturedContext: false);
        
            // Example case: response string is some JSON object
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
        
        var dataInput = await this.Context.Request.Content.ReadAsStringAsync().ConfigureAwait(false);
        
        var columns = JsonConvert.DeserializeObject<List<string>>(columnInput);
      
        var batchData = JsonConvert.DeserializeObject<List<List<string>>>(dataInput);
        
        var output = new List<Dictionary<string, string>>(batchData.Count);

        foreach (var row in batchData)
        {
            int additionalCol = 1;
            bool useGuid = !string.IsNullOrEmpty(guidColumn) && !string.IsNullOrEmpty(primaryColumn);
            //if GuidColumn and PrimaryColumn is provided
            if(useGuid) additionalCol++;
            var primaryNameValue = string.Empty;

            var item = new Dictionary<string, string>(columns.Count + additionalCol);
            item["@odata.type"] = $"Microsoft.Dynamics.CRM.{ tableName }";           
            item[trackerColumn] = trackerValue;
            
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
        int startRow = int.Parse(startCell.Substring(1)) + (firstRow - 1);
        int endRow = int.Parse(endCell.Substring(1));
        int numRows = endRow - startRow + 1;
        int numBatches = (numRows + batchSize - 1) / batchSize;
    
        for (int i = 0; i < numBatches; i++)
        {
            int startBatchRow = startRow + i * batchSize;
            int endBatchRow = Math.Min(startBatchRow + batchSize - 1, endRow);
            string batchRange = $"{startCell.Substring(0, 1)}{startBatchRow}:{endCell.Substring(0, 1)}{endBatchRow}";
            batches.Add(batchRange);
        }
        
        var result = new {
            Batch = batches, TotalRows = numRows
        };

        HttpResponseMessage response = new HttpResponseMessage(HttpStatusCode.OK);

        response.Content = CreateJsonContent(JsonConvert.SerializeObject(result).ToString());
        return response;
    }
}
