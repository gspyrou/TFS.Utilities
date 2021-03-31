using Newtonsoft.Json.Linq;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Net.Http;
using System.Net.Http.Headers;
using System.Text;
using System.Threading.Tasks;

namespace TestRunsExport
{
    class Program
    {

        static void Main(string[] args)
        {
            Task.Run(() => MainAsync());
            Console.ReadLine();
        }
        static async Task MainAsync()
        {
            try
            {
                var personalaccesstoken = "hwgvv4rtd6mzk5xs47qsfubvtd6jjy4hz3kehog5xs34v2gpkw4a";
                using (HttpClient client = new HttpClient())
                {
                    client.DefaultRequestHeaders.Accept.Add(
                        new System.Net.Http.Headers.MediaTypeWithQualityHeaderValue("application/json"));

                    client.DefaultRequestHeaders.Authorization = new AuthenticationHeaderValue("Basic",
                        Convert.ToBase64String(
                            System.Text.ASCIIEncoding.ASCII.GetBytes(
                                string.Format("{0}:{1}", "", personalaccesstoken))));

                    using (HttpResponseMessage response = await client.GetAsync(
                                "http://era-srd-dev:8080/tfs/ERA-SRD_Collection/SRD/_apis/test/runs?includeRunDetails=true"))
                                
                    {
                        response.EnsureSuccessStatusCode();
                        string responseBody = await response.Content.ReadAsStringAsync();
                        dynamic data = JObject.Parse(responseBody);
                        foreach (var item in data.value)
                        {    
                            Console.WriteLine(item.id);
                        }
                        //DocumentFormat.OpenXml.Spreadsheet.Cell newcell = new DocumentFormat.OpenXml.Spreadsheet.Cell();
                        //newcell.DataType = DocumentFormat.OpenXml.Spreadsheet.CellValues.String;
                        //newcell.CellValue = new DocumentFormat.OpenXml.Spreadsheet.CellValue(
                        //    (string)(data.fields["System.Description"].Value).Replace("<div>", "").Replace("</div>", Environment.NewLine).Replace("&quot;", "\"").Replace("&nbsp;", "")); //
                        //row.Append(newcell);
                    }
                }
            }

            
            catch (Exception x)
            {
                Console.WriteLine(x.Message);
            }
        }

    }
}
