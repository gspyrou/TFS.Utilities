using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using HtmlAgilityPack;
using Newtonsoft.Json.Linq;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Net.Http;
using System.Net.Http.Headers;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using System.Configuration;

namespace ConsoleApp3
{
    class Program
    {
        static void Main(string[] args)
        {
            Console.WriteLine("Please enter the full path to the Excel file:");
            var path = Console.ReadLine();
            Task.Run(() => MainAsync(path));
            Console.ReadLine();
        }
        static async Task MainAsync(string excelFilePath)
        {
            try
            {
                var personalaccesstoken = ConfigurationManager.AppSettings["PersonalAccessToken"];

                var document = SpreadsheetDocument.Open(excelFilePath, true);
                SharedStringTable sharedStringTable = document.WorkbookPart.SharedStringTablePart.SharedStringTable;
                Console.WriteLine("Getting TestCases data ....");
                foreach (WorksheetPart worksheetPart in document.WorkbookPart.WorksheetParts)
                {
                    foreach (SheetData sheetData in worksheetPart.Worksheet.Elements<SheetData>())
                    {
                        if (sheetData.HasChildren)
                        {
                            foreach (Row row in sheetData.Elements<Row>())
                            {
                                // Loop through each of the cells in the current row.
                                var cells = row.Elements<Cell>().Take(1);
                                foreach (var cell in cells)
                                {
                                    // Here is where you would do something with the values of the spreadsheet.
                                    int Id = 0;
                                    if (int.TryParse(GetValue(document, cell), out Id))
                                    {
                                        Console.WriteLine("TestCaseId : {0}", Id);
                                        using (HttpClient client = new HttpClient())
                                        {
                                            client.DefaultRequestHeaders.Accept.Add(
                                                new System.Net.Http.Headers.MediaTypeWithQualityHeaderValue("application/json"));

                                            client.DefaultRequestHeaders.Authorization = new AuthenticationHeaderValue("Basic",
                                                Convert.ToBase64String(
                                                    System.Text.ASCIIEncoding.ASCII.GetBytes(
                                                        string.Format("{0}:{1}", "", personalaccesstoken))));

                                            using (HttpResponseMessage response = await client.GetAsync(
                                                        String.Format("{0}/_apis/wit/workItems/{1}", ConfigurationManager.AppSettings["RestApiBaseUri"], Id)))
                                            {
                                                response.EnsureSuccessStatusCode();
                                                string responseBody = await response.Content.ReadAsStringAsync();
                                                dynamic data = JObject.Parse(responseBody);

                                                DocumentFormat.OpenXml.Spreadsheet.Cell newcell = new DocumentFormat.OpenXml.Spreadsheet.Cell();
                                                newcell.DataType = DocumentFormat.OpenXml.Spreadsheet.CellValues.String;
                                                newcell.CellValue = new DocumentFormat.OpenXml.Spreadsheet.CellValue(
                                                    (string)(data.fields["System.Description"].Value).Replace("<div>", "").Replace("</div>", Environment.NewLine).Replace("&quot;", "\"").Replace("&nbsp;", "")); //
                                                row.Append(newcell);
                                            }
                                        }
                                    }
                                }

                            }
                        }
                    }

                    Console.WriteLine("Done .");
                }
                document.Save();
                document.Close();


            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.ToString());
                Console.Read();
            }
        }

        private static string GetValue(SpreadsheetDocument doc, Cell cell)
        {
            string value = cell.CellValue.InnerText;
            if (cell.DataType != null && cell.DataType.Value == CellValues.SharedString)
            {
                return doc.WorkbookPart.SharedStringTablePart.SharedStringTable.ChildElements.GetItem(int.Parse(value)).InnerText;
            }
            return value;
        }

    }
}
