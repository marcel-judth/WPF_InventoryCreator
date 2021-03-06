using ExcelDataReader;
using Newtonsoft.Json;
using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading.Tasks;
using WPF_InventoryListCreator.Models;
using Excel = Microsoft.Office.Interop.Excel;

namespace WPF_InventoryListCreator.Code
{
    class Utile
    {
        public List<Article> LoadArticles(string filename)
        {
            List<Article> results = new List<Article>();

            using (FileStream stream = File.Open(filename, FileMode.Open))
            {
                IExcelDataReader excelReader = ExcelReaderFactory.CreateReader(stream);
                var result = excelReader.AsDataSet(new ExcelDataSetConfiguration()
                {
                    ConfigureDataTable = (_) => new ExcelDataTableConfiguration()
                    {
                        UseHeaderRow = true
                    }
                });

                for (int row = 0; row < result.Tables[0].Rows.Count; row++)
                {
                    results.Add(new Article
                    {
                        Number = GetValueFromTable(result, "Artikelnummer", row),
                        Description = GetValueFromTable(result, "Bezeichnung", row),
                        BatchID = GetValueFromTable(result, "ChargeIdentnummer", row)
                    });
                }
            }

            SaveArticlesToSettings(results);

            return results;
        }

        private string GetValueFromTable(DataSet result, string column, int row)
        {
            return result.Tables[0].Rows[row][column].ToString().Trim();
        }

        internal List<InventoryItem> LoadInventoryItems(string filename)
        {
            List<InventoryItem> results = new List<InventoryItem>();

            using (var reader = new StreamReader(filename))
            {
                reader.ReadLine();

                while (!reader.EndOfStream)
                {
                    var line = reader.ReadLine();
                    var values = line.Split('\t');

                    if (values.Length > 3)
                    {
                        results.Add(new InventoryItem
                        {
                            ArticleNr = values[0],
                            BatchID = values[1],
                            Quantity = values[2],
                            Unit = values[3],
                            Date = DateTime.Now.ToString("dd.MM.yyyy")
                        });
                    }
                }
            }
            return results;
        }

        internal void ExportInventoryList(List<InventoryItem> allItems, string filename)
        {
            //var table = ConvertToDataTable(allItems);

            //GenerateExcel(table, filename);

            string csv = GenerateCSVString(allItems);

            File.WriteAllText(filename, csv);
        }

        public string GenerateCSVString(List<InventoryItem> allItems)
        {
            var csv = new StringBuilder();
            csv.AppendLine("Artikelnummer;Menge1;Inventurdatum;Zahlliste;Arbeitnehmer");

            foreach (var item in allItems)
                csv.AppendLine($"{item.article?.Number};{item.Quantity};{item.Date};;{item.BatchID}");

            return csv.ToString();
        }

        public DataTable ConvertToDataTable(List<InventoryItem> models)
        {
            DataTable table = new DataTable();
            table.Columns.Add("Artikelnummer", typeof(string));
            table.Columns.Add("Menge1", typeof(string));
            table.Columns.Add("Inventurdatum", typeof(string));
            table.Columns.Add("Zahlliste", typeof(string));
            table.Columns.Add("Arbeitnehmer", typeof(string));

            foreach (var item in models)
                table.Rows.Add(item.article.Number, item.Quantity, item.Date, "", item.BatchID);

            return table;
        }

        public void GenerateExcel(DataTable dataTable, string path)
        {

            DataSet dataSet = new DataSet();
            dataSet.Tables.Add(dataTable);

            // create a excel app along side with workbook and worksheet and give a name to it  
            Excel.Application excelApp = new Excel.Application();
            Excel.Workbook excelWorkBook = excelApp.Workbooks.Add();
            Excel._Worksheet xlWorksheet = excelWorkBook.Sheets[1];
            Excel.Range xlRange = xlWorksheet.UsedRange;

            foreach (DataTable table in dataSet.Tables)
            {
                //Add a new worksheet to workbook with the Datatable name  
                Excel.Worksheet excelWorkSheet = excelWorkBook.Sheets.Add();
                excelWorkSheet.Name = table.TableName;

                // add all the columns  
                for (int i = 1; i < table.Columns.Count + 1; i++)
                {
                    excelWorkSheet.Cells[1, i] = table.Columns[i - 1].ColumnName;
                }

                // add all the rows  
                for (int j = 0; j < table.Rows.Count; j++)
                {
                    for (int k = 0; k < table.Columns.Count; k++)
                    {
                        excelWorkSheet.Cells[j + 2, k + 1] = table.Rows[j].ItemArray[k].ToString();
                    }
                }
            }

            // excelWorkBook.Save(); -> this is save to its default location  
            excelApp.DisplayAlerts = false;
            excelWorkBook.SaveCopyAs(path); // -> this will do the custom  
            excelWorkBook.Close();
            excelApp.Quit();
        }

        internal void CreateInventoryList(List<Article> allArticles, List<InventoryItem> allItems)
        {
            //todo make error when no article

            foreach (var item in allItems)
            {
                var article = allArticles.FirstOrDefault(a => a.BatchID?.ToLower() == item.BatchID?.ToLower() && a.Number.StartsWith(item.ArticleNr));

                item.article = article;
            }
        }

        public List<Article> GetArticlesFromSettings()
        {
            string articles = Properties.Settings.Default.ListArticles;

            if (!string.IsNullOrEmpty(articles))
                return JsonConvert.DeserializeObject<List<Article>>(articles);

            return new List<Article>();
        }

        private void SaveArticlesToSettings(List<Article> articles)
        {
            Properties.Settings.Default.ListArticles = JsonConvert.SerializeObject(articles);
            Properties.Settings.Default.Save();
        }
    }
}
