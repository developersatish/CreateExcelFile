using ExcelLibrary.SpreadSheet;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace App
{
    public class DataModel
    {
        public int ID { get; set; }
        public string FirstName { get; set; }
        public string LastName { get; set; }        
        public DateTime Date { get; set; }
    }
    class Program
    {
        static void Main(string[] args)
        {

            Do();
        }
        public static void Do()
        {
            try
            {
                using (var db = new welldbEntities())
                {
                    IEnumerable<DataModel> data = db.tblLandMen.Select(s => new DataModel()
                    {
                        FirstName = s.FirstName,
                        LastName = s.LastName,
                        Date = s.CreateDate,
                        ID = s.ID
                    }).ToList();
                    string file = @"E:\Book1.xls";
                    //if (!File.Exists(file))
                    //    File.Create(file);

                    Workbook workbook = new Workbook();
                    Worksheet worksheet = new Worksheet("First Sheet");

                    //Write Header
                    DataModel tbl = new DataModel();

                    var headers = tbl.GetType().GetProperties();
                    int head = 0;
                    foreach (var prop in headers)
                    {
                        // Console.WriteLine("{0}={1}", prop.Name, prop.GetValue(item, null));
                        worksheet.Cells[0, head] = new Cell(prop.Name);
                        head++;
                    }

                    int cell = 1, row = 0;
                    foreach (var item in data)
                    {
                        row = 0;
                        var pro = item.GetType().GetProperties();
                        foreach (var prop in pro)
                        {
                            var value = prop.GetValue(item, null);

                            Cell cellData = null;

                            if (prop.PropertyType == typeof(DateTime))
                                cellData = new Cell((DateTime)value);
                            else
                                cellData = new Cell(value ?? "");

                            worksheet.Cells[cell, row] = cellData;
                            row++;
                        }

                        cell++;
                    }
                    // Console.ReadLine();

                    //for (int i = 0; i < data.Count(); i++)
                    //{
                    //    var pro = data[i].GetType().GetProperties();
                    //}

                    workbook.Worksheets.Add(worksheet);
                    workbook.Save(file);
                }



                //worksheet.Cells[2, 0] = new Cell(9999999);
                //worksheet.Cells[3, 3] = new Cell((decimal)3.45);
                //worksheet.Cells[2, 2] = new Cell("Text string");              

            }
            catch (Exception ex)
            {

                throw;
            }
        }
    }
}

//create new xls file string file = "C:\newdoc.xls"; Workbook workbook = new Workbook(); Worksheet worksheet = new Worksheet("First Sheet"); worksheet.Cells[0, 1] = new Cell((short)1); worksheet.Cells[2, 0] = new Cell(9999999); worksheet.Cells[3, 3] = new Cell((decimal)3.45); worksheet.Cells[2, 2] = new Cell("Text string"); worksheet.Cells[2, 4] = new Cell("Second string"); worksheet.Cells[4, 0] = new Cell(32764.5, "#,##0.00"); worksheet.Cells[5, 1] = new Cell(DateTime.Now, @"YYYY-MM-DD"); worksheet.Cells.ColumnWidth[0, 1] = 3000; workbook.Worksheets.Add(worksheet); workbook.Save(file);

// open xls file Workbook book = Workbook.Load(file); Worksheet sheet = book.Worksheets[0];

// traverse cells foreach (Pair, Cell> cell in sheet.Cells) { dgvCells[cell.Left.Right, cell.Left.Left].Value = cell.Right.Value; }

// traverse rows by Index for (int rowIndex = sheet.Cells.FirstRowIndex; rowIndex <= sheet.Cells.LastRowIndex; rowIndex++) { Row row = sheet.Cells.GetRow(rowIndex); for (int colIndex = row.FirstColIndex; colIndex <= row.LastColIndex; colIndex++) { Cell cell = row.GetCell(colIndex); } } ```