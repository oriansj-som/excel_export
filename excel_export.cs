using OfficeOpenXml;
using OfficeOpenXml.Style;
using sqlite_Example;
using System;
using System.Collections.Generic;
using System.Data;
using System.Drawing;
using System.IO;
using System.Text;


namespace Export_Excel
{
    class excel_export
    {
        private SQLiteDatabase mydatabase;
        private ExcelWorksheet ws;
        private ExcelPackage Report;
        private FileInfo output;

        private void require(Boolean b, String msg)
        {
            if (!b)
            {
                Console.WriteLine("Failed requirement");
                Console.WriteLine(msg);
                Environment.Exit(1);
            }
        }

        public void initialize(string output_file, string Database)
        {

            output = new FileInfo(output_file);
            Report = new ExcelPackage();

            if (!File.Exists(Database))
            {
                Console.WriteLine("Unable to find " + Database);
                Environment.Exit(4);
            }

            // Connect to Database
            mydatabase = new SQLiteDatabase(Database);            
        }

        public void Cleanup()
        {
            mydatabase.SQLiteDatabase_Close();
            MultiSave();
        }

        private void Special_Coloring(String rule, int row, int col, Color c)
        {
            ExcelAddress _formatRangeAddress = new ExcelAddress(row, col, row, col);
            var _cond4 = ws.ConditionalFormatting.AddExpression(_formatRangeAddress);
            _cond4.Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
            _cond4.Style.Fill.BackgroundColor.Color = c;
            _cond4.Formula = String.Format(rule, _formatRangeAddress.ToString());
        }

        private void make_header(String contents, int row, int col, Color c)
        {
            ExcelBorderStyle a = ExcelBorderStyle.Medium;
            try
            {
                ws.Cells[row, col].Value = contents;
            }
            catch (Exception e)
            {
                Console.WriteLine(e);
                Environment.Exit(7);
            }

            ws.Cells[row, col].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
            ws.Cells[row, col].Style.Fill.BackgroundColor.SetColor(c);
            ws.Cells[row, col].Style.Border.Top.Style = a;
            ws.Cells[row, col].Style.Border.Bottom.Style = a;
            ws.Cells[row, col].Style.Border.Left.Style = a;
            ws.Cells[row, col].Style.Border.Right.Style = a;
        }

        private void insert_content(string contents, int row, int col)
        {
            try
            {
                ws.Cells[row, col].Value = Decimal.Parse(contents);
            }
            catch
            {
                ws.Cells[row, col].Value = contents;
            }
        }

        public void Generate(List<string> list)
        {
            foreach (string tables_list in list)
            {
                string[] tables = tables_list.Split(',');
                List<int> widths = new List<int>();
                foreach(string t in tables)
                {
                    try
                    {
                        int i = int.Parse(t);
                        widths.Add(i);
                    }
                    catch
                    {
                        Console.WriteLine(string.Format("Starting to process table: {0}", t));
                    }
                }

                string table = tables[0];
                require(mydatabase.VerifyTableExists(table), string.Format("There is no table named: {0} in the database", table));
                ws = Report.Workbook.Worksheets.Add(table);
                populate(table, widths);
            }
        }

        private void populate(string table, List<int> column_widths)
        {
            int row = 1;
            int col = 1;
            DataTable ds = mydatabase.GetDataTable(string.Format("SELECT * FROM {0}", table));
            List<string> headers = new List<string>();

            foreach(DataColumn dc in ds.Columns)
            {
                make_header(dc.ColumnName, row, col, Color.Gray);
                headers.Add(dc.ColumnName);
                try
                {
                    ws.Column(col).Width = column_widths[col - 1];
                }
                catch (Exception ex)
                {
                    Console.WriteLine(string.Format("No width set for column: {0} (column: {1}), setting to default", dc.ColumnName, col));
                }
                col = col + 1;
            }

            row = row + 1;
            foreach(DataRow dr in ds.Rows)
            {
                col = 1;
                foreach (string s in headers)
                {
                    try
                    {
                        insert_content(dr[s].ToString(), row, col);
                    }
                    catch (Exception ex)
                    {
                        Console.WriteLine(ex);
                        Environment.Exit(8);
                    }
                    col = col + 1;
                }

                row = row + 1;
            }
        }

        // Enable us to incrementally save the workbook
        public void MultiSave()
        {
            var holdingstream = new MemoryStream();
            Report.Stream.CopyTo(holdingstream);
            Report.SaveAs(output);
            holdingstream.SetLength(0);
            Report.Stream.Position = 0;
            Report.Stream.CopyTo(holdingstream);
            Report.Load(holdingstream);
        }
    }
}
