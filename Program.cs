using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Excel = Microsoft.Office.Interop.Excel;
namespace GenericList2XL
{
    class Program
    {
        static void Main(string[] args)
        {
            // Create a list of parts.
            List<Part> parts = new List<Part>();

            // Add parts to the list.
            parts.Add(new Part() { PartId = 1234 , PartName = "Test1 arm", PartAmount=3232});
            parts.Add(new Part() { PartId = 1334, PartName = "chain ring", PartAmount = 3232 });
            parts.Add(new Part() { PartId = 1434, PartName = "regular seat", PartAmount = 3232 });
            parts.Add(new Part() { PartId = 1444, PartName = "banana seat", PartAmount = 3232 });
            parts.Add(new Part() { PartId = 1534,PartName = "cassette", PartAmount = 3232 });
            parts.Add(new Part() { PartId = 1634, PartName = "shift lever", PartAmount = 3232 });
            XportExcel(GetDataTable.ToDataTable(parts));
        }

        public class Part
        {
            public int PartId { get; set; }

            public string PartName { get; set; }
            public double PartAmount { get; set; }
        }

        public static void XportExcel(DataTable Tbl)
        {
            Excel.Application xlSamp = new Microsoft.Office.Interop.Excel.Application();
            if (xlSamp == null) 
            { 
                Console.WriteLine("Excel is not Insatalled"); 
                Console.ReadKey(); 
                return; 
            }
            Excel.Workbook xlWorkBook;
            Excel.Worksheet xlWorkSheet;
            Excel.Range xlWorkSheetRange;
            
            object misValue = System.Reflection.Missing.Value;
            xlWorkBook = xlSamp.Workbooks.Add(misValue);
            xlWorkSheet = (Excel.Worksheet)xlWorkBook.Worksheets.get_Item(1);

            for (int i = 0; i < Tbl.Columns.Count; i++)
            {
                xlWorkSheet.Cells[1, (i + 1)] = Tbl.Columns[i].ColumnName;
            }

            for (int i = 0; i < Tbl.Rows.Count; i++)
            {
                // to do: format datetime values before printing
                for (int j = 0; j < Tbl.Columns.Count; j++)
                {
                    xlWorkSheet.Cells[(i + 2), (j + 1)] = Tbl.Rows[i][j];
                }
            }

            xlWorkSheetRange = xlWorkSheet.Rows.get_Range("1:1");
            xlWorkSheetRange.Select();

            xlWorkSheetRange.Font.Bold = true;
            xlWorkSheetRange.Font.Italic = true;

            string desktopPath = Environment.GetFolderPath(Environment.SpecialFolder.Desktop);
            string location = Path.Combine(desktopPath, Tbl.TableName + DateTime.Now.ToString("yyyyMMddHHmmssffff") + ".xls");
            xlWorkBook.SaveAs(location, Excel.XlFileFormat.xlWorkbookNormal, misValue, misValue, misValue, misValue, Excel.XlSaveAsAccessMode.xlExclusive, misValue, misValue, misValue, misValue, misValue);
            xlWorkBook.Close(true, misValue, misValue);
            xlSamp.Visible = true;
            //xlSamp.Quit();

            try
            {
                System.Runtime.InteropServices.Marshal.ReleaseComObject(xlSamp);
                xlSamp = null;
            }
            catch (Exception ex)
            {
                xlSamp = null;
                Console.Write("Error " + ex.ToString());
            }
            finally
            {
                GC.Collect();
            }
        }
    }
}
