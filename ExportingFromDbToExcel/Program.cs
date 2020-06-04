using System;
using System.Data;
using System.Configuration;
using System.Data.SqlClient;


namespace ExportingFromDbToExcel
{
    class Program
    {
        static void Main(string[] args)
        {
            //string connectionString = ConfigurationManager.ConnectionStrings["DefaultConnection"].ConnectionString;

            //using (SqlConnection con = new SqlConnection(connectionString))
            //{

            //    SqlCommand cmd = new SqlCommand("select * from ProjectDetails where ProjectModified >=   '2020-04-18 11:32:54.993' And ProjectModified <= '2020-05-01 17:06:49.057'", con);
            //    SqlDataAdapter sda = new SqlDataAdapter();
            //    sda.SelectCommand = cmd;
            //    DataTable dt = new DataTable();

            //    sda.Fill(dt);
            //    object misValue = System.Reflection.Missing.Value;
            //    Microsoft.Office.Interop.Excel.Application app = new Microsoft.Office.Interop.Excel.Application();
            //    app.Visible = false;

            //    Microsoft.Office.Interop.Excel.Workbook wb = app.Workbooks.Add(Microsoft.Office.Interop.Excel.XlWBATemplate.xlWBATWorksheet);
            //    Microsoft.Office.Interop.Excel.Worksheet ws = (Microsoft.Office.Interop.Excel.Worksheet)wb.ActiveSheet;

            //    for (int i = 0; i < dt.Columns.Count; i++)
            //    {
            //        ws.Cells[1, i + 1] = dt.Columns[i].ColumnName;
            //    }

            //    for (int i = 0; i < dt.Rows.Count; i++)
            //    {
            //        for (int j = 0; j < dt.Columns.Count; j++)
            //        {
            //            ws.Cells[i + 2, j + 1] = dt.Rows[i][j].ToString();
            //        }
            //    }
            //    try
            //    {
            //        ws.Name = dt.TableName;
            //    }
            //    catch (Exception ex)
            //    {
            //        //  lbl.Text = ex.ToString();
            //    }
            //    try
            //    {
            //        wb.SaveAs("Ope.xls", Microsoft.Office.Interop.Excel.XlFileFormat.xlWorkbookNormal, misValue, misValue, misValue, misValue, Microsoft.Office.Interop.Excel.XlSaveAsAccessMode.xlExclusive, misValue, misValue, misValue, misValue, misValue);
            //        Console.WriteLine("Saved");
            //        Console.ReadLine();
            //    }
            //    catch(Exception ex)
            //    {
            //        Console.WriteLine("failed | " + ex.Message);
            //        Console.ReadLine();
            //    }
            //    wb.Close(true, misValue, misValue);
            //    app.Quit();

            //}

            Console.WriteLine(ExcelManager.ExcelSend());
            Console.ReadLine(); 
        }
    }
}