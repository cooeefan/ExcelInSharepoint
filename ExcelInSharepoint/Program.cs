using System;
using System.Collections.Generic;
using System.Text;
using System.Threading.Tasks;
using Microsoft.Office.Interop.Excel;
using System.Net;
using System.Data.SqlClient;
using System.Configuration;

namespace ExcelInSharepoint
{
    class Program
    {
        static void Main(string[] args)
        {
            string excelPath = ConfigurationManager.AppSettings["ExcelPath"];
            string targetTable = ConfigurationManager.AppSettings["TargetTable"];
            string connnectionString = ConfigurationManager.AppSettings["DBConnectionString"];

            string url = WebUtility.HtmlEncode(excelPath);
            Console.WriteLine(url);

            LogWriter.Write("Programme started...");

            Application excel = null;
            Workbooks workbooks = null;
            Workbook excelWorkbook = null;

            try
            {
                excel = new Application();
                excel.Visible = false;
                workbooks = excel.Workbooks;

                LogWriter.Write("Opening excel...");

                excelWorkbook = workbooks.Open(url, 0, false, 5, "", "", false, XlPlatform.xlWindows, "", true, false, 0, true, false, false);
                Sheets sheets = excelWorkbook.Worksheets;
                Worksheet worksheet = (Worksheet)sheets.get_Item(4);

                LogWriter.Write("Excel opened");

                Range xlrange = worksheet.UsedRange;

                int rowCount = xlrange.Rows.Count;
                int colCount = xlrange.Columns.Count;

                //Put cells into a array to increase the performance
                object[,] xlArray = xlrange.get_Value(Type.Missing);

                //Get last updated date
                string lastUpdateDate = "";
                if (xlArray[1, 4] != null)
                    lastUpdateDate = xlArray[1, 4].ToString();

                //Build insert SQL
                string insertSQLPrefix = "INSERT INTO " + targetTable + " ([";

                for (int a = 1; a <= colCount; a++)
                {
                    if (xlArray[2, a] != null)
                        insertSQLPrefix = insertSQLPrefix + xlArray[2, a].ToString() + "],[";
                }
                insertSQLPrefix = insertSQLPrefix + "LastUpdateDate]) VALUES( '";
                //Console.WriteLine(insertSQLPrefix);

                using (SqlConnection sqlconn = new SqlConnection(connnectionString))
                {
                    using (SqlCommand sqlcomm = new SqlCommand())
                    {
                        sqlconn.Open();
                        sqlcomm.Connection = sqlconn;

                        if(rowCount>0 && colCount>0)
                        {
                            string truncateSQL = "truncate table " + targetTable;
                            sqlcomm.CommandText = truncateSQL;
                            sqlcomm.ExecuteNonQuery();
                        }

                        for (int i = 3; i <= rowCount; i++)
                        {
                            if (xlArray[i, 1] != null)
                            {
                                StringBuilder insertSQL = new StringBuilder(insertSQLPrefix);
                                for (int j = 1; j <= colCount; j++)
                                {
                                    if (xlArray[i, j] != null)
                                    {
                                        //Console.WriteLine(xlrange.Cells[i, j].Value2.ToString());
                                        insertSQL.Append(xlArray[i, j].ToString().Replace("'", "''"));
                                        insertSQL.Append("','");
                                    }
                                }
                                insertSQL.Append(lastUpdateDate + "')");
                                LogWriter.Write(insertSQL.ToString());
                                sqlcomm.CommandText = insertSQL.ToString();
                                sqlcomm.ExecuteNonQuery();
                            }

                        }
                    }
                }

                LogWriter.Write("Job finished");
            }

            catch(Exception ex)
            {
                LogWriter.Write(ex.ToString());
            }

            finally
            {
                //Clean job
                excelWorkbook.Close();
                excel.Quit();

                GC.Collect();
                GC.WaitForPendingFinalizers();

                System.Runtime.InteropServices.Marshal.FinalReleaseComObject(excelWorkbook);
                System.Runtime.InteropServices.Marshal.FinalReleaseComObject(excel);
            }


        }
    }
}
