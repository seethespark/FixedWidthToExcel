using System;
using System.Collections.Generic;
using System.Data;
using System.Data.SqlClient;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Excel = Microsoft.Office.Interop.Excel;

namespace FixedWidthToExcel
{
    class Program
    {
        static void Main(string[] args)
        {
            string dbServerName = "MCMDSSQLDESQL01";
            string dbDatabaseName = "MacmillanDataStoreDev2";
            if (args.Length == 0)
            {
                System.Console.WriteLine("Please enter the path to the fixed width file or ? for help.");
                Console.ReadKey();
                return;
                
            }
            if (args.Length == 1 && (args[0] == "?" || args[0] == "??" || args[0] == "--help" || args[0] == "-h" )){
                Console.WriteLine("This application will open a fixed width file formatted in columns for Excel.  ");
                Console.WriteLine("Pass in the file name as the first parameter or drop it into the application");
                Console.WriteLine("Optional parameters are 2) server name, for column width and 3) database name for column width");
                Console.WriteLine("Defaults for parameters 2) and 3) are {0} and {1}", dbServerName, dbDatabaseName);
                Console.WriteLine("Under length rows are highlighted with the last cell blue.");
                Console.WriteLine("Over length rows are highlighted with a pink cell after the end of the row.");
                Console.WriteLine("Press any key to close.");
                Console.ReadKey();
                return;
            }
            if (args.Length == 2)
            {
                dbServerName = args[1];
            }
            if (args.Length == 3)
            {
                dbDatabaseName = args[2];
            }
            LoadExcel(args[0], dbServerName, dbDatabaseName);
            //LoadExcel(@"C:\Projects\TestImportFiles\Main files\Phrasis supporter import_3012_2015_04_02_150.txt", dbServerName, dbDatabaseName);

        }

        static void LoadExcel(string textFilePath, string dbServerName, string dbDatabaseName)
        {
            
            if (textFilePath == null | textFilePath == "" | !File.Exists(textFilePath))
            {
                Console.WriteLine("File foesn't exist: {0}", textFilePath);
                Console.ReadKey();
                return;
            }
            var sscsb = new SqlConnectionStringBuilder();
            sscsb.DataSource = dbServerName;
            sscsb.InitialCatalog = dbDatabaseName;
            sscsb.IntegratedSecurity = true;
            sscsb.ConnectTimeout = 10;
 

            //using (SqlConnection conn = new SqlConnection("Server=" + dbServerName + ";Database=" + dbDatabaseName + ";Integrated Security=SSPI;"))
            using (SqlConnection conn = new SqlConnection(sscsb.ConnectionString))
            using (SqlCommand cmd = new SqlCommand("SELECT ColumnWidth, FileColumnName FROM control.SupplierImport si JOIN control.SupplierImportColumn sic ON sic.SupplierImportID = si.SupplierImportID WHERE @fileName LIKE REPLACE(si.FileNameMask, '.', '%') AND si.FileType = 'Fixed width' ORDER BY sic.ColumnOrdinal ASC;  ", conn))

            {
                try
                {
                    //// DB bits
                    SqlParameter param = new SqlParameter("@fileName", System.Data.SqlDbType.VarChar, 50);
                    param.Value = textFilePath.Replace(@"\", @"/").Split(@"/".ToCharArray()[0])[textFilePath.Replace(@"\", @"/").Split(@"/".ToCharArray()[0]).Length - 1];
                    cmd.Parameters.Add(param);
                    DataTable dt = new DataTable();
                    SqlDataAdapter da = new SqlDataAdapter(cmd);
                    da.Fill(dt);
                    if (dt.Rows.Count == 0)
                    {
                        Console.WriteLine("No entry for '{0}' in the database.", param.Value);
                        Console.ReadKey();
                        return;
                    }

                    /// Excel bits
                    Excel.Application excel = new Excel.Application();
                    excel.Visible = true;
                    Excel.Workbook workbook = excel.Workbooks.Add(Excel.XlSheetType.xlWorksheet);
                    Excel.Worksheet sheet = workbook.Sheets[1];
                    sheet.Cells[1, 1].EntireRow.Font.Bold = true;

                    /// File bits

                    StreamReader file = new StreamReader(textFilePath);


                    string line;
                    string colText;
                    int colLen;
                    int rowNo = 1;
                    int colNo = 1;
                    int rowPosition;
                    int lineLen;

                    foreach (DataRow row in dt.Rows)
                    {
                        sheet.Cells[rowNo, colNo] = row["FileColumnName"];
                        colNo++;
                    }
                    rowNo++;
                    /// Skip the first row
                    file.ReadLine();
                    while ((line = file.ReadLine()) != null)
                    {
                        lineLen = line.Length;
                        colNo = 1;
                        rowPosition = 0;
                        foreach (DataRow row in dt.Rows)
                        {
                            colLen = int.Parse(row[0].ToString());
                            if ((rowPosition + colLen) > lineLen)
                            {
                                colText = line.Substring(rowPosition, lineLen - rowPosition);
                                sheet.Cells[rowNo, colNo] = colText;
                                sheet.Cells[rowNo, colNo].Interior.Color = Excel.XlRgbColor.rgbLightSkyBlue;
                                break;
                            }
                            else
                            {
                                colText = line.Substring(rowPosition, colLen);
                                sheet.Cells[rowNo, colNo] = colText;
                                rowPosition += colLen;
                                colNo++;
                            }

                        }
                        /// Handle over length rows
                        if (lineLen > rowPosition)
                            if (sheet.Cells[rowNo, colNo].Interior.ColorIndex < 0)
                                sheet.Cells[rowNo, colNo].Interior.Color = Excel.XlRgbColor.rgbMediumOrchid;

                        rowNo++;
                    }
                    file.Close();
                    file.Dispose();
                    Console.ReadKey();
                }
                catch (Exception ex)
                {
                    Console.WriteLine(ex.Message);
                    if (ex.InnerException!= null)
                        Console.WriteLine(ex.InnerException.Message);
                    Console.ReadKey();
                }
            }


        }
    }
}
