using Microsoft.Data.SqlClient;
using System;
using System.Threading.Tasks;
using NPOI.XSSF.UserModel;
using Microsoft.VisualBasic.FileIO;
using System.ComponentModel.DataAnnotations;
using System.Data;
using NPOI.OpenXmlFormats.Wordprocessing;
using NPOI.SS.UserModel;
using System.IO;
using MathNet.Numerics.Optimization;
using Microsoft.Extensions.Configuration;
using WinSCP;
using NPOI.HSSF.UserModel;
using NPOI.SS.Formula.Functions;
using Microsoft.EntityFrameworkCore.Metadata.Internal;
using static Microsoft.EntityFrameworkCore.DbLoggerCategory.Database;
using NPOI.OpenXmlFormats.Spreadsheet;
using NPOI.HPSF;
using Org.BouncyCastle.Bcpg;
using NPOI.Util;

namespace ExcelSQLImporter
{
    class Program
    {
        static async Task<int> Main(string[] args)
        {
            Console.WriteLine("\nImport Excel File to SQL Table");
            Console.WriteLine("=========================================\n");
            Console.WriteLine("Copyright Robin Wilson");

            string configFile = "appsettings.json";
            string? customConfigFile = null;
            if (args.Length >= 1)
            {
                customConfigFile = args[0];
            }

            if (!string.IsNullOrEmpty(customConfigFile))
            {
                configFile = customConfigFile;
            }

            Console.WriteLine($"\nUsing Config File {configFile}");

            var builder = new ConfigurationBuilder()
                .SetBasePath(Directory.GetCurrentDirectory())
                .AddJsonFile(configFile, optional: false);

            IConfiguration config;
            try
            {
                config = builder.Build();
            }
            catch (Exception e)
            {
                Console.WriteLine("Error: {0}", e);
                return 1;
            }
            
            var databaseConnection = config.GetSection("DatabaseConnection");
            var excelFile = config.GetSection("ExcelFile");
            var ftpConnection = config.GetSection("FTPConnection");
            string excelFilePath = excelFile["Folder"] + "\\" + excelFile["FileName"];

            var sqlConnection = new SqlConnectionStringBuilder
            {
                DataSource = databaseConnection["Server"],
                UserID = databaseConnection["Username"],
                Password = databaseConnection["Password"],
                IntegratedSecurity = databaseConnection.GetValue<bool>("UseWindowsAuth", false),
                InitialCatalog = databaseConnection["Database"],
                TrustServerCertificate = true
            };

            //If not using windows auth then need username and password values too
            if (sqlConnection.IntegratedSecurity == false) {
                sqlConnection.UserID = databaseConnection["Username"];
                sqlConnection.Password = databaseConnection["Password"];
            }

            var connectionString = sqlConnection.ConnectionString;

            //Output Values into Command Window
            //for (int row = 0; row <= sheet.LastRowNum; row++)
            //{
            //    if (sheet.GetRow(row) != null) //null is when the row only contains empty cells 
            //    {
            //        Console.WriteLine(string.Format("Row {0} = {1}", row, sheet.GetRow(row).GetCell(1).StringCellValue));
            //    }
            //}

            //Get Excel File

            if (ftpConnection.GetValue<bool?>("DownloadFile", false) == true)
            {
                // Setup session options
                SessionOptions sessionOptions = new SessionOptions
                {
                    HostName = ftpConnection["Server"],
                    PortNumber = ftpConnection.GetValue<int>("Port", 21),
                    UserName = ftpConnection["Username"],
                    Password = ftpConnection["Password"]
                };

                switch(ftpConnection["Type"])
                {
                    case "FTP":
                        sessionOptions.Protocol = Protocol.Ftp;
                        break;
                    case "FTPS":
                        sessionOptions.Protocol = Protocol.Ftp;
                        sessionOptions.FtpSecure = FtpSecure.Explicit;
                        break;
                    case "SFTP":
                        sessionOptions.Protocol = Protocol.Sftp;
                        break;
                    default:
                        sessionOptions.Protocol = Protocol.Ftp;
                        break;
                }

                switch (ftpConnection["Mode"])
                {
                    case "Active":
                        sessionOptions.FtpMode = FtpMode.Active;
                        break;
                    case "Passive":
                        sessionOptions.FtpMode = FtpMode.Passive;
                        break;
                    default:
                        sessionOptions.FtpMode = FtpMode.Passive;
                        break;
                }

                Console.WriteLine($"Downloding File {excelFile["FileName"]} From {sessionOptions.HostName}");

                try
                {
                    using (Session session = new Session())
                    {
                        // Connect
                        session.Open(sessionOptions);

                        // Upload files
                        TransferOptions transferOptions = new TransferOptions();
                        transferOptions.TransferMode = TransferMode.Binary;

                        TransferOperationResult transferResult;
                        transferResult =
                            session.GetFiles("/" + excelFile["FileName"], @excelFilePath, false, transferOptions);

                        // Throw on any error
                        transferResult.Check();

                        // Print results
                        foreach (TransferEventArgs transfer in transferResult.Transfers)
                        {
                            Console.WriteLine("Download of {0} succeeded", transfer.FileName);
                        }
                    }
                }
                catch (Exception e)
                {
                    Console.WriteLine("Error: {0}", e);
                    return 1;
                }
            }
            else
            {
                Console.WriteLine($"Not Uploading File to FTP as Option in Config is False");
            }

            //Load Excel File
            Console.WriteLine($"Loading Excel File from {excelFilePath}");

            IWorkbook book;
            DataTable table;
            if (System.IO.File.Exists(excelFilePath))
            {
                using (FileStream file = new FileStream(@excelFilePath, FileMode.Open, FileAccess.Read))
                {
                    book = WorkbookFactory.Create(file);
                }

                //ISheet sheet = book.GetSheet("Sheet1");
                ISheet sheet = book.GetSheetAt(0);
                string sheetName = sheet.SheetName;
                table = new DataTable(sheetName);
                table.Rows.Clear();
                table.Columns.Clear();

                for (int rowIndex = 0; rowIndex <= sheet.LastRowNum; rowIndex++)
                {
                    DataRow? tableRow = null;
                    IRow? row = sheet.GetRow(rowIndex);
                    IRow? row2 = null;
                    IRow? row3 = null;

                    if (rowIndex == 0)
                    {
                        //If first row then need second row to check data types (check row 3 too)
                        row2 = sheet.GetRow(rowIndex + 1);
                        row3 = sheet.GetRow(rowIndex + 2);
                    }

                    if (row != null) //Check row does not only contains empty cells
                    {
                        //Don't want to insert header line as a row
                        if (rowIndex > 0)
                        {
                            tableRow = table.NewRow();
                        }

                        int colIndex = 0;

                        //For each cell in the row
                        foreach (ICell cell in row.Cells)
                        {
                            object? cellValue = null;
                            string? cellType = "";
                            string[] cellType2 = new string[2];

                            if (rowIndex == 0)
                            {
                                //Row 0 Should Contain Row Titles in Excel
                                for (int i = 0; i < 2; i++)
                                {
                                    ICell? cell2 = null;
                                    if (i == 0)
                                    {
                                        cell2 = row2?.GetCell(cell.ColumnIndex);
                                    }
                                    else
                                    {
                                        cell2 = row3?.GetCell(cell.ColumnIndex);
                                    }

                                    if (cell2 != null)
                                    {
                                        switch (cell2.CellType)
                                        {
                                            case CellType.Blank: break;
                                            case CellType.Boolean: cellType2[i] = "System.Boolean"; break;
                                            case CellType.String: cellType2[i] = "System.String"; break;
                                            case CellType.Numeric:
                                                if (HSSFDateUtil.IsCellDateFormatted(cell2))
                                                {
                                                    cellType2[i] = "System.DateTime";
                                                }
                                                else
                                                {
                                                    cellType2[i] = "System.Double";
                                                }
                                                break;

                                            case CellType.Formula:
                                                bool cont = true;
                                                switch (cell2.CachedFormulaResultType)
                                                {
                                                    case CellType.Boolean: cellType2[i] = "System.Boolean"; break;
                                                    case CellType.String: cellType2[i] = "System.String"; break;
                                                    case CellType.Numeric:
                                                        if (HSSFDateUtil.IsCellDateFormatted(cell2)) { cellType2[i] = "System.DateTime"; }
                                                        else
                                                        {
                                                            try
                                                            {
                                                                //Check if Boolean
                                                                if (cell2.CellFormula == "TRUE()") { cellType2[i] = "System.Boolean"; cont = false; }
                                                                if (cont && cell2.CellFormula == "FALSE()") { cellType2[i] = "System.Boolean"; cont = false; }
                                                                if (cont) { cellType2[i] = "System.Double"; cont = false; }
                                                            }
                                                            catch { }
                                                        }
                                                        break;
                                                }
                                                break;
                                            default:
                                                cellType2[i] = "System.String"; break;
                                        }
                                    }
                                }

                                //Resolve different types
                                if (cellType2[0] == cellType2[1]) { cellType = cellType2[0]; }
                                else
                                {
                                    if (cellType2[0] == null) cellType = cellType2[1];
                                    if (cellType2[1] == null) cellType = cellType2[0];
                                    if (cellType == "") cellType = "System.String";
                                }

                                //Get the name of the column
                                string colName = "Column_{0}";
                                try { colName = cell.StringCellValue; }
                                catch { colName = string.Format(colName, colIndex); }

                                //Check the name of the column is not repeated
                                foreach (DataColumn col in table.Columns)
                                {
                                    if (col.ColumnName == colName) colName = string.Format("{0}_{1}", colName, colIndex);
                                }

                                //Add field to the table
                                DataColumn tableColumn = new DataColumn(colName, Type.GetType(cellType) ?? typeof(string));
                                table.Columns.Add(tableColumn); colIndex++;
                            }
                            else
                            {
                                //All Other Rows Aside from Header
                                switch (cell.CellType)
                                {
                                    case CellType.Blank: cellValue = DBNull.Value; break;
                                    case CellType.Boolean: cellValue = cell.BooleanCellValue; break;
                                    case CellType.String: cellValue = cell.StringCellValue; break;
                                    case CellType.Numeric:
                                        if (HSSFDateUtil.IsCellDateFormatted(cell)) { cellValue = cell.DateCellValue; }
                                        else { cellValue = cell.NumericCellValue; }
                                        break;
                                    case CellType.Formula:
                                        switch (cell.CachedFormulaResultType)
                                        {
                                            case CellType.Blank: cellValue = DBNull.Value; break;
                                            case CellType.String: cellValue = cell.StringCellValue; break;
                                            case CellType.Boolean: cellValue = cell.BooleanCellValue; break;
                                            case CellType.Numeric:
                                                if (HSSFDateUtil.IsCellDateFormatted(cell)) { cellValue = cell.DateCellValue; }
                                                else { cellValue = cell.NumericCellValue; }
                                                break;
                                        }
                                        break;
                                    default: cellValue = cell.StringCellValue; break;
                                }
                                //If the cell has a blank value then make it null in the SQL Table
                                if (cellValue?.ToString()?.Length == 0)
                                {
                                    cellValue = DBNull.Value;
                                }
                                //Add the cell to the row
                                if (tableRow != null)
                                {
                                    if (cell.ColumnIndex <= table.Columns.Count - 1) tableRow[cell.ColumnIndex] = cellValue;
                                }
                            }
                        }

                        //Add the row to the table
                        if (tableRow != null)
                        {
                            if (rowIndex > 0) table.Rows.Add(tableRow);
                        }
                    }
                    table.AcceptChanges();
                }
            }
            else
            {
                Console.WriteLine($"The File at {excelFilePath} Could Not Be Found");
                return 1;
            }


            //Save to Database
            Console.WriteLine($"Creating Table {table.TableName} in Database");
            await using var connection = new SqlConnection(connectionString);
            try
            {
                await connection.OpenAsync();

                string createTableSQL = CreateTableSQL(table.TableName, table);
                //Console.WriteLine($"{createTableSQL}");

                using (SqlCommand command = new SqlCommand(createTableSQL, connection))
                    command.ExecuteNonQuery();
            }
            catch (Exception e)
            {
                Console.WriteLine(e.ToString());
                return 1;
            }

            Console.WriteLine($"Uploading {table.Rows.Count} Rows of Data into Table {table.TableName} in Database");

            try
            {
                SqlBulkCopy bulkcopy = new SqlBulkCopy(connection);
                bulkcopy.DestinationTableName = table.TableName;

                bulkcopy.WriteToServer(table);
                connection.Close();
            }
            catch (Exception e)
            {
                Console.WriteLine(e.ToString());
                return 1;
            }

            return 0;
        }

        public static string CreateTableSQL(string tableName, DataTable table)
        {
            string sqlsc;
            sqlsc = "\n DROP TABLE IF EXISTS [" + tableName + "];";
            sqlsc += "\n CREATE TABLE [" + tableName + "] (";

            //Check Cell Value
            //Console.WriteLine(table.Rows[1][4].ToString());

            for (int i = 0; i < table.Columns.Count; i++)
            {
                sqlsc += "\n [" + table.Columns[i].ColumnName + "]";
                System.Type columnType = table.Columns[i].DataType;

                if (columnType == typeof(Int32))
                {
                    sqlsc += " INT";
                }
                else if (columnType == typeof(Int64))
                {
                    sqlsc += " BIGINT";
                }
                else if (columnType == typeof(Int16))
                {
                    sqlsc += " SMALLINT";
                }
                else if (columnType == typeof(Byte))
                {
                    sqlsc += " TINYINT";
                }
                else if (columnType == typeof(System.Decimal))
                {
                    sqlsc += " DECIMAL";
                }
                else if (columnType == typeof(Double))
                {
                    sqlsc += " FLOAT";
                }
                else if (columnType == typeof(DateTime))
                {
                    sqlsc += " DATETIME";
                }
                else if (columnType == typeof(string))
                {
                    int rowLength = 0;
                    int maxRowLength = 0;
                    string maxRowLengthString = "";
                    if (table.Columns[i].MaxLength == -1)
                    {
                        for (int rowIndex = 1; rowIndex < table.Rows.Count; rowIndex++)
                        {
                            rowLength = (table.Rows[rowIndex][i].ToString() ?? "").Length;
                            if (rowLength > maxRowLength)
                            {
                                maxRowLength = rowLength;
                            }
                        }
                    }
                    else
                    {
                        maxRowLength = table.Columns[i].MaxLength;
                    }

                    if (maxRowLength > 0 && maxRowLength < 4000)
                    {
                        maxRowLengthString = maxRowLength.ToString();
                    }
                    else
                    {
                        maxRowLengthString = "MAX";
                    }
                    sqlsc += string.Format(" NVARCHAR({0})", maxRowLengthString);
                }
                else
                {
                    int rowLength = 0;
                    int maxRowLength = 0;
                    string maxRowLengthString = "";
                    if (table.Columns[i].MaxLength == -1)
                    {
                        for (int rowIndex = 1; rowIndex < table.Rows.Count; rowIndex++)
                        {
                            rowLength = (table.Rows[rowIndex][i].ToString() ?? "").Length;
                            if (rowLength > maxRowLength)
                            {
                                maxRowLength = rowLength;
                            }
                        }
                    }
                    else
                    {
                        maxRowLength = table.Columns[i].MaxLength;
                    }

                    if (maxRowLength > 0 && maxRowLength < 4000)
                    {
                        maxRowLengthString = maxRowLength.ToString();
                    }
                    else
                    {
                        maxRowLengthString = "MAX";
                    }
                    sqlsc += string.Format(" NVARCHAR({0})", maxRowLengthString);
                }
                
                if (table.Columns[i].AutoIncrement)
                    sqlsc += " IDENTITY(" + table.Columns[i].AutoIncrementSeed.ToString() + "," + table.Columns[i].AutoIncrementStep.ToString() + ")";
                if (!table.Columns[i].AllowDBNull)
                    sqlsc += " NOT NULL";
                sqlsc += ",";
            }
            return sqlsc.Substring(0, sqlsc.Length - 1) + "\n)";
        }
    }
}