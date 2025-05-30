﻿using Microsoft.Data.SqlClient;
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
using Microsoft.IdentityModel.Tokens;
using System.Reflection;
using System.Globalization;
using ExcelSQLImporter.Services;

namespace ExcelSQLImporter
{
    class Program
    {
        static async Task<int> Main(string[] args)
        {
            bool? logToFile = true;
            bool? outputToScreen = true;

            string? toolName = Assembly.GetExecutingAssembly().GetName().Name;
            string logFileName = $"{toolName} - {DateTime.Now.ToString("yyyy-MM-dd HH-mm-ss")}.log";

            await LoggingService.Log(toolName, logFileName, logToFile, outputToScreen);
            await LoggingService.Log("Import Excel File to SQL Table", logFileName, logToFile, outputToScreen);
            await LoggingService.Log("=========================================", logFileName, logToFile, outputToScreen);

            string? productVersion = Assembly.GetExecutingAssembly().GetName().Version?.ToString();
            await LoggingService.Log($"\nVersion {productVersion}", logFileName, logToFile, outputToScreen);
            await LoggingService.Log($"\nCopyright Robin Wilson", logFileName, logToFile, outputToScreen);

            string configFile = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "appsettings.json");
            string? customConfigFile = null;
            if (args.Length >= 1)
            {
                customConfigFile = args[0];
            }

            if (!string.IsNullOrEmpty(customConfigFile))
            {
                configFile = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, customConfigFile);
            }

            await LoggingService.Log($"\nUsing Config File {configFile}", logFileName, logToFile, outputToScreen);

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
                await LoggingService.Log($"Error: {e}", logFileName, logToFile, outputToScreen);
                return 1;
            }

            await LoggingService.Log($"\nSetting Locale To {config["Locale"]}", logFileName, logToFile, outputToScreen);

            //Set locale to ensure dates and currency are correct
            CultureInfo culture = new CultureInfo(config["Locale"] ?? "en-GB");
            Thread.CurrentThread.CurrentCulture = culture;
            Thread.CurrentThread.CurrentUICulture = culture;
            CultureInfo.DefaultThreadCurrentCulture = culture;
            CultureInfo.DefaultThreadCurrentUICulture = culture;

            var databaseConnection = config.GetSection("DatabaseConnection");
            var databaseTable = config.GetSection("DatabaseTable");
            string? schemaName = databaseTable["Schema"] ?? "dbo";
            var excelFile = config.GetSection("ExcelFile");
            var ftpConnection = config.GetSection("FTPConnection");
            var storedProcedure = config.GetSection("StoredProcedure");
            string[]? filePaths = { @excelFile["Folder"] ?? "", excelFile["FileName"] ?? "" };
            string excelFilePath = Path.Combine(filePaths);
            string? excelFileNameNoExtension = excelFile["FileName"]?.Substring(0, excelFile["FileName"]!.LastIndexOf("."));

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
            //        await LoggingService.Log(string.Format("Row {0} = {1}", row, sheet.GetRow(row).GetCell(1).StringCellValue));
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

                switch(ftpConnection?["Type"])
                {
                    case "FTP":
                        sessionOptions.Protocol = Protocol.Ftp;
                        break;
                    case "FTPS":
                        sessionOptions.Protocol = Protocol.Ftp;
                        sessionOptions.FtpSecure = FtpSecure.Explicit;
                        sessionOptions.GiveUpSecurityAndAcceptAnyTlsHostCertificate = true;
                        break;
                    case "SFTP":
                        sessionOptions.Protocol = Protocol.Sftp;
                        sessionOptions.GiveUpSecurityAndAcceptAnyTlsHostCertificate = true;
                        break;
                    case "SCP":
                        sessionOptions.Protocol = Protocol.Scp;
                        sessionOptions.GiveUpSecurityAndAcceptAnyTlsHostCertificate = true;
                        break;
                    default:
                        sessionOptions.Protocol = Protocol.Ftp;
                        break;
                }

                if (ftpConnection?["SSHHostKeyFingerprint"]?.Length > 0)
                {
                    sessionOptions.SshHostKeyFingerprint = ftpConnection["SSHHostKeyFingerprint"];
                    sessionOptions.GiveUpSecurityAndAcceptAnyTlsHostCertificate = false;
                }

                switch (ftpConnection?["Mode"])
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

                await LoggingService.Log("\nDownloading Excel File", logFileName, logToFile, outputToScreen);
                await LoggingService.Log($"Downloding File {excelFile["FileName"]} From {sessionOptions.HostName}", logFileName, logToFile, outputToScreen);

                try
                {
                    using (Session session = new Session())
                    {
                        //When publishing to a self-contained exe file need to specify the location of WinSCP.exe
                        session.ExecutablePath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "WinSCP.exe");

                        // Connect
                        session.Open(sessionOptions);

                        // Download files
                        string downloadpath = Path.Combine("/", ftpConnection?["FolderPath"] ?? "");

                        if (downloadpath.Substring(downloadpath.Length - 1) != "/")
                        {
                            downloadpath = downloadpath + "/";
                        }

                        TransferOptions transferOptions = new TransferOptions();
                        transferOptions.TransferMode = TransferMode.Binary;

                        TransferOperationResult transferResult;
                        transferResult =
                            session.GetFiles(downloadpath + excelFile["FileName"], @excelFilePath, false, transferOptions);

                        // Throw on any error
                        transferResult.Check();

                        // Print results
                        foreach (TransferEventArgs transfer in transferResult.Transfers)
                        {
                            await LoggingService.Log($"Download of {transfer.FileName} succeeded from {downloadpath}", logFileName, logToFile, outputToScreen);
                        }
                    }
                }
                catch (Exception e)
                {
                    await LoggingService.Log($"Error: {e}", logFileName, logToFile, outputToScreen);
                    return 1;
                }
            }
            else
            {
                await LoggingService.Log($"Not Downloading File to FTP as Option in Config is False", logFileName, logToFile, outputToScreen);
            }

            //Load Excel File
            await LoggingService.Log("\nLoading Data from Excel", logFileName, logToFile, outputToScreen);
            await LoggingService.Log($"Loading Excel File from {excelFilePath}", logFileName, logToFile, outputToScreen);

            IWorkbook book;
            DataTable table;
            if (System.IO.File.Exists(excelFilePath))
            {
                using (FileStream file = new FileStream(@excelFilePath, FileMode.Open, FileAccess.Read))
                {
                    book = WorkbookFactory.Create(file);
                }

                ISheet sheet;

                //Get first sheet in Excel file if not name specified
                if (!String.IsNullOrEmpty(excelFile["SheetName"])) {
                    sheet = book.GetSheet(excelFile["SheetName"]);
                }
                else {
                    sheet = book.GetSheetAt(0);
                }

                string sheetName = sheet.SheetName;
                string? tableNameOverride = null;

                if (databaseTable?["TableNameOverride"]?.Length > 0)
                {
                    tableNameOverride = databaseTable?["TableNameOverride"];
                }

                switch(databaseTable?["TableNamingMethod"])
                {
                    case "SheetName":
                        table = new DataTable(databaseTable?["TablePrefix"] + (tableNameOverride ?? sheetName));
                        break;
                    case "FileName":
                        table = new DataTable(databaseTable?["TablePrefix"] + (tableNameOverride ?? excelFileNameNoExtension));
                        break;
                    default:
                        table = new DataTable(databaseTable?["TablePrefix"] + (tableNameOverride ?? sheetName));
                        break;
                }

                table.Rows.Clear();
                table.Columns.Clear();

                for (int rowIndex = 0; rowIndex <= sheet.LastRowNum; rowIndex++)
                {
                    DataRow? tableRow = null;
                    IRow? row = sheet.GetRow(rowIndex);

                    if (row != null) //Check row does not only contains empty cells
                    {
                        //Don't want to insert header line as a row
                        if (rowIndex > 0)
                        {
                            tableRow = table.NewRow();
                        }

                        int colIndex = 0;

                        //Process rows and place into data table with named columns, correct data types
                        foreach (ICell cell in row.Cells)
                        {
                            object? cellValue = null;
                            string? fieldTypeToUse = "";
                            List<string> fieldTypes = new List<string>();

                            //If first row then add named columns and don't add a row
                            if (rowIndex == 0)
                            {
                                //Check all cell types for this column excluding header row to pick best type
                                for (int i = 0; i < sheet.LastRowNum; i++)
                                {
                                    IRow? rowToCheck = sheet.GetRow(i + 1);
                                    ICell? cellInRow = rowToCheck?.GetCell(cell.ColumnIndex);
                                    string fieldType = "System.String";

                                    if (cellInRow != null)
                                    {
                                        switch (cellInRow.CellType)
                                        {
                                            case CellType.Blank: cellValue = DBNull.Value; break;
                                            case CellType.Boolean: cellValue = cellInRow.BooleanCellValue; fieldType = "System.Boolean"; break;
                                            case CellType.String: cellValue = cellInRow.StringCellValue; fieldType = "System.String"; break;
                                            case CellType.Numeric:
                                                if (HSSFDateUtil.IsCellDateFormatted(cellInRow)) { cellValue = cellInRow.DateCellValue; fieldType = "System.DateTime"; }
                                                else { cellValue = cellInRow.NumericCellValue; fieldType = "System.Double"; }
                                                break;

                                            case CellType.Formula:
                                                bool cont = true;
                                                switch (cellInRow.CachedFormulaResultType)
                                                {
                                                    case CellType.Blank: cellValue = DBNull.Value; break;
                                                    case CellType.String: cellValue = cellInRow.StringCellValue; fieldType = "System.String"; break;
                                                    case CellType.Boolean: cellValue = cellInRow.BooleanCellValue; fieldType = "System.Boolean"; break;
                                                    case CellType.Numeric:
                                                        if (HSSFDateUtil.IsCellDateFormatted(cellInRow)) { cellValue = cellInRow.DateCellValue; fieldType = "System.DateTime"; }
                                                        else
                                                        {
                                                            try
                                                            {
                                                                //Check if Boolean
                                                                if (cellInRow.CellFormula == "TRUE()") { cellValue = cellInRow.BooleanCellValue; fieldType = "System.Boolean"; cont = false; }
                                                                if (cont && cellInRow.CellFormula == "FALSE()") { cellValue = cellInRow.BooleanCellValue; fieldType = "System.Boolean"; cont = false; }
                                                                if (cont) { cellValue = cellInRow.NumericCellValue; fieldType = "System.Double"; cont = false; }
                                                            }
                                                            catch { }
                                                        }
                                                        break;
                                                }
                                                break;
                                            default:
                                                fieldType = "System.String"; break;
                                        }

                                        if (cellValue?.ToString()?.Length > 0)
                                        {
                                            fieldTypes.Add(fieldType);

                                            //Output types to check for specific column
                                            //if (colIndex == 35)
                                            //{
                                            //    await LoggingService.Log($"Value '{cellValue}' is {cellInRow.CellType}");
                                            //}
                                        }
                                    }
                                }

                                //Pick best type of field
                                if (fieldTypes.Contains("System.String"))
                                {
                                    //If any of the rows contain a string then need to set row to that to be able to store all values
                                    fieldTypeToUse = "System.String";
                                }
                                else if (fieldTypes.Contains("System.Double") && fieldTypes.Contains("System.DateTime"))
                                {
                                    //If rows are mixed types then store as string
                                    fieldTypeToUse = "System.String";
                                }
                                else if (fieldTypes.Contains("System.Int32") && fieldTypes.Contains("System.DateTime"))
                                {
                                    //If rows are mixed types then store as string
                                    fieldTypeToUse = "System.String";
                                }
                                else if (fieldTypes.Contains("System.Boolean") && fieldTypes.Contains("System.DateTime"))
                                {
                                    //If rows are mixed types then store as string
                                    fieldTypeToUse = "System.String";
                                }
                                else if (fieldTypes.Contains("System.Double"))
                                {
                                    fieldTypeToUse = "System.Double";
                                }
                                else if (fieldTypes.Contains("System.Int32"))
                                {
                                    fieldTypeToUse = "System.Int32";
                                }
                                else if (fieldTypes.Contains("System.Boolean"))
                                {
                                    fieldTypeToUse = "System.Boolean";
                                }
                                else if (fieldTypes.Contains("System.DateTime"))
                                {
                                    fieldTypeToUse = "System.DateTime";
                                }
                                else
                                {
                                    fieldTypeToUse = "System.String";
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
                                DataColumn tableColumn = new DataColumn(colName, Type.GetType(fieldTypeToUse) ?? typeof(string));
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
                                        bool cont = true;
                                        switch (cell.CachedFormulaResultType)
                                        {
                                            case CellType.Blank: cellValue = DBNull.Value; break;
                                            case CellType.String: cellValue = cell.StringCellValue; break;
                                            case CellType.Boolean: cellValue = cell.BooleanCellValue; break;
                                            case CellType.Numeric:
                                                if (HSSFDateUtil.IsCellDateFormatted(cell)) { cellValue = cell.DateCellValue; }
                                                else
                                                {
                                                    try
                                                    {
                                                        //Check if Boolean
                                                        if (cell.CellFormula == "TRUE()") { cellValue = cell.BooleanCellValue; cont = false; }
                                                        if (cont && cell.CellFormula == "FALSE()") { cellValue = cell.BooleanCellValue; cont = false; }
                                                        if (cont) { cellValue = cell.NumericCellValue; cont = false; }
                                                    }
                                                    catch { }
                                                }
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
                await LoggingService.Log($"The File at {excelFilePath} Could Not Be Found", logFileName, logToFile, outputToScreen);
                return 1;
            }

            await LoggingService.Log($"Loaded {table?.Rows.Count} rows of data from file", logFileName, logToFile, outputToScreen);

            //Save to Database
            await LoggingService.Log("\nSaving Data To Database", logFileName, logToFile, outputToScreen);
            await LoggingService.Log($"Creating Table {table?.TableName} in Database", logFileName, logToFile, outputToScreen);
            await using var connection = new SqlConnection(connectionString);

            try
            {
                await connection.OpenAsync();

                if (table != null)
                {
                    string createTableSQL = CreateTableSQL(schemaName ?? "dbo", table?.TableName ?? "Imported_Excel_File", table!);
                    //await LoggingService.Log($"{createTableSQL}");

                    using (SqlCommand command = new SqlCommand(createTableSQL, connection))
                        await command.ExecuteNonQueryAsync();
                }
            }
            catch (Exception e)
            {
                await LoggingService.Log(e.ToString(), logFileName, logToFile, outputToScreen);

                if (connection != null)
                {
                    await connection.CloseAsync();
                }

                return 1;
            }

            await LoggingService.Log($"Uploading {table?.Rows.Count} Rows of Data into Table {table?.TableName} in Database", logFileName, logToFile, outputToScreen);

            try
            {
                SqlBulkCopy bulkcopy = new SqlBulkCopy(connection);
                bulkcopy.DestinationTableName = table?.TableName;

                await bulkcopy.WriteToServerAsync(table);
            }
            catch (Exception e)
            {
                await LoggingService.Log(e.ToString(), logFileName, logToFile, outputToScreen);

                if (connection != null)
                {
                    await connection.CloseAsync();
                }

                return 1;
            }

            //Run Stored Procedure On Completion
            if (storedProcedure.GetValue<bool?>("RunTask", false) == true)
            {
                await LoggingService.Log("\nRunning Post Import Stored Procedure", logFileName, logToFile, outputToScreen);
                await LoggingService.Log($"Running Stored Procedure: {storedProcedure["Database"]}.{storedProcedure["Schema"]}.{storedProcedure["StoredProcedure"]}", logFileName, logToFile, outputToScreen);

                if (storedProcedure["StoredProcedure"]?.Length > 0)
                {
                    try
                    {
                        if (table != null)
                        {
                            string customTaskSQL = $"EXEC {storedProcedure["Database"]}.{storedProcedure["Schema"]}.{storedProcedure["StoredProcedure"]}";
                            //await LoggingService.Log($"{createTableSQL}");

                            using (SqlCommand command = new SqlCommand(customTaskSQL, connection))
                                await command.ExecuteNonQueryAsync();

                            await LoggingService.Log($"Stored Procedure Completed", logFileName, logToFile, outputToScreen);
                        }
                    }
                    catch (Exception e)
                    {
                        await LoggingService.Log(e.ToString(), logFileName, logToFile, outputToScreen);

                        if (connection != null)
                        {
                            await connection.CloseAsync();
                        }

                        return 1;
                    }
                }
                else
                {
                    await LoggingService.Log($"Cannot run stored procedure as it has not been specified in the config file", logFileName, logToFile, outputToScreen);
                }
            }

            //Close database connection
            if (connection != null)
            {
                await connection.CloseAsync();
            }

            return 0;
        }

        public static string CreateTableSQL(string schemaName, string tableName, DataTable table)
        {
            string sqlsc;
            sqlsc = $"\n DROP TABLE IF EXISTS [{schemaName ?? "dbo"}].[{tableName}];";
            sqlsc += $"\n CREATE TABLE [{schemaName ?? "dbo"}].[{tableName}] (";

            //Check Cell Value
            //await LoggingService.Log(table.Rows[1][4].ToString());

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
                        for (int rowIndex = 0; rowIndex < table.Rows.Count; rowIndex++)
                        {
                            rowLength = (table.Rows[rowIndex][i].ToString() ?? "").Length;
                            //await LoggingService.Log($"Column: {i}, Row: {rowIndex}, Length: {rowLength}");

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