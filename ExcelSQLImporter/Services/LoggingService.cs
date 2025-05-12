using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ExcelSQLImporter.Services
{
    public static class LoggingService
    {
        public static string? LogFilePath { get; set; }

        public static async Task<bool?> Log(string? message, string? logFileName, bool? logToFile, bool? outputToScreen)
        {
            if (string.IsNullOrEmpty(message))
            {
                return false;
            }

            string? toolName = System.Reflection.Assembly.GetExecutingAssembly().GetName().Name;
            string currentPath = Directory.GetCurrentDirectory();
            string? logFilePathAndName = Path.Combine(currentPath, logFileName ?? "Log.txt");

            try
            {
                if (!File.Exists(logFilePathAndName))
                {
                    File.Create(logFilePathAndName).Close();
                }

                StreamWriter logFileContents = new StreamWriter(logFilePathAndName, append: true);

                await logFileContents.WriteLineAsync(message);
                logFileContents.Close();
                logFileContents.Dispose();
            }
            catch (Exception ex)
            {
                Console.WriteLine($"\nError creating log file: {ex.Message}");
                return false;
            }

            if (outputToScreen == true)
            {
                Console.WriteLine(message);
            }

            return true;
        }
    }
}
