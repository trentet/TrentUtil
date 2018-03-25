using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Linq;
using System.Text;
using TableBuilderLibrary;
using System.Windows.Forms;

namespace TrentUtil
{
    public class FileIO
    {
        //This class library is currently missing exception handling

        /// <summary>
        /// Given an xlsx file path, it will import and return a DataTable
        /// </summary>
        /// <param name="filePath">Xlsx import path</param>
        /// <returns></returns>
        public static DataTable ImportDataTableFromExcel(string filePath)
        {
            DataSet d = Excel.ReadXlsx(filePath);
            d.Clear();

            DataTable table = d.Tables[0].Copy();
            d.Dispose();
            return table;
        }

        /// <summary>
        /// Opens OpenFileDialog and returns the selected file path.
        /// </summary>
        /// <param name="fileType">Desired file type for imported file. Example: "csv"</param>
        /// <returns></returns>
        public static string GetOpenFilePath(string fileType)
        {
            string importFileName = "";
            OpenFileDialog openFileDialog = new OpenFileDialog
            {
                Filter = fileType + " files (*." + fileType + ")|*." + fileType + "|All files (*.*)|*.*",
                RestoreDirectory = true
            };
            DialogResult result = openFileDialog.ShowDialog(); // Show the dialog.
            if (result == DialogResult.OK) // Test result.
            {
                importFileName = openFileDialog.FileName;
            }
            openFileDialog.Dispose();
            return importFileName;
        }

        /// <summary>
        /// Opens SaveFileDialog and returns the selected file path.
        /// </summary>
        /// <param name="fileType">Desired file type for exported file. Example: "csv"</param>
        /// <returns></returns>
        public static string GetSaveFilePath(string fileType)
        {
            string exportFileName = "";
            SaveFileDialog saveFileDialog = new SaveFileDialog
            {
                Filter = fileType + " file |*." + fileType + "|All files (*.*)|*.*" //"JPeg Image|*.jpg|Bitmap Image|*.bmp|Gif Image|*.gif";
            };
            //saveFileDialog.Title = "Save an Image File";
            DialogResult result = saveFileDialog.ShowDialog();

            if (result == DialogResult.OK) // Test result.
            {
                exportFileName = saveFileDialog.FileName;
            }
            saveFileDialog.Dispose();
            return exportFileName;
        }

        /// <summary>
        /// Opens an OpenFileDialog for any file type.
        /// </summary>
        /// <returns>Import file path</returns>
        public static string GetOpenFilePath()
        {
            string importFileName = "";
            OpenFileDialog openFileDialog = new OpenFileDialog
            {
                Filter = "|All files (*.*)|*.*",
                RestoreDirectory = true
            };
            DialogResult result = openFileDialog.ShowDialog(); // Show the dialog.
            if (result == DialogResult.OK) // Test result.
            {
                importFileName = openFileDialog.FileName;
            }
            openFileDialog.Dispose();
            return importFileName;
        }

        /// <summary>
        /// Opens a SaveFileDialog for any file type.
        /// </summary>
        /// <returns>Export file path</returns>
        public static string GetSaveFilePath()
        {
            string exportFileName = "";
            SaveFileDialog saveFileDialog = new SaveFileDialog
            {
                Filter = "|All files (*.*)|*.*" //"JPeg Image|*.jpg|Bitmap Image|*.bmp|Gif Image|*.gif";
            };

            DialogResult result = saveFileDialog.ShowDialog();
            if (result == DialogResult.OK) // Test result.
            {
                exportFileName = saveFileDialog.FileName;
            }
            saveFileDialog.Dispose();
            return exportFileName;
        }

        /// <summary>
        /// Exports a DataTable to .xlsx
        /// </summary>
        /// <param name="table">DataTable to export to .xlxs</param>
        public static void ExportTable(DataTable table)
        {
            Console.WriteLine("\nSelect a save destination...");
            //Export datatable to xlsx file
            string saveFilePath = GetSaveFilePath("xlsx");
            Excel.CreateExcelDocument(table, saveFilePath);
            if (File.Exists(saveFilePath))
            {
                Console.WriteLine("\nFile saved...");
            }
            else
            {
                Console.WriteLine("\nFile not saved successfully. Something went wrong...");
            }
        }

        /// <summary>
        /// Exports a DataTable to .xlsx while excluding specific columns
        /// </summary>
        /// <param name="table">DataTable to export to .xlxs</param>
        /// <param name="excludedColumnNames">List of column names to exclude from the export</param>
        public static void ExportTable(DataTable table, List<string> excludedColumnNames)
        {
            Console.WriteLine("\nSelect a save destination...");
            //Export datatable to xlsx file
            string saveFilePath = GetSaveFilePath("xlsx");

            var trimmedTable = table.Copy();
            foreach (string column in excludedColumnNames)
            {
                if (table.Columns.Contains(column))
                {
                    if (trimmedTable.PrimaryKey.Contains(trimmedTable.Columns[column]))
                    {
                        trimmedTable.PrimaryKey = null;
                    }
                    trimmedTable.Columns.Remove(column);
                }
            }

            Excel.CreateExcelDocument(trimmedTable, saveFilePath);
            if (File.Exists(saveFilePath))
            {
                Console.WriteLine("\nFile saved...");
            }
            else
            {
                Console.WriteLine("\nFile not saved successfully. Something went wrong...");
            }
        }

        /// <summary>
        /// Given a filepath, will import file to a new List<string>
        /// </summary>
        /// <param name="filepath">Import file path</param>
        /// <returns>A list of type string containing the lines of the imported file</returns>
        public static List<string> ImportFileToStringList(string filepath)
        {
            List<string> lines = System.IO.File.ReadAllLines(filepath).ToList();
            return lines;
        }

        /// <summary>
        /// Opens a OpenFileDialog and will import file to a new List<string>
        /// </summary>
        /// <param name="filepath">Import file path</param>
        /// <returns>A list of type string containing the lines of the imported file</returns>
        public static List<string> ImportFileToStringList()
        {
            string filepath = GetOpenFilePath();
            List<string> lines = System.IO.File.ReadAllLines(filepath).ToList();
            return lines;
        }

        /// <summary>
        /// Exports an IEnumerable of strings to file
        /// </summary>
        /// <param name="filepath">Export file path, including file extension</param>
        /// <param name="lines">List of strings to export to file</param>
        public static void ExportStringsToFile(string filepath, IEnumerable<string> lines)
        {
            System.IO.File.WriteAllLines(filepath, lines);
        }

        /// <summary>
        /// Imports a CSV file to a list of strings
        /// </summary>
        /// <param name="filepath">Import file path</param>
        /// <returns></returns>
        public static List<string> ImportCsvToStringList(string filepath)
        {
            if (!filepath.EndsWith(".csv"))
            {
                filepath += ".csv";
            }
            List<string> csv = System.IO.File.ReadAllLines(filepath).ToList();
            return csv;
        }

        /// <summary>
        /// Imports a CSV file to a single string
        /// </summary>
        /// <param name="filepath">Import file path</param>
        /// <returns>A string contianing the entire imported CSV</returns>
        public static string ImportCsvToString(string filepath)
        {
            string csv = File.ReadAllText(filepath, Encoding.ASCII);

            return csv;
        }

        /// <summary>
        /// Imports a CSV file and outputs a DataTable (All non-Guid columns as type string)
        /// </summary>
        /// <param name="csvFilePath">File path of CSV file to be imported</param>
        /// <param name="delimiter">Delimiter used in the CSV file</param>
        /// <param name="needsGuid">True will add a Guid to DataTable, false will treat the first column as the Guid</param>
        /// <returns>A DataTable</returns>
        public static DataTable ImportTableFromCsv(string csvFilePath, char delimiter, bool needsGuid)
        {
            //Check if file exists
            if (File.Exists(csvFilePath)) //If exists
            {
                Console.WriteLine("CSV file exists. Importing...");

                //Import file
                List<string> csv = ImportCsvToStringList(csvFilePath);

                DataTable table = TableBuilder.BuildTableSchema("A Table", TableBuilder.GetHeaders(csv, delimiter), needsGuid);

                //Populate table from file
                if (csvFilePath.EndsWith(".csv"))
                {
                    csvFilePath = csvFilePath.Remove(csvFilePath.Length - 4);
                }

                table.PopulateTableFromCsv(csvFilePath, delimiter, true, needsGuid);
                return table;
            }
            else //If not exists
            {
                return null;
            }
        }

        /// <summary>
        /// Imports a CSV file and outputs a DataTable (Uses provided column types)
        /// </summary>
        /// <param name="csvFilePath">File path of CSV file to be imported</param>
        /// <param name="delimiter">Delimiter used in the CSV file</param>
        /// <param name="columnTypes">Array of type Type with all of the desired DataTypes for the DataTable's columns, in order.</param>
        /// <param name="needsGuid">True will add a Guid to DataTable, false will treat the first column as the Guid</param>
        /// <returns></returns>
        public static DataTable ImportTableFromCsv(string csvFilePath, char delimiter, Type[] columnTypes, bool needsGuid)
        {
            //Check if file exists
            if (File.Exists(csvFilePath)) //If exists
            {
                Console.WriteLine("CSV file exists. Importing...");

                //Import file
                List<string> csv = ImportCsvToStringList(csvFilePath);

                DataTable table = table = TableBuilder.BuildTableSchema("A Table", TableBuilder.GetHeaders(csv, delimiter), columnTypes, needsGuid);

                //Populate table from file
                if (csvFilePath.EndsWith(".csv"))
                {
                    csvFilePath = csvFilePath.Remove(csvFilePath.Length - 4);
                }

                table.PopulateTableFromCsv(csvFilePath, delimiter, true, needsGuid);
                return table;
            }
            else //If not exists
            {
                return null;
            }
        }

        /// <summary>
        /// Exports DataTable to CSV
        /// </summary>
        /// <param name="table">DataTable to export</param>
        /// <param name="folderPath">Export directory path</param>
        /// <param name="fileName">Export file name without file extension</param>
        public static void ExportDataTableToCSV(DataTable table, string folderPath, string fileName)
        {
            using (StreamWriter writer = new StreamWriter(folderPath + "\\" + fileName + ".csv"))
            {
                WriteDataTable(table, writer, true);
                writer.Close();
            }
        }

        /// <summary>
        /// Exports DataTable to CSV
        /// </summary>
        /// <param name="table">DataTable to export</param>
        /// <param name="fullPath">Export file path</param>
        public static void ExportDataTableToCSV(DataTable table, string fullPath)
        {
            using (StreamWriter writer = new StreamWriter(fullPath))
            {
                WriteDataTable(table, writer, true);
                writer.Close();
            }
        }

        /// <summary>
        /// Writes DataTable to file using filepath in TextWriter
        /// </summary>
        /// <param name="sourceTable">DataTable to export</param>
        /// <param name="writer">Writer used to write the file</param>
        /// <param name="includeHeaders">Whether to include headers during the export</param>
        private static void WriteDataTable(DataTable sourceTable, TextWriter writer, bool includeHeaders)
        {
            if (includeHeaders)
            {
                IEnumerable<String> headerValues = sourceTable.Columns
                    .OfType<DataColumn>()
                    .Select(column => QuoteValue(column.ColumnName));

                writer.WriteLine(String.Join("|", headerValues));
            }

            IEnumerable<String> items = null;

            foreach (DataRow row in sourceTable.Rows)
            {
                items = row.ItemArray.Select(o => QuoteValue(o.ToString()));
                writer.WriteLine(String.Join(",", items));
            }

            writer.Flush();
        }

        private static string QuoteValue(string value)
        {
            return String.Concat("\"",
                                value.Replace("\"",
                                "\"\""),
                                "\"");
        }
    }
}