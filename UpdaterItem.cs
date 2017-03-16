using System;
using System.Collections.Generic;
using System.Linq;
using System.Data;
using System.IO;
using System.Runtime.Serialization;
using System.Runtime.Serialization.Formatters.Binary;
using Excel;
using System.Threading.Tasks;
using CsvHelper;
using System.Windows.Forms;

namespace SpreadSheetConnector
{
    public class DataReader
    {
        public static DataTable Read(string path)
        {
            using (FileStream stream = File.Open(path, FileMode.Open, FileAccess.Read))
            {
                IExcelDataReader excelReader = GetReader(path, stream);
                if (excelReader != null)
                {
                    DataSet set = excelReader.AsDataSet();
                    if (set.Tables.Count > 0)
                    {
                        return set.Tables[0];
                    }
                }
                using (TextReader reader = new StreamReader(stream))
                {
                    DataTable dt = new DataTable();
                    var csv = new CsvReader(reader);
                    csv.ReadHeader();
                    DataColumn[] columns = new DataColumn[csv.FieldHeaders.Length];
                    for (int i = 0; i < csv.FieldHeaders.Length; i++)
                    {
                        columns[i] = new DataColumn(csv.FieldHeaders[i]);
                    }
                    dt.Columns.AddRange(columns);
                    DataRow headerRow = dt.NewRow();
                    for (int i = 0; i < columns.Length; i++)
                    {
                        headerRow[columns[i]] = columns[i];
                    }
                    while (csv.Read())
                    {
                        var row = dt.NewRow();
                        foreach (DataColumn column in dt.Columns)
                        {
                            row[column.ColumnName] = csv.GetField(column.DataType, column.ColumnName);
                        }
                        dt.Rows.Add(row);
                    }
                    return dt;
                }
            }

        }

        private static IExcelDataReader GetReader(string path, Stream stream)
        {
            if (new FileInfo(path).Extension == ".xls")
            {
                return ExcelReaderFactory.CreateBinaryReader(stream);
            }
            if (new FileInfo(path).Extension == ".xlsx")
            {
                return ExcelReaderFactory.CreateOpenXmlReader(stream);
            }
            return null;
        }
    }

    public enum UpdaterAction
    {
        Overwrite,
        Append
    }

    [Serializable]
    public class UpdaterItem : IDisposable
    {
        public static GoogleConnector connector;
        public const string SaveFolder = "Data";
        private string name;
        private string localPath;
        private string googlePath;
        private UpdaterAction action;
        private Tuple<int, int> range;
        private DateTime date;

        public string Name => name;
        public string LocalPath => localPath;
        public string GooglePath => googlePath;
        public string GoogleSpreadsheetID
        {
            get
            {
                if (googlePath.Split('=').Length > 0)
                    return googlePath.Split('=').Last();
                else
                    return null;
            }
        }
        public UpdaterAction Action => action;
        public Tuple<int, int> Range => range;
        //Last used
        public DateTime Date => date;
        [NonSerialized]
        public FileSystemWatcher watchdog;

        public UpdaterItem(string name, string localPath,
            string googlePath, UpdaterAction action,
            Tuple<int, int> range)
        {
            if (!Directory.Exists(SaveFolder))
                Directory.CreateDirectory(SaveFolder);
            if (!Directory.Exists(localPath))
                throw new FileNotFoundException("File not exists", localPath);
            if (googlePath.Split('=').Length == 0)
                throw new ArgumentException("Google path not exists", googlePath);

            this.name = name;
            this.localPath = localPath;
            this.googlePath = googlePath;
            this.action = action;
            this.range = range;
        }

        public void StartWatch()
        {
            watchdog = new FileSystemWatcher()
            {
                Path = localPath,
                NotifyFilter = NotifyFilters.LastWrite | NotifyFilters.CreationTime,
                Filter = "*.csv|*.xls|*.xlsx"
            };
            watchdog.Changed += new FileSystemEventHandler(watchdog_Changed);
            watchdog.EnableRaisingEvents = true;
        }

        private async void watchdog_Changed(object sender, FileSystemEventArgs e)
        {
            date = DateTime.Now;
            DataTable table = DataReader.Read(e.FullPath);
            if (table != null)
            {
                if (range != null)
                {
                    if (table.Rows.Count > range.Item2)
                    {
                        foreach (var i in Enumerable.Range(range.Item1, range.Item2))
                        {
                            table.Rows[i].Delete();
                        }
                    }
                }

                try
                {
                    if (action == UpdaterAction.Append)
                    {
                        await connector.Append(table, GoogleSpreadsheetID);
                    }
                    if (action == UpdaterAction.Overwrite)
                    {
                        await connector.Overwrite(table, GoogleSpreadsheetID);
                    }
                }
                catch (AggregateException ae)
                {
                    ae.Handle((x) =>
                    {
                        if (x is IOException)
                        {
                            MessageBox.Show("Please, close " + e.FullPath);
                        }
                        if (x is Google.GoogleApiException)
                        {
                            MessageBox.Show("Spreadsheets", "Max cells number is 2000000");
                        }
                        return false;
                    });
                }
            }
        }

        public void Dispose()
        {
            string path = Path.Combine(SaveFolder, this.GetHashCode().ToString() + ".dat");
            if (File.Exists(path))
            {
                File.Delete(path);
            }
        }

        public override int GetHashCode()
        {
            return localPath.GetHashCode() + googlePath.GetHashCode();
        }

        public static void SerializeUpdaterItems(List<UpdaterItem> items)
        {
            IFormatter formatter = new BinaryFormatter();
            foreach (UpdaterItem item in items)
            {
                string path = Path.Combine(SaveFolder, item.GetHashCode().ToString() + ".dat");
                if (!File.Exists(path))
                {
                    using (Stream stream = new FileStream(path,
                        FileMode.CreateNew, FileAccess.ReadWrite))
                    {
                        formatter.Serialize(stream, item);
                    }
                }
            }
        }

        public static List<UpdaterItem> DeserializeUpdaterItems()
        {
            List<UpdaterItem> items = new List<UpdaterItem>();
            IFormatter formatter = new BinaryFormatter();

            if (Directory.Exists(SaveFolder))
            {
                foreach (FileInfo file in Directory.EnumerateFiles(UpdaterItem.SaveFolder)
                    .Select(path => new FileInfo(path))
                    .Where(file => file.Extension == ".dat"))
                {
                    using (Stream stream = new FileStream(file.FullName, FileMode.Open, FileAccess.ReadWrite))
                    {
                        items.Add(formatter.Deserialize(stream) as UpdaterItem);
                    }
                }
            }

            return items;
        }
    }
}

