using System;
using System.Linq;
using System.Data;
using System.ComponentModel;
using System.IO;
using System.Runtime.Serialization;
using System.Runtime.Serialization.Formatters.Binary;
using Excel;
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
                else
                {
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
            return new DataTable();
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
                if (googlePath.Contains("/d/"))
                {
                    int pFrom = googlePath.IndexOf("/d/") + "/d/".Length;
                    int pTo = googlePath.LastIndexOf("/");
                    return googlePath.Substring(pFrom, pTo - pFrom);
                }
                return "";
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
                throw new FileNotFoundException("Directory not exists", localPath);
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
                NotifyFilter = NotifyFilters.LastWrite,
                Filter = "*.*"
            };
            watchdog.Changed += new FileSystemEventHandler(Changed);
            watchdog.EnableRaisingEvents = true;
        }

        private async void Changed(object sender, FileSystemEventArgs e)
        {
            try
            {
                var ext = new FileInfo(e.FullPath).Extension;

                if (ext != ".csv" & 
                    ext != ".xls" & 
                    ext != ".xlsx") return;
                watchdog.EnableRaisingEvents = false;
                date = DateTime.Now;
                DataTable table = DataReader.Read(e.FullPath);
                if (table != null)
                {
                    int deleteRowsRange = 0;
                    if(range != null)
                    {
                        deleteRowsRange = range.Item2;
                    }

                    
                    if (action == UpdaterAction.Append)
                    {
                        await connector.Append(table, GoogleSpreadsheetID, deleteRowsRange);
                    }
                    if (action == UpdaterAction.Overwrite)
                    {
                        await connector.Overwrite(table, GoogleSpreadsheetID, deleteRowsRange);
                    }
                }
            }
            catch (IOException)
            {
                
            }
            catch (Google.GoogleApiException exc)
            {
                MessageBox.Show(exc.Message);
            }
            finally
            {
                watchdog.EnableRaisingEvents = true;
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

        public static void SerializeUpdaterItems(BindingList<UpdaterItem> items)
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

        public static BindingList<UpdaterItem> DeserializeUpdaterItems()
        {
            BindingList<UpdaterItem> items = new BindingList<UpdaterItem>();
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

