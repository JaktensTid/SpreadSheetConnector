using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.IO;
using System.Runtime.Serialization;
using System.Runtime.Serialization.Formatters.Binary;

namespace SpreadSheetConnector
{
    public enum UpdaterAction
    {
        Overwrite,
        Replace
    }

    [Serializable]
    public class UpdaterItem : IDisposable
    {
        public const string SaveFolder = "Data";
        private string localPath;
        private string googlePath;
        private Tuple<int, int> range;
        private DateTime date;

        public string LocalPath { get { return localPath; } }
        public string GooglePath { get { return googlePath; } }
        public Tuple<int, int> Range { get { return range; } }
        public DateTime Date { get { return date; } }
        [NonSerialized]
        public FileSystemWatcher watchdog;

        public UpdaterItem(string localPath, string googlePath, Tuple<int, int> range, DateTime date)
        {
            if (!Directory.Exists(SaveFolder))
            {
                Directory.CreateDirectory(SaveFolder);
            }
            if (!File.Exists(localPath))
                throw new FileNotFoundException("File not exists", localPath);
            if (googlePath.Split('=').Length == 0)
                throw new ArgumentException("Google path not exists", googlePath);

            this.localPath = localPath;
            this.googlePath = googlePath;
            this.range = range;
            this.date = date;
        }

        public void StartWatch()
        {
            watchdog = new FileSystemWatcher() { Path = new FileInfo(localPath).Directory.FullName,
                NotifyFilter = NotifyFilters.LastWrite | NotifyFilters.CreationTime,
            Filter = localPath};
            watchdog.Changed += new FileSystemEventHandler(watchdog_Changed);
            watchdog.EnableRaisingEvents = true;
        }

        private void watchdog_Changed(object sender, FileSystemEventArgs e)
        {
            FileInfo file = new FileInfo(e.FullPath);

        }

        private void StopWatch()
        { }

        public void Dispose()
        {
            StopWatch();
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
                string path = Path.Combine(UpdaterItem.SaveFolder, item.GetHashCode().ToString() + ".dat");
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

            if (Directory.Exists(UpdaterItem.SaveFolder))
            {
                foreach (FileInfo file in Directory.EnumerateFiles(UpdaterItem.SaveFolder).Select(path => new FileInfo(path)))
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

