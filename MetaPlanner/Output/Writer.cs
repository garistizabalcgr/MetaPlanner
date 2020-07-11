using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Windows.Storage;
using CsvHelper;
using System.Globalization;
using MetaPlanner.Model;
using System.IO;

namespace MetaPlanner.Output
{
    class Writer
    {

        public async void WritePlans(IEnumerable<Plan> list, string name)
        {
            StorageFolder storageFolder = Windows.Storage.ApplicationData.Current.LocalFolder;
            String fileName = String.Format("{0:D4}", DateTime.Now.Year) + "-" + String.Format("{0:D2}", DateTime.Now.Month) + "-" + String.Format("{0:D2}", DateTime.Now.Day) + "_" + String.Format("{0:D2}", DateTime.Now.Hour) + "_" + String.Format("{0:D2}", DateTime.Now.Minute) + "_" + String.Format("{0:D2}", DateTime.Now.Second) + " " + name + ".csv";
            // Create sample file; replace if exists.
            Windows.Storage.StorageFile file = await storageFolder.CreateFileAsync(fileName, Windows.Storage.CreationCollisionOption.ReplaceExisting);
            var writer = new StreamWriter(file.Path, false, Encoding.UTF8);
            var csv = new CsvWriter(writer, CultureInfo.InvariantCulture);
            csv.Configuration.HasHeaderRecord = true;
            csv.WriteRecords(list);
        }

        public async void WriteBuckets(IEnumerable<Bucket> list, string name)
        {
            StorageFolder storageFolder = Windows.Storage.ApplicationData.Current.LocalFolder;
            String fileName = String.Format("{0:D4}", DateTime.Now.Year) + "-" + String.Format("{0:D2}", DateTime.Now.Month) + "-" + String.Format("{0:D2}", DateTime.Now.Day) + "_" + String.Format("{0:D2}", DateTime.Now.Hour) + "_" + String.Format("{0:D2}", DateTime.Now.Minute) + "_" + String.Format("{0:D2}", DateTime.Now.Second) + " " + name + ".csv";
            // Create sample file; replace if exists.
            Windows.Storage.StorageFile file = await storageFolder.CreateFileAsync(fileName, Windows.Storage.CreationCollisionOption.ReplaceExisting);
            var writer = new StreamWriter(file.Path, false, Encoding.UTF8);
            var csv = new CsvWriter(writer, CultureInfo.InvariantCulture);
            csv.Configuration.HasHeaderRecord = true;
            csv.WriteRecords(list);
        }


        public async void WriteTasks(IEnumerable<PTask> list, string name)
        {
            StorageFolder storageFolder = Windows.Storage.ApplicationData.Current.LocalFolder;
            String fileName = String.Format("{0:D4}", DateTime.Now.Year) + "-" + String.Format("{0:D2}", DateTime.Now.Month) + "-" + String.Format("{0:D2}", DateTime.Now.Day) + "_" + String.Format("{0:D2}", DateTime.Now.Hour) + "_" + String.Format("{0:D2}", DateTime.Now.Minute) + "_" + String.Format("{0:D2}", DateTime.Now.Second) + " " + name + ".csv";
            // Create sample file; replace if exists.
            Windows.Storage.StorageFile file = await storageFolder.CreateFileAsync(fileName, Windows.Storage.CreationCollisionOption.ReplaceExisting);
            var writer = new StreamWriter(file.Path, false, Encoding.UTF8);
            var csv = new CsvWriter(writer, CultureInfo.InvariantCulture);
            csv.Configuration.HasHeaderRecord = true;
            csv.WriteRecords(list);
        }


        public async void WriteAssignees(IEnumerable<Assignment> list, string name)
        {
            StorageFolder storageFolder = Windows.Storage.ApplicationData.Current.LocalFolder;
            String fileName = String.Format("{0:D4}", DateTime.Now.Year) + "-" + String.Format("{0:D2}", DateTime.Now.Month) + "-" + String.Format("{0:D2}", DateTime.Now.Day) + "_" + String.Format("{0:D2}", DateTime.Now.Hour) + "_" + String.Format("{0:D2}", DateTime.Now.Minute) + "_" + String.Format("{0:D2}", DateTime.Now.Second) + " " + name + ".csv";
            // Create sample file; replace if exists.
            Windows.Storage.StorageFile file = await storageFolder.CreateFileAsync(fileName, Windows.Storage.CreationCollisionOption.ReplaceExisting);
            var writer = new StreamWriter(file.Path, false, Encoding.UTF8);
            var csv = new CsvWriter(writer, CultureInfo.InvariantCulture);
            csv.Configuration.HasHeaderRecord = true;
            csv.WriteRecords(list);
        }
    }
}
