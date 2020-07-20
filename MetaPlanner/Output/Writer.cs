using System;
using System.Text;
using Windows.Storage;
using CsvHelper;
using System.Globalization;
using System.IO;
using System.Collections;
using Windows.UI.Xaml.Data;
using System.Threading.Tasks;

namespace MetaPlanner.Output
{
    class Writer
    {
        public async Task Write(IEnumerable list, StorageFolder storageFolder,string fileName)
        {
            // Create  file; replace if exists.
            Windows.Storage.StorageFile file = await storageFolder.CreateFileAsync(fileName, Windows.Storage.CreationCollisionOption.ReplaceExisting);
            var writer = new StreamWriter(file.Path, false, Encoding.UTF8);
            var csv = new CsvWriter(writer, CultureInfo.InvariantCulture);
            csv.Configuration.HasHeaderRecord = true;
            csv.WriteRecords(list);
            csv.Flush();
            writer.Close();
        }
    }
}
