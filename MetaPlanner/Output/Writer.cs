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
using System.Collections;

namespace MetaPlanner.Output
{
    class Writer
    {
        public async void Write(IEnumerable list, StorageFolder storageFolder,string fileName)
        {
            // Create  file; replace if exists.
            Windows.Storage.StorageFile file = await storageFolder.CreateFileAsync(fileName, Windows.Storage.CreationCollisionOption.ReplaceExisting);
            var writer = new StreamWriter(file.Path, false, Encoding.UTF8);
            var csv = new CsvWriter(writer, CultureInfo.InvariantCulture);
            csv.Configuration.HasHeaderRecord = true;
            csv.WriteRecords(list);
        }
    }
}
