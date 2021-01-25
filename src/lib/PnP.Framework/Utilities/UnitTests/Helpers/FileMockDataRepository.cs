using Newtonsoft.Json;
using PnP.Framework.Utilities.UnitTests.Model;
using System;
using System.Collections.Generic;
using System.IO;
using System.Text;

namespace PnP.Framework.Utilities.UnitTests.Helpers
{
    public class FileMockDataRepository<T> : IMockDataRepository<MockResponseEntry<T>>
    {
        public string FilePath { get; set; }
        public FileMockDataRepository(string filePath)
        {
            FilePath = filePath;
        }
        public List<MockResponseEntry<T>> LoadMockData()
        {
            string serializedData = File.ReadAllText(FilePath);
            List<MockResponseEntry<T>> result = JsonConvert.DeserializeObject<List<MockResponseEntry<T>>>(serializedData, MockResponseEntry.SerializerSettings);

            return result;
        }

        public void SaveMockData(List<MockResponseEntry<T>> mockedData)
        {
            string serializedData = JsonConvert.SerializeObject(mockedData, MockResponseEntry.SerializerSettings);
            using (StreamWriter outputFile = new StreamWriter(FilePath))
            {
                outputFile.Write(serializedData);
            }
        }
    }
}
