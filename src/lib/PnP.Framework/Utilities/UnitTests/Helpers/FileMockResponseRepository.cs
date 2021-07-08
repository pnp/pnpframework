using Newtonsoft.Json;
using PnP.Framework.Utilities.UnitTests.Model;
using System;
using System.Collections.Generic;
using System.IO;
using System.Text;

namespace PnP.Framework.Utilities.UnitTests.Helpers
{
    public class FileMockResponseRepository : IMockDataRepository<MockResponse>
    {
        public string FilePath { get; set; }
        public FileMockResponseRepository(string filePath)
        {
            FilePath = filePath;
        }
        public List<MockResponse> LoadMockData()
        {
            if (File.Exists(FilePath))
            {
                string serializedData = File.ReadAllText(FilePath);
                List<MockResponse> result = JsonConvert.DeserializeObject<List<MockResponse>>(serializedData, MockResponseEntry.SerializerSettings);

                return result;
            }
            return new List<MockResponse>();
        }

        public void SaveMockData(List<MockResponse> mockedData)
        {
            string serializedData = JsonConvert.SerializeObject(mockedData, MockResponseEntry.SerializerSettings);
            if(!Directory.Exists(FilePath))
            {
                Directory.CreateDirectory(Path.GetDirectoryName(FilePath));
            }
            File.WriteAllText(FilePath, serializedData);
        }
    }
}
