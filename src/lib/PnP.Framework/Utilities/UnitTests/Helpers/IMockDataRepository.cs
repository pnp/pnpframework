using PnP.Framework.Utilities.UnitTests.Model;
using System;
using System.Collections.Generic;
using System.Text;

namespace PnP.Framework.Utilities.UnitTests.Helpers
{
    public interface IMockDataRepository : IMockDataRepository<MockResponseEntry<object>>
    {
    }
    public interface IMockDataRepository<T>
    {
        void SaveMockData(List<T> mockedData);
        List<T> LoadMockData();
    }
}
