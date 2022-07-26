using System;
using System.Collections;
using System.Collections.Generic;

namespace XLSXIO.NetFramework.AuxiliaryTypes
{
    public class XLSColumnTemplatesCollection : IEnumerable<XLSColumnTemplate>
    {
        List<XLSColumnTemplate> columns = new List<XLSColumnTemplate>();

        public void Add(string name, Type type)
        {
            columns.Add(new XLSColumnTemplate(name, type));
        }

        public IEnumerator<XLSColumnTemplate> GetEnumerator()
        {
            return columns.GetEnumerator();
        }

        IEnumerator IEnumerable.GetEnumerator()
        {
            return GetEnumerator();
        }
    }
}
