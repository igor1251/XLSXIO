using System;
using System.Collections;
using System.Collections.Generic;

namespace XLSXIO.AuxiliaryTypes
{
    public class XLSColumnTemplatesCollection : IEnumerable<XLSColumnTemplate>
    {
        List<XLSColumnTemplate> columns = new List<XLSColumnTemplate>();
        /// <summary>
        /// Добавляет объявление ожидаемого столбца в документе на импорт
        /// </summary>
        /// <param name="inDocumentName">Имя столбца в документе xls(x)</param>
        /// <param name="inDatabaseName">Имя столбца в выходном DataTable</param>
        /// <param name="type">Предполагаемый тип данных, хранимый в ячейке</param>
        public void Add(string inDocumentName, string inDatabaseName, Type type)
        {
            columns.Add(new XLSColumnTemplate(inDocumentName, inDatabaseName, type));
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
