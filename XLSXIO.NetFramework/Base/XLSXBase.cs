using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using XLSXIO.AuxiliaryTypes;

namespace XLSXIO.Base
{
    public class XLSXBase
    {
        protected readonly string NEW_LINE_CHAR = Environment.NewLine;
        /// <summary>
        /// Коллекция описаний столбцов, которые должны находиться в документе на импорт/экспорт
        /// </summary>
        protected Dictionary<XLSColumnTemplate, int> Columns = new Dictionary<XLSColumnTemplate, int>();
        /// <summary>
        /// Конструктор класса
        /// </summary>
        /// <param name="templates">Коллекция описаний столбцов, которые должны находиться в документе на импорт/экспорт</param>
        public XLSXBase(IEnumerable<XLSColumnTemplate> templates)
        {
            Columns.Clear();
            foreach (var template in templates)
            {
                Columns.Add(template, -1);
            }
        }
    }
}
