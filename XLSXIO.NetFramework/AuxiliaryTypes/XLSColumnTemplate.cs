using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace XLSXIO.NetFramework.AuxiliaryTypes
{
    public class XLSColumnTemplate
    {
        string name;
        Type type;
        /// <summary>
        /// Название столбца
        /// </summary>
        public string Name
        {
            get => name;
            set
            {
                if (string.IsNullOrEmpty(value)) throw new ArgumentException("Наименование столбца не должно быть пустым");
                name = value;
            }
        }
        /// <summary>
        /// Предполагаемый тип данных, хранимый в ячейке
        /// </summary>
        public Type Type
        {
            get => type;
            set
            {
                if (value == null) throw new ArgumentNullException("Предполагаемый тип столбца не может быть NULL");
                type = value;
            }
        }

        public XLSColumnTemplate(string name, Type type)
        {
            this.name = name;
            this.type = type;
        }
    }
}
