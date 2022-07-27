using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace XLSXIO.AuxiliaryTypes
{
    public class XLSColumnTemplate
    {
        string inDocumentName;
        string inDatabaseName;
        Type type;
        /// <summary>
        /// Имя столбца в документе xls(x)
        /// </summary>
        public string InDocumentName
        {
            get => inDocumentName;
            set
            {
                if (string.IsNullOrEmpty(value)) throw new ArgumentException("Наименование столбца не должно быть пустым");
                inDocumentName = value;
            }
        }
        /// <summary>
        /// Имя столбца в выходном DataTable
        /// </summary>
        public string InDatabaseName
        {
            get => inDatabaseName;
            set
            {
                if (string.IsNullOrEmpty(value)) throw new ArgumentException("Наименование столбца в БД не должно быть пустым");
                inDatabaseName = value;
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
        /// <summary>
        /// Конструктор объявления ожидаемого столбца в документе на импорт
        /// </summary>
        /// <param name="inDocumentName">Имя столбца в документе xls(x)</param>
        /// <param name="inDatabaseName">Имя столбца в выходном DataTable</param>
        /// <param name="type">Предполагаемый тип данных, хранимый в ячейке</param>
        public XLSColumnTemplate(string inDocumentName, string inDatabaseName, Type type)
        {
            this.inDocumentName = inDocumentName;
            this.inDatabaseName = inDatabaseName;
            this.type = type;
        }
    }
}
