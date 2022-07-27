using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.IO;
using System.Data;
using NPOI.SS.UserModel;
using NPOI.HSSF.UserModel;
using NPOI.XSSF.UserModel;
using XLSXIO.AuxiliaryTypes;
using XLSXIO.Base;

namespace XLSXIO.Import
{
    public class XLSXImport : XLSXBase
    {
        public XLSXImport(IEnumerable<XLSColumnTemplate> templates) : base(templates) { }
        /// <summary>
        /// Загружает книгу из документа
        /// </summary>
        /// <param name="filename">Имя xls(x) файла</param>
        /// <returns>Рабочая книга</returns>
        /// <exception cref="Exception">Возникает при попытке разобрать файл с неподдерживаемым расширением</exception>
        IWorkbook LoadWorkbook(string filename)
        {
            var extension = Path.GetExtension(filename);
            using (var stream = new FileStream(filename, FileMode.Open, FileAccess.Read))
            {
                switch (extension.ToLower())
                {
                    case ".xls":
                        return new HSSFWorkbook(stream);
                    case ".xlsx":
                        return new XSSFWorkbook(stream);
                    default:
                        throw new Exception("Неподдерживаемое расширение файла");
                }
            }
        }
        /// <summary>
        /// Загружает столбцы из документа и запоминает индекс каждого, который имеет вхождение в 
        /// коллекцию описанных столбцов, переданных в конструктор
        /// </summary>
        /// <param name="sheet">Лист Excel</param>
        void LoadColumns(ISheet sheet)
        {
            var headersRow = sheet.GetRow(0);
            for (int i = 0; i < headersRow.LastCellNum; i++)
            {
                if (headersRow.Cells[i] == null) break;
                var column = Columns.FirstOrDefault(x => x.Key.InDocumentName.ToLower() == headersRow.Cells[i].ToString().ToLower());
                if (!column.Equals(default(KeyValuePair<XLSColumnTemplate, int>)))
                {
                    Columns[column.Key] = i;
                }
            }
        }
        /// <summary>
        /// Проверяет, были ли найдены все требуемые столбцы в документе
        /// </summary>
        /// <exception cref="Exception">Вызывается, если индекс хотя бы одного столбца равен -1 (значение по-умолчанию)</exception>
        void CheckDocumentForErrors()
        {
            var errorText = new StringBuilder();
            bool columnNotFound = false;
            foreach (var column in Columns)
            {
                columnNotFound = column.Value < 0;
                if (columnNotFound)
                {
                    errorText.AppendLine(column.Key.InDocumentName);
                }
            }
            if (errorText.Length > 0)
            {
                errorText.Insert(0, $"Не найдены следующие столбцы:{NEW_LINE_CHAR}");
                throw new Exception(errorText.ToString());
            }
        }
        /// <summary>
        /// Выполняет разбор значений, хранимых в активном листе документа
        /// </summary>
        /// <param name="sheet">Лист Excel</param>
        /// <returns>Содержимое документа в виде объекта DataTable</returns>
        /// <exception cref="Exception">Возникает при несоответствии ожидаемого и хранимого форматов в документе</exception>
        DataTable ParseSheet(ISheet sheet)
        {
            var result = new DataTable();
            foreach (var template in Columns)
            {
                result.Columns.Add(new DataColumn(template.Key.InDatabaseName));
            }

            object[] valuesCollection = new object[Columns.Count];
            int j = 0;
            for (int i = 1; i <= sheet.LastRowNum; i++)
            {
                var row = sheet.GetRow(i);
                foreach (var key in Columns.Keys)
                {
                    var cellIndex = Columns[key];
                    var cellValue = row.GetCell(cellIndex);
                    if (cellValue != null)
                    {
                        try
                        {
                            if (key.Type == typeof(UInt16)) valuesCollection[j] = Convert.ToUInt16(cellValue.NumericCellValue);
                            else if (key.Type == typeof(UInt32)) valuesCollection[j] = Convert.ToUInt32(cellValue.NumericCellValue);
                            else if (key.Type == typeof(UInt64)) valuesCollection[j] = Convert.ToUInt64(cellValue.NumericCellValue);
                            else if (key.Type == typeof(Int16)) valuesCollection[j] = Convert.ToInt16(cellValue.NumericCellValue);
                            else if (key.Type == typeof(Int32)) valuesCollection[j] = Convert.ToInt32(cellValue.NumericCellValue);
                            else if (key.Type == typeof(Int64)) valuesCollection[j] = Convert.ToInt64(cellValue.NumericCellValue);
                            else if (key.Type == typeof(DateTime)) valuesCollection[j] = Convert.ToDateTime(cellValue.DateCellValue);
                            else if (key.Type == typeof(Double)) valuesCollection[j] = cellValue.NumericCellValue;
                            else if (key.Type == typeof(Single)) valuesCollection[j] = Convert.ToSingle(cellValue.NumericCellValue);
                            else if (key.Type == typeof(Boolean)) valuesCollection[j] = Convert.ToBoolean(cellValue.BooleanCellValue);
                            else valuesCollection[j] = cellValue.ToString();
                        }
                        catch (Exception ex)
                        {
                            throw new Exception($"Значение в ячейке {cellValue.Address} имеет неправильный формат.{NEW_LINE_CHAR}Ожидаемый формат: {key.Type}{NEW_LINE_CHAR}Значение, вызвавшее ошибку: {cellValue}");
                        }
                    }
                    else valuesCollection[j] = null;
                    j++;
                }
                j = 0;
                result.Rows.Add(valuesCollection);
            }

            return result;
        }
        /// <summary>
        /// Выполняет загрузку содержимого файла
        /// </summary>
        /// <param name="filename">Имя xls(x) файла</param>
        /// <returns>Объект DataTable с содержимым файла</returns>
        public DataTable Load(string filename)
        {
            var book = LoadWorkbook(filename);
            var sheet = book.GetSheetAt(0);

            LoadColumns(sheet);
            CheckDocumentForErrors();
            var loadResult = ParseSheet(sheet);
            return loadResult;
        }
    }
}
