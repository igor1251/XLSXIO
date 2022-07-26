using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.IO;
using XLSXIO.NetFramework.AuxiliaryTypes;
using System.Data;
using NPOI.SS.UserModel;
using NPOI.HSSF.UserModel;
using NPOI.XSSF.UserModel;

namespace XLSXIO.NetFramework.Import
{
    public class XLSXImport
    {
        Dictionary<XLSColumnTemplate, int> columns = new Dictionary<XLSColumnTemplate, int>();
        DataTable result = new DataTable();

        public XLSXImport(IEnumerable<XLSColumnTemplate> columnTemplates)
        {
            foreach (var template in columnTemplates)
            {
                result.Columns.Add(new DataColumn(template.Name.ToLower()));
                columns.Add(template, -1);
            }
        }

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

        public DataTable Load(string filename)
        {
            var book = LoadWorkbook(filename);
            var sheet = book.GetSheetAt(0);

            var headersRow = sheet.GetRow(0);
            for (int i = 0; i < headersRow.LastCellNum; i++)
            {
                if (headersRow.Cells[i] == null) break;
                var column = columns.FirstOrDefault(x => x.Key.Name.ToLower() == headersRow.Cells[i].ToString().ToLower());
                if (!column.Equals(default(KeyValuePair<XLSColumnTemplate, int>)))
                {
                    columns[column.Key] = i;
                }
            }

            var errorText = new StringBuilder();
            bool columnNotFound = false;
            foreach (var column in columns)
            {
                columnNotFound = column.Value < 0;
                if (columnNotFound)
                {
                    errorText.AppendLine(column.Key.Name);
                }
            }
            if (errorText.Length > 0)
            {
                errorText.Insert(0, "Не найдены следующие столбцы:\n");
                throw new Exception(errorText.ToString());
            }

            object[] valuesCollection = new object[columns.Count];
            int j = 0;
            for (int i = 1; i <= sheet.LastRowNum; i++)
            {
                var row = sheet.GetRow(i);
                foreach (var key in columns.Keys)
                {
                    var cellIndex = columns[key];
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
                            else if (key.Type == typeof(Double)) valuesCollection[j] = Convert.ToDouble(cellValue.NumericCellValue);
                            else if (key.Type == typeof(Single)) valuesCollection[j] = Convert.ToSingle(cellValue.NumericCellValue);
                            else if (key.Type == typeof(Boolean)) valuesCollection[j] = Convert.ToBoolean(cellValue.BooleanCellValue);
                            else valuesCollection[j] = cellValue.ToString();
                        }
                        catch (Exception ex)
                        {
                            throw new Exception($"Значение в ячейке {cellValue.Address} имеет неправильный формат.\nОжидаемый формат: {key.Type}\nЗначение, вызвавшее ошибку: {cellValue}");
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
    }
}
