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
using NPOI.XSSF.Model;

namespace XLSXIO.Export
{
    public class XLSXExport : XLSXBase
    {
        public XLSXExport(IEnumerable<XLSColumnTemplate> templates) : base(templates) { }
        /// <summary>
        /// Применить стиль к ячейке
        /// </summary>
        /// <param name="cell">Ячейка, к котрой применяется стиль</param>
        void ApplyCellStyle(ICell cell)
        {
            cell.CellStyle.BorderLeft = BorderStyle.Thick;
            cell.CellStyle.BorderRight = BorderStyle.Thick;
            cell.CellStyle.BorderBottom = BorderStyle.Thick;
            cell.CellStyle.BorderTop = BorderStyle.Thick;
        }
        /// <summary>
        /// Создает строку с заголовками столбцов
        /// </summary>
        /// <param name="sheet">Таблица, в которую будут заноситься сведения</param>
        void CreateHeadersRow(ISheet sheet)
        {
            IRow headersRow = sheet.CreateRow(0);
            for (int i = 0; i < Columns.Count; i++)
            {
                var cell = headersRow.CreateCell(i);
                ApplyCellStyle(cell);
                cell.SetCellValue(Columns.ElementAt(i).Key.InDocumentName);
                Columns[Columns.ElementAt(i).Key] = i;
            }
        }
        /// <summary>
        /// Заполняет созданную таблицу данными
        /// </summary>
        /// <param name="sheet">Таблица, в которую будут заноситься сведения</param>
        /// <param name="data">Объект DataTable, содержащий сведения для экспорта</param>
        void FillSheetWithValues(ISheet sheet, DataTable data)
        {
            for (int rowIndex = 0; rowIndex < data.Rows.Count; rowIndex++)
            {
                var exportRow = sheet.CreateRow(rowIndex + 1);
                foreach (var key in Columns.Keys)
                {
                    exportRow.CreateCell(Columns[key]).SetCellValue(data.Rows[rowIndex][key.InDatabaseName].ToString());
                }
            }
        }
        /// <summary>
        /// Выгружает данные из объекта DataTable в файл xlsx
        /// </summary>
        /// <param name="filename">Имя целевого файла</param>
        /// <param name="data">Данные для экспорта</param>
        public void Upload(string filename, DataTable data)
        {
            using (var stream = new FileStream(filename, FileMode.OpenOrCreate, FileAccess.Write))
            {
                IWorkbook book = new XSSFWorkbook();
                ISheet sheet = book.CreateSheet($"Export_{DateTime.Today:yyyy-MM-dd}");
                
                CreateHeadersRow(sheet);
                FillSheetWithValues(sheet, data);
                
                book.Write(stream);
            }
        }
    }
}
