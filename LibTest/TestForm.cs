using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.IO;
using System.Windows.Forms;
using XLSXIO.NetFramework.Import;
using XLSXIO.NetFramework.AuxiliaryTypes;

namespace LibTest
{
    public partial class TestForm : Form
    {
        public TestForm()
        {
            InitializeComponent();
        }

        string PickFileToImport()
        {
            var fileDialog = new OpenFileDialog();
            fileDialog.Filter = "Файлы Excel|*.xls;*.xlsx";
            if (fileDialog.ShowDialog() == DialogResult.OK)
            {
                return fileDialog.FileName;
            }
            return string.Empty;
        }

        private void launchTestButton_Click(object sender, EventArgs e)
        {
            var xls = new XLSXImport(new XLSColumnTemplatesCollection
            {
                { "номер агента", typeof(UInt64) },
                { "номер субагента", typeof(UInt32) },
                { "дата операции БО", typeof(DateTime) },
                { "тип БО", typeof(Int32) },
                { "Сумма", typeof(Double) },
                { "Валюта", typeof(string) },
                { "комментарий", typeof(string) },
                { "номер реферала", typeof(Int32) },
                { "номер кассы", typeof(Int32) },
                { "сумма в валюте кассы (если имеется)", typeof(Double) },
            });

            var filename = PickFileToImport();
            if (!string.IsNullOrEmpty(filename))
            {
                try
                {
                    var result = xls.Load(filename);
                    MessageBox.Show("Импорт завершен успешно", "Результат", MessageBoxButtons.OK, MessageBoxIcon.Information);
                }
                catch (Exception ex)
                {
                    MessageBox.Show(ex.Message, "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
            }
        }
    }
}
