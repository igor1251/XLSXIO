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
using XLSXIO.Import;
using XLSXIO.AuxiliaryTypes;
using XLSXIO.Export;

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
                { "Ключ", "Key", typeof(string) },
                { "Значение", "Value", typeof(string) },
            });

            var filename = PickFileToImport();
            DataTable result = new DataTable();
            if (!string.IsNullOrEmpty(filename))
            {
                try
                {
                    result = xls.Load(filename);
                    MessageBox.Show("Импорт завершен успешно", "Результат", MessageBoxButtons.OK, MessageBoxIcon.Information);
                }
                catch (Exception ex)
                {
                    MessageBox.Show(ex.Message, "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
            }

            var xlsx = new XLSXExport(new XLSColumnTemplatesCollection
            {
                { "Ключ", "Key", typeof(string) },
                { "Значение", "Value", typeof(string) },
            });

            try
            {
                xlsx.Upload($"{filename}.export.xlsx", result);
                MessageBox.Show("Экспорт завершен успешно", "Результат", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }
    }
}
