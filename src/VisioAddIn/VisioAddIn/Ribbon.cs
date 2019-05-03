using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.IO;
using System.Windows.Forms;
using Microsoft.Office.Tools.Ribbon;

namespace VisioAddIn
{
    public partial class Ribbon
    {
        private void Ribbon_Load(object sender, RibbonUIEventArgs e)
        {
        }

        private void openFileButton_Click(object sender, RibbonControlEventArgs e)
        {
            if (openFileDialog.ShowDialog() == DialogResult.OK)
            {
                string input = File.ReadAllText(openFileDialog.FileName);
                try
                {
                    Globals.ThisAddIn.ShowGraph(input);
                }
                catch (Exception exc)
                {
                    ErrorMessage("Во время импорта графа возникла следующая ошибка:\n\n" + exc.Message, "Не удалось отобразить граф");
                    Globals.ThisAddIn.RemovePageIfError();
                }
            }
        }

        private void exportGraphButton_Click(object sender, RibbonControlEventArgs e)
        {
            if (saveFileDialog.ShowDialog() == DialogResult.OK)
            {
                try
                {
                    Globals.ThisAddIn.ExportGraph(saveFileDialog.FileName);
                }
                catch (Exception exc)
                {
                    ErrorMessage("Во время экспорта графа возникла следующая ошибка:\n\n" + exc.Message, "Не удалось экспортировать граф");
                }
            }
        }

        /// <summary>
        /// Метод для отображения сообщения об ошибке
        /// </summary>
        /// <param name="message"></param>
        /// <param name="caption"></param>
        public void ErrorMessage(string message, string caption)
        {
            MessageBoxButtons buttons = MessageBoxButtons.OK;
            DialogResult result;

            // Отобразить окошко об ошибке
            result = MessageBox.Show(message, caption, buttons);
        }
    }
}
