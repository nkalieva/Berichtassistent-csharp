using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Microsoft.Office.Interop.Excel;
using System.Runtime.InteropServices;
using System.Linq;
using ExcelApp = Microsoft.Office.Interop.Excel.Application;
using WinFormsApp = System.Windows.Forms.Application;

namespace Resave
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }
        private string excelComments = "";

        private void btnShow_Click(object sender, EventArgs e)
        {
            OpenFileDialog dlg = new OpenFileDialog();
            dlg.Title = "Datei auswählen";
            dlg.Filter = "Excel Files (*.xlsx)|*.xlsx|All Files (*.*)|*.*";

            if (dlg.ShowDialog() == DialogResult.OK)
            {
                // Проверяем, что это Excel файл
                if (dlg.FileName.EndsWith(".xlsx"))
                {
                    btnShow.Text = dlg.FileName;

                    // Читаем комментарии из Excel файла
                    ReadExcelComments(dlg.FileName);

                    // Активируем кнопку "Speichern"
                    btnSave.Enabled = true;
                }
                else
                {
                    MessageBox.Show("Bitte wählen Sie eine Excel-Datei aus.", "Ungültige Datei", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                }
            }
        }
        private void ReadExcelComments(string filePath)
        {
            ExcelApp excelApp = new ExcelApp();
            Workbook workbook = null;
            bool foundComments = false; // Флаг, указывающий, были ли найдены комментарии

            try
            {
                // Открываем Excel файл
                workbook = excelApp.Workbooks.Open(filePath);

                // Проходим по всем листам в книге
                foreach (Worksheet worksheet in workbook.Sheets)
                {
                    // Читаем все комментарии на текущем листе
                    foreach (Range cell in worksheet.UsedRange)
                    {
                        if (cell.Comment != null)  // Если есть комментарий в ячейке
                        {
                            foundComments = true; // Устанавливаем флаг, что комментарий найден
                            break; // Прерываем внутренний цикл, так как нам достаточно первого комментария
                        }
                    }

                    if (foundComments) // Если хотя бы один комментарий найден
                    {
                        break; // Прерываем внешний цикл, так как нам не нужно искать дальше
                    }
                }

                // Показываем сообщение с результатами чтения комментариев
                if (foundComments)
                {
                    MessageBox.Show("Kommentare gefunden!", "Gefundene Mandant Kommentare", MessageBoxButtons.OK, MessageBoxIcon.Information);
                }
                else
                {
                    MessageBox.Show("Keine Kommentare gefunden.", "Keine Kommentare", MessageBoxButtons.OK, MessageBoxIcon.Information);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("Fehler beim Auslesen der Excel-Datei: " + ex.Message, "Fehler", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            finally
            {
                // Закрываем Excel
                workbook?.Close(false);
                Marshal.ReleaseComObject(workbook);
                Marshal.ReleaseComObject(excelApp);
            }
        }


        private void btnSave_Click(object sender, EventArgs e)
        {
            SaveFileDialog saveFileDialog1 = new SaveFileDialog();
            saveFileDialog1.Filter = "Excel Files (*.xlsx)|*.xlsx";
            saveFileDialog1.Title = "Notizen speichern";

            if (saveFileDialog1.ShowDialog() == DialogResult.OK)
            {
                string savePath = saveFileDialog1.FileName;

                try
                {
                    if (savePath.EndsWith(".xlsx"))
                    {
                        ExcelApp excelApp = new ExcelApp();
                        excelApp.DisplayAlerts = false; // Отключаем уведомления Excel
                        Workbook sourceWorkbook = excelApp.Workbooks.Open(btnShow.Text); // Открываем исходный Excel файл
                        Worksheet sourceWorksheet = sourceWorkbook.Sheets[1];

                        // Создаем новый Excel файл для сохранения значений
                        Workbook newWorkbook = excelApp.Workbooks.Add();
                        Worksheet newWorksheet = newWorkbook.Sheets[1];

                        // Переменная для отслеживания найденных значений
                        bool valuesFound = false;

                        // Проходим по всем ячейкам с комментариями в исходном файле
                        foreach (Range cell in sourceWorksheet.UsedRange)
                        {
                            if (cell.Comment != null)
                            {
                                string commentText = cell.Comment.Text();
                                string mandantValue = ExtractValue(commentText, "Mandant :");
                                string jahrValue = ExtractValue(commentText, "Jahr:");

                                // Записываем значение Mandant, если оно найдено
                                if (!string.IsNullOrEmpty(mandantValue))
                                {
                                    Range newCellMandant = newWorksheet.Range[cell.Address];
                                    newCellMandant.Value = mandantValue;
                                    valuesFound = true;
                                }

                                // Записываем значение Jahr, если оно найдено
                                if (!string.IsNullOrEmpty(jahrValue))
                                {
                                    // В ячейке ниже текущей записываем Jahr (например, $F$18 -> $F$19)
                                    Range newCellJahr = newWorksheet.Range[cell.Offset[1, 0].Address];
                                    newCellJahr.Value = jahrValue;
                                    valuesFound = true;
                                }
                            }
                        }

                        // Если значения не найдены, показываем сообщение
                        if (!valuesFound)
                        {
                            MessageBox.Show("Keine Mandant- oder Jahr-Werte gefunden.", "Information", MessageBoxButtons.OK, MessageBoxIcon.Information);
                        }

                        // Сохраняем новый Excel файл
                        newWorkbook.SaveAs(savePath);
                        newWorkbook.Close(false);
                        sourceWorkbook.Close(false);

                        // Освобождаем COM-объекты
                        Marshal.ReleaseComObject(sourceWorksheet);
                        Marshal.ReleaseComObject(sourceWorkbook);
                        Marshal.ReleaseComObject(newWorksheet);
                        Marshal.ReleaseComObject(newWorkbook);
                        Marshal.ReleaseComObject(excelApp);

                        MessageBox.Show("Werte erfolgreich gespeichert!", "Erfolg", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    }
                    else
                    {
                        MessageBox.Show("Bitte wählen Sie eine Excel-Datei zum Speichern aus.", "Ungültige Datei", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    }
                }
                catch (Exception ex)
                {
                    MessageBox.Show("Fehler beim Speichern der Datei: " + ex.Message, "Fehler", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }

                btnSave.Enabled = false;
                btnShow.Text = "Excel Datei öffnen";
            }
        }

        // Вспомогательная функция для извлечения значения по ключевому слову
        private string ExtractValue(string commentText, string keyword)
        {
            int startIndex = commentText.IndexOf(keyword);
            if (startIndex >= 0)
            {
                startIndex += keyword.Length;
                int endIndex = commentText.IndexOf("\n", startIndex); // Ищем конец строки или комментария
                if (endIndex < 0) endIndex = commentText.Length;
                string value = commentText.Substring(startIndex, endIndex - startIndex).Trim();
                return value;
            }
            return string.Empty;
        }




        private void umwandeln_Click(object sender, EventArgs e)
        {
            FolderBrowserDialog dlg = new FolderBrowserDialog();
            dlg.ShowDialog();
            btnShow.Text = dlg.SelectedPath;
        }

        private void Form1_Load(object sender, EventArgs e)
        {

        }

    }
}
