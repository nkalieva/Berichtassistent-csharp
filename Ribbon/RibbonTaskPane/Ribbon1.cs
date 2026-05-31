using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Xml.Linq;
using Excel = Microsoft.Office.Interop.Excel;
using Office = Microsoft.Office.Core;
using Microsoft.Office.Tools.Excel;
using System.Windows.Forms;
using Microsoft.Office.Tools;
using Microsoft.Office.Tools.Ribbon;
using Microsoft.Office.Interop.Excel;
using System.Collections;
using System.Security.Policy;
using System.Drawing;
using System.Web.Script.Serialization;
using DocumentFormat.OpenXml.Drawing.Diagrams;
using DocumentFormat.OpenXml.Wordprocessing;
using DocumentFormat.OpenXml.Bibliography;
using Newtonsoft.Json;
using System.Net;
using System.Threading.Tasks;
using System.Net.Http;
using System.Diagnostics;
using Color = System.Drawing.Color;
using static Berichtsassistent2.Ribbon1;
using Worksheet = Microsoft.Office.Interop.Excel.Worksheet;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Drawing;
using DocumentFormat.OpenXml.Spreadsheet;
using System.Security.Cryptography;
using Microsoft.Win32;
using RibbonTaskPane;
using System.Threading;

namespace Berichtsassistent2
{

    public partial class Ribbon1
    {

        private CustomTaskPane _taskPane;
        private Excel.Application Application;
        private bool isSelectionBoxOpen = false;
        private object startRow;
        private object startColumn;
        private object endRow;
        private object endColumn;
        private object host;

        public object Mandanten_ID { get; private set; }
        public object Jahr_ID { get; private set; }
        public object UStruktur_ID { get; private set; }
        public object Datenart_ID { get; private set; }
        public static int AnzahlZeilen { get; private set; }
        public static int AnzahlSpalten { get; private set; }
        public Worksheet YourExcelWorksheet { get; private set; }
        public string Mandanten { get; private set; }
        public string Jahr { get; private set; }
        public string UStruktur { get; private set; }
        public string Datenart { get; private set; }

        private void BA2Btn_Click(object sender, RibbonControlEventArgs e)
        {
            ShowTaskPane();
        }

        private void ShowTaskPane()
        {
            Excel.Window activeWindow = Globals.ThisAddIn.Application.ActiveWindow;
            if (_taskPane != null && !_taskPane.Visible)
            {
                _taskPane.Visible = true;
                activeWindow.Activate();
            }
            else if (_taskPane != null && _taskPane.Visible)
            {
                _taskPane.Visible = false;
                activeWindow.Activate();
            }

        }
        private void Ribbon1_Load_1(object sender, RibbonUIEventArgs e)
        {
            Application = Globals.ThisAddIn.Application;
            MyUserControl myUserControl = new MyUserControl(); // Ersetze "MyUserControl" durch den Namen deines benutzerdefinierten MyUserControls
            _taskPane = Globals.ThisAddIn.CustomTaskPanes.Add(myUserControl, "API Einstellung");
            _taskPane.DockPosition = Office.MsoCTPDockPosition.msoCTPDockPositionRight;
            // _taskPane.DockPosition = Microsoft.Office.Core.MsoCTPDockPosition.msoCTPDockPositionRight;
            _taskPane.Width = 430;
            _taskPane.Visible = false;
        }

        private void MandantenAuswahl(object sender, RibbonControlEventArgs e)
        {
            Excel.Application excelApp = Globals.ThisAddIn.Application;
            Excel.Workbook workbook = excelApp.ActiveWorkbook; // Aktuelle Arbeitsmappe erhalten
            Excel.Worksheet worksheet = workbook.ActiveSheet;

            Excel.Range currentCell = excelApp.ActiveCell;

            // Dropdown-Liste für erste Liste erstellen
            Excel.Range formulaCell6 = currentCell.Offset[0, 6];
            Excel.Range dropdownListRange1 = currentCell.Offset[1, 2].Resize[2, 1]; 
            Excel.Range dropdownAuswahlMandanten = formulaCell6.Resize[40];
            Excel.Validation validation1 = currentCell.Offset[1, 1].Validation;
            validation1.Delete();
            validation1.Add(
                Excel.XlDVType.xlValidateList,
                Formula1: "=" + dropdownAuswahlMandanten.Address
            );
            validation1.IgnoreBlank = true;
            validation1.InCellDropdown = true;

            // Dropdown-Liste für zweite Liste erstellen
            Excel.Range formulaCell7 = currentCell.Offset[0, 4];
            Excel.Range dropdownAuswahlJahr = formulaCell7.Resize[20];
            Excel.Validation validation2 = currentCell.Offset[2, 1].Validation;
            validation2.Delete();
            validation2.Add(
                Excel.XlDVType.xlValidateList,
                Formula1: "=" + dropdownAuswahlJahr.Address
            );
            validation2.IgnoreBlank = true;
            validation2.InCellDropdown = true;

            //dropdown Ustruktur
            Excel.Range formulaCell8 = currentCell.Offset[0, 8];
            Excel.Range dropdownAuswahlUStruktur = formulaCell8.Resize[40];
            Excel.Validation validation3 = currentCell.Offset[3, 1].Validation;
            validation3.Delete();
            validation3.Add(
                Excel.XlDVType.xlValidateList,
                Formula1: "=" + dropdownAuswahlUStruktur.Address
            );
            validation3.IgnoreBlank = true;
            validation3.InCellDropdown = true;

            //dropdown Datenart

            Excel.Validation validation4 = currentCell.Offset[4, 1].Validation;
            validation4.Delete();
            validation4.Add(
                Excel.XlDVType.xlValidateList,
                Excel.XlDVAlertStyle.xlValidAlertStop,
                Excel.XlFormatConditionOperator.xlBetween,
                "Plan; Ist; Vorschau",
                Type.Missing
            );

            validation4.IgnoreBlank = true;
            validation4.InCellDropdown = true;

            Excel.Range daten = currentCell.Offset[0, 1];
            Excel.Range auswahlmandanten = currentCell.Offset[1, 1];//Zelle für die erste Formel
                                                                    //Excel.Range formulaCell2 = currentCell.Offset[1, 1]; // Zelle für die zweite Formel
            Excel.Range auswahljahr = currentCell.Offset[2, 1];
            Excel.Range mandanten = currentCell.Offset[1, 0];
            Excel.Range jahre = currentCell.Offset[2, 0];
            Excel.Range ustruktur = currentCell.Offset[3, 0];
            Excel.Range datenart = currentCell.Offset[4, 0];
            Excel.Range auswahlustruktur = currentCell.Offset[3, 1];
            Excel.Range auswahldatenart = currentCell.Offset[4, 1];

            //int rowNumber = currentCell.Offset[1, 1].Row;
            string auswahlMandantenAddress = currentCell.Offset[1, 1].Address;
            string auswahlJahrAddress = currentCell.Offset[2, 1].Address;

            daten.Value = "Daten:";
            auswahlmandanten.Value = "Auswahl Mandanten";
            auswahljahr.Value = "Auswahl Jahre";
            mandanten.Value = "Mandanten";
            jahre.Value = "Jahre";
            // Set the array formula
            // Select the cell and simulate pressing Enter to confirm the formula
            
            //RemoveAtSymbol(formulaCell6);
            //formulaCell7.Value = auswahlmandanten.Value;
            formulaCell6.FormulaArray = "=Liste_Stamm_Mandanten()";
            formulaCell7.FormulaArray = "=Liste_Stamm_Geschaeftsjahre_Zusammen(" + auswahlMandantenAddress + ")";
            formulaCell8.FormulaArray = "=Stamm_UStruktur(" + auswahlMandantenAddress + "," + auswahlJahrAddress + ")";
           

            // formulaCell8.Value = auswahlustruktur.Value;

            ustruktur.Value = "UStruktur";
            datenart.Value = "Datenart";
            auswahlustruktur.Value = "Auswahl UStruktur";
            auswahldatenart.Value = "Auswahl Datenart";

            // formulaCell2.Formula = "=BenutzerdefinierteFunktion(" + dropdownListRange1.Address + ")";
            currentCell.Offset[1, 0].Interior.Color = Excel.XlRgbColor.rgbMoccasin; // Hintergrundfarbe für "Auswahl Mandanten" 
            currentCell.Offset[2, 0].Interior.Color = Excel.XlRgbColor.rgbMoccasin; // Hintergrundfarbe für "Auswahl Jahre" 
            currentCell.Offset[0, 1].Interior.Color = Excel.XlRgbColor.rgbOrange;
            currentCell.Offset[1, 1].Interior.Color = Excel.XlRgbColor.rgbSilver;
            currentCell.Offset[2, 1].Interior.Color = Excel.XlRgbColor.rgbSilver;
            currentCell.Offset[0, 0].Interior.Color = Excel.XlRgbColor.rgbGray;
            currentCell.Offset[3, 0].Interior.Color = Excel.XlRgbColor.rgbMoccasin;
            currentCell.Offset[4, 0].Interior.Color = Excel.XlRgbColor.rgbMoccasin;
            currentCell.Offset[3, 1].Interior.Color = Excel.XlRgbColor.rgbSilver;
            currentCell.Offset[4, 1].Interior.Color = Excel.XlRgbColor.rgbSilver;
            ConfirmFormula(formulaCell6);
            //System.Threading.Thread.Sleep(100);
            ConfirmFormula(formulaCell7);
            //System.Threading.Thread.Sleep(100);
            ConfirmFormula(formulaCell8);

            GC.Collect();
        }

        // Function to set the focus on a cell and press F2 and Enter
        void ConfirmFormula(Excel.Range cell)
        {
            Excel.Application excelApp = Globals.ThisAddIn.Application;
            cell.Select();
            //System.Threading.Thread.Sleep(100);
            excelApp.SendKeys("{F2}");
            //System.Threading.Thread.Sleep(100);
            excelApp.SendKeys("{ENTER}");
            System.Windows.Forms.Application.DoEvents();
        }
        private void RemoveAtSymbol(Excel.Range cell)
        {
            if (cell.Formula.Contains("@"))
            {
                cell.Formula = cell.Formula.Replace("@", "");
            }
        }

        // private RegistryKey key;
        private const string RegistrySubKey = "BA2";
        private const string RegistryValueServer = "BA2_Server";
        private const string RegistryValuePort = "BA2_Port";
        private static string Mandanten_Name;

        private static string ReadRegistry()
        {
            string server = string.Empty;
            string port = string.Empty;

            using (var key = Registry.CurrentUser.OpenSubKey(RegistrySubKey))
            {
                if (key != null)
                {
                    server = key.GetValue(RegistryValueServer)?.ToString();
                    port = key.GetValue(RegistryValuePort)?.ToString();
                }
            }

            return server + ":" + port;
        }

        private void SaveDataToExcel(string option, string targetCell)
        {
            Excel.Application excelApp = new Excel.Application();
            Excel.Workbook workbook = excelApp.Workbooks.Add();
            Excel.Worksheet worksheet = workbook.ActiveSheet;

            worksheet.Range[targetCell].Value = option;
            workbook.SaveAs("C:\\Pfad\\zur\\Excel-Datei.xlsx");

            workbook.Close();
            excelApp.Quit();
        }

        private static object Liste_Daten_GuV(string host, int Mandanten_ID, int Jahr_ID, int UStruktur_ID, int Datenart_ID)

        {
            object[,] ergebnis;
            try
            {
                string ur1 = $"http://{host}/api/data/income/" + Mandanten_ID + "/" + Jahr_ID + "/" + UStruktur_ID + "/" + Datenart_ID;
                var awaiter = CallURL(ur1);

                var myData = JsonConvert.DeserializeObject<List<GuV_Liste>>(awaiter.Result);

                ergebnis = new object[myData.Count + 1, 15];

                ergebnis[0, 0] = "Reihenfolge";
                ergebnis[0, 1] = "Name";
                ergebnis[0, 2] = "Monat1";
                ergebnis[0, 3] = "Monat2";
                ergebnis[0, 4] = "Monat3";
                ergebnis[0, 5] = "Monat4";
                ergebnis[0, 6] = "Monat5";
                ergebnis[0, 7] = "Monat6";
                ergebnis[0, 8] = "Monat7";
                ergebnis[0, 9] = "Monat8";
                ergebnis[0, 10] = "Monat9";
                ergebnis[0, 11] = "Monat10";
                ergebnis[0, 12] = "Monat11";
                ergebnis[0, 13] = "Monat12";
                ergebnis[0, 14] = "Jahreswert";

                {

                    for (int i = 0; i < myData.Count; i++)
                    {
                        ergebnis[i + 1, 0] = myData[i].Reihenfolge;
                        ergebnis[i + 1, 1] = myData[i].Position_Name;
                        ergebnis[i + 1, 2] = myData[i].Monat1;
                        ergebnis[i + 1, 3] = myData[i].Monat2;
                        ergebnis[i + 1, 4] = myData[i].Monat3;
                        ergebnis[i + 1, 5] = myData[i].Monat4;
                        ergebnis[i + 1, 6] = myData[i].Monat5;
                        ergebnis[i + 1, 7] = myData[i].Monat6;
                        ergebnis[i + 1, 8] = myData[i].Monat7;
                        ergebnis[i + 1, 9] = myData[i].Monat8;
                        ergebnis[i + 1, 10] = myData[i].Monat9;
                        ergebnis[i + 1, 11] = myData[i].Monat10;
                        ergebnis[i + 1, 12] = myData[i].Monat11;
                        ergebnis[i + 1, 13] = myData[i].Monat12;
                        ergebnis[i + 1, 14] = myData[i].Jahreswert;
                    }
                }

            }
            catch (Exception ex)
            {
                ergebnis = new object[1, 2];
                ergebnis[0, 0] = "Fehler:";
                Type exceptionType = ex.GetType();
                ergebnis[0, 1] = "Parameter nicht erkannt. Bitte geben Sie einen gültigen Servernamen, Port und Mandanten-ID ein. ";
                //ergebnis[0, 2] = "Bitte stellen Sie sicher, dass die Datenbankverbindung hergestellt ist und versuchen Sie es erneut: " + exceptionType.Namespace;
            }
            return ergebnis;
        }
        public class GuV_Liste
        {
            public long MandantenID { get; set; }
            public int DatenartID { get; set; }
            public int Pos_GuV_ID { get; set; }
            public int U_Struktur_ID { get; set; }
            public int Jahr_ID { get; set; }
            public int Planungsart_ID { get; set; }
            public int Reihenfolge { get; set; }
            public string Position_Name { get; set; }
            public string Position_Name_Lang { get; set; }
            public float Monat1 { get; set; }
            public float Monat2 { get; set; }
            public float Monat3 { get; set; }
            public float Monat4 { get; set; }
            public float Monat5 { get; set; }
            public float Monat6 { get; set; }
            public float Monat7 { get; set; }
            public float Monat8 { get; set; }
            public float Monat9 { get; set; }
            public float Monat10 { get; set; }
            public float Monat11 { get; set; }
            public float Monat12 { get; set; }
            public float Jahreswert { get; set; }
            public int ZL_Schrift_Groesse { get; set; }
            public string ZL_Schrift_Name { get; set; }
            public int ZL_Schrift_Style { get; set; }
            public int ZL_Textfarbe { get; set; }
            public int ZL_Transparenz { get; set; }
            public int ZL_Zeilenfarbe { get; set; }
        }

        private static object Liste_Daten_Passiva(string host, int Mandanten_ID, int Jahr_ID, int UStruktur_ID, int Datenart_ID)
        {
            object[,] ergebnis;
            try
            {
                string ur1 = $"http://{host}/api/data/balance/" + Mandanten_ID + "/" + Jahr_ID + "/" + UStruktur_ID + "/" + Datenart_ID;
                var awaiter = CallURL(ur1);

                var myData = JsonConvert.DeserializeObject<List<Passiva>>(awaiter.Result);

                ergebnis = new object[myData.Count + 1, 8];

                ergebnis[0, 0] = "ID";
                ergebnis[0, 1] = "Reihenfolge";
                ergebnis[0, 2] = "Position_Name";
                ergebnis[0, 3] = "Anfangsbestand";
                ergebnis[0, 4] = "Monat_Z";
                ergebnis[0, 5] = "Monat_A";
                ergebnis[0, 6] = "Monat_S";
                ergebnis[0, 7] = "Endbestand";

                for (int i = 0; i < myData.Count; i++)
                {
                    {
                        ergebnis[i + 1, 0] = myData[i].Pos_Bilanz_ID;
                        ergebnis[i + 1, 1] = myData[i].Reihenfolge;
                        ergebnis[i + 1, 2] = myData[i].Position_Name;
                        ergebnis[i + 1, 3] = (myData[i].Anfangsbestand);
                        ergebnis[i + 1, 4] = (myData[i].Monat1_Z + myData[i].Monat2_Z + myData[i].Monat3_Z + myData[i].Monat4_Z + myData[i].Monat5_Z + myData[i].Monat6_Z + myData[i].Monat7_Z + myData[i].Monat8_Z + myData[i].Monat9_Z + myData[i].Monat10_Z + myData[i].Monat11_Z + myData[i].Monat12_Z);
                        ergebnis[i + 1, 5] = (myData[i].Monat1_A + myData[i].Monat2_A + myData[i].Monat3_A + myData[i].Monat4_A + myData[i].Monat5_A + myData[i].Monat6_A + myData[i].Monat7_A + myData[i].Monat8_A + myData[i].Monat9_A + myData[i].Monat10_A + myData[i].Monat11_A + myData[i].Monat12_A);
                        ergebnis[i + 1, 6] = (myData[i].Monat1_S + myData[i].Monat2_S + myData[i].Monat3_S + myData[i].Monat4_S + myData[i].Monat5_S + myData[i].Monat6_S + myData[i].Monat7_S + myData[i].Monat8_S + myData[i].Monat9_S + myData[i].Monat10_S + myData[i].Monat11_S + myData[i].Monat12_S);
                        ergebnis[i + 1, 7] = (myData[i].Endbestand);
                    }
                }
            }
            catch (Exception ex)
            {
                ergebnis = new object[1, 2];
                ergebnis[0, 0] = "Fehler:";
                Type exceptionType = ex.GetType();
                ergebnis[0, 1] = "Parameter nicht erkannt. Bitte geben Sie einen gültigen Servernamen, Port und Mandanten-ID ein. ";
                //ergebnis[0, 2] = "Bitte stellen Sie sicher, dass die Datenbankverbindung hergestellt ist und versuchen Sie es erneut: " + exceptionType.Namespace;
            }
            return ergebnis;
        }

        public class Passiva
        {
            public long MandantID { get; set; }
            public int DatenartID { get; set; }
            public string Bilanzseite { get; set; }
            public int Pos_Bilanz_ID { get; set; }
            public int UStruktur_ID { get; set; }
            public int Jahr_ID { get; set; }
            public int Planungsart_ID { get; set; }
            public int Reihenfolge { get; set; }
            public string Position_Name { get; set; }
            public string Position_Name_Lang { get; set; }
            public float Anfangsbestand { get; set; }
            public float Monat1_Z { get; set; }
            public float Monat2_Z { get; set; }
            public float Monat3_Z { get; set; }
            public float Monat4_Z { get; set; }
            public float Monat5_Z { get; set; }
            public float Monat6_Z { get; set; }
            public float Monat7_Z { get; set; }
            public float Monat8_Z { get; set; }
            public float Monat9_Z { get; set; }
            public float Monat10_Z { get; set; }
            public float Monat11_Z { get; set; }
            public float Monat12_Z { get; set; }
            public float Monat1_A { get; set; }
            public float Monat2_A { get; set; }
            public float Monat3_A { get; set; }
            public float Monat4_A { get; set; }
            public float Monat5_A { get; set; }
            public float Monat6_A { get; set; }
            public float Monat7_A { get; set; }
            public float Monat8_A { get; set; }
            public float Monat9_A { get; set; }
            public float Monat10_A { get; set; }
            public float Monat11_A { get; set; }
            public float Monat12_A { get; set; }
            public float Monat1_S { get; set; }
            public float Monat2_S { get; set; }
            public float Monat3_S { get; set; }
            public float Monat4_S { get; set; }
            public float Monat5_S { get; set; }
            public float Monat6_S { get; set; }
            public float Monat7_S { get; set; }
            public float Monat8_S { get; set; }
            public float Monat9_S { get; set; }
            public float Monat10_S { get; set; }
            public float Monat11_S { get; set; }
            public float Monat12_S { get; set; }
            public float Monat1_B { get; set; }
            public float Monat2_B { get; set; }
            public float Monat3_B { get; set; }
            public float Monat4_B { get; set; }
            public float Monat5_B { get; set; }
            public float Monat6_B { get; set; }
            public float Monat7_B { get; set; }
            public float Monat8_B { get; set; }
            public float Monat9_B { get; set; }
            public float Monat10_B { get; set; }
            public float Monat11_B { get; set; }
            public float Monat12_B { get; set; }
            public float Endbestand { get; set; }
            public int ZL_Schrift_Groesse { get; set; }
            public string ZL_Schrift_Name { get; set; }
            public int ZL_Schrift_Style { get; set; }
            public int ZL_Textfarbe { get; set; }
            public int ZL_Transparenz { get; set; }
            public int ZL_Zeilenfarbe { get; set; }
        }
        public static async Task<string> CallURL(string url)
        {
            HttpClient client = new HttpClient();
            ServicePointManager.SecurityProtocol = SecurityProtocolType.Tls;
            client.DefaultRequestHeaders.Accept.Clear();
            var response = client.GetStringAsync(url);
            return await response;
        }
        private static string GetMandantenIDFromName(string Mandanten_Name)
        {
            int Mandanten_ID = 0;
            if (!string.IsNullOrEmpty(Mandanten_Name))
            {
                string mandantenIdNurZahlen = new string(Mandanten_Name.TakeWhile(char.IsDigit).ToArray());
                Int32.TryParse(mandantenIdNurZahlen, out Mandanten_ID);
            }
            return Mandanten_ID.ToString();
        }

        private static int GetUStrukturIDFromName(string UStruktur_Name)
        {
            int UStruktur_ID = 0;
            if (!string.IsNullOrEmpty(UStruktur_Name))
            {
                string ustrukturIdNurZahlen = new string(UStruktur_Name.TakeWhile(char.IsDigit).ToArray());
                Int32.TryParse(ustrukturIdNurZahlen, out UStruktur_ID);
            }
            return UStruktur_ID;
        }
        private void Layout(object sender, RibbonControlEventArgs e)
        {

            string server = "localhost";  // Замените на фактическое значение сервера
            string port = "8080";  // Замените на фактическое значение порта

            if (string.IsNullOrEmpty(server) || string.IsNullOrEmpty(port))
            {
                MessageBox.Show("Server oder Port", "Fehler", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }

            Excel.Application excelApp = Globals.ThisAddIn.Application;
            Excel.Workbook workbook = excelApp.ActiveWorkbook;
            Excel.Worksheet worksheet = workbook.ActiveSheet;
            Excel.Range currentCell = excelApp.ActiveCell;

            excelApp.ScreenUpdating = false;
            excelApp.Calculation = Excel.XlCalculation.xlCalculationManual;
            excelApp.DisplayAlerts = false;

            try
            {
                ProcessRange(worksheet.UsedRange);
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Server oder Port ist nicht angegeben. Bitte überprüfen Sie die Einstellungen", "Fehler", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            finally
            {
                excelApp.ScreenUpdating = true;
                excelApp.Calculation = Excel.XlCalculation.xlCalculationAutomatic;
                excelApp.DisplayAlerts = true;

                worksheet.Calculate();
            }
        }

        private void ProcessRange(Excel.Range range)
        {
            Excel.Application excelApp = Globals.ThisAddIn.Application;
            Excel.Workbook workbook = excelApp.ActiveWorkbook;
            Excel.Worksheet worksheet = workbook.ActiveSheet;

            foreach (Excel.Range cell in range)
            {
                if (cell.HasFormula)
                {
                    string formula = cell.Formula;

                    if (formula.StartsWith("=Liste_Daten_GuV"))
                    {
                        string[] formulaParts = formula.Split(',');
                        if (formulaParts.Length >= 4)
                        {
                            string Mandanten_ID_Name = ExtractValue(formulaParts[0], 17);
                            string Jahr_ID = ExtractValue(formulaParts[1]);
                            string UStruktur_ID_Name = ExtractValue(formulaParts[2]);
                            string selectedDatenartID = ExtractValue(formulaParts[3], 0, true);

                            FormatFormulaA(cell, Mandanten_ID_Name, Jahr_ID, UStruktur_ID_Name, selectedDatenartID);
                        }
                    }
                    else if (formula.StartsWith("=Liste_Daten_Passiva"))
                    {
                        string[] formulaParts = formula.Split(',');
                        if (formulaParts.Length >= 4)
                        {
                            string Mandanten_ID = ExtractValue(formulaParts[0], 21);
                            string Jahr_ID = ExtractValue(formulaParts[1]);
                            string UStruktur_ID = ExtractValue(formulaParts[2]);
                            string Datenart_ID = ExtractValue(formulaParts[3], 0, true);

                            FormatFormulaB(cell, Mandanten_ID, Jahr_ID, UStruktur_ID, Datenart_ID);
                        }
                    }
                    else
                    {
                        // Wenn weder FormelA noch FormelB in der Zelle sind, Zellinhalt löschen
                        // cell.ClearContents();
                    }
                }
            }
        }

        private string ExtractValue(string input, int startIndex = 0, bool trimEnd = false)
        {
            Excel.Application excelApp = Globals.ThisAddIn.Application;
            Excel.Workbook workbook = excelApp.ActiveWorkbook;
            Excel.Worksheet worksheet = workbook.ActiveSheet;

            string value = input.Substring(startIndex, input.Length - startIndex);
            if (trimEnd) value = value.TrimEnd(')');

            bool isNumber = int.TryParse(value, out int _);

            if (isNumber)
            {
                return value;
            }
            else
            {
                int[] coordinates = ConvertCellToCoordinates(value);
                int exclamationIndex = value.IndexOf('!');
                if (exclamationIndex != -1)
                {
                    string[] parts = value.Split('!');
                    Excel.Worksheet worksheet2 = workbook.Worksheets[parts[0]];
                    object cellValue = worksheet2.Cells[coordinates[0], coordinates[1]].Value2;
                    return cellValue.ToString();
                }
                else
                {
                    object cellValue = worksheet.Cells[coordinates[0], coordinates[1]].Value2;
                    return cellValue.ToString();
                }
            }
        }
        private void FormatFormulaA(Excel.Range cell, string Mandanten_ID_Name, string Jahr_ID, string UStruktur_ID_Name, string Datenart_ID)
        {
            Mandanten_ID = GetMandantenIDFromName(Mandanten_ID_Name);
            UStruktur_ID = GetUStrukturIDFromName(UStruktur_ID_Name);
            int selectedDatenartID;
            
            if (!string.IsNullOrEmpty(Datenart_ID))
            {
                switch (Datenart_ID.ToLower())
                {
                    case "plan":
                        selectedDatenartID = 2;
                        break;
                    case "ist":
                        selectedDatenartID = 3;
                        break;
                    case "vorschau":
                        selectedDatenartID = 4;
                        break;
                    default:
                        if (int.TryParse(Datenart_ID, out int parsedID))
                        {
                            selectedDatenartID = parsedID;
                        }
                        else
                        {
                            // Hier entscheiden, was passieren soll, wenn keine gültige Zahl oder Text-ID übergeben wurde
                            throw new ArgumentException("Ungültige Datenart_ID.");
                        }
                        break;
                }

                // Hier wird die Methode Liste_Daten_GuV aufgerufen und das Ergebnis zurückgegeben 
            }
            else
            {
                throw new ArgumentNullException("Datenart_ID", "Datenart_ID darf nicht leer sein.");
            }

            string host = ReadRegistry();

            string ur1 = $"http://{host}/api/data/income/" + Mandanten_ID_Name + "/" + Jahr_ID + "/" + UStruktur_ID_Name + "/" + selectedDatenartID;

            var awaiter = CallURL(ur1);

            var myData = JsonConvert.DeserializeObject<List<GuV_Liste>>(awaiter.Result);

            //for (int spalte = 0; spalte < 16; spalte++)
            {
                for (int zeile = 0; zeile < myData.Count; zeile++)
                {
                    Excel.Range currentCell = cell.Offset[zeile + 1, 0].Resize[1, 15];

                    currentCell.NumberFormat = "#,##";

                    currentCell.Interior.Color = ConvertDelphiColorToExcelColor(myData[zeile].ZL_Zeilenfarbe);
                    currentCell.Font.Size = myData[zeile].ZL_Schrift_Groesse;
                    currentCell.Font.Name = myData[zeile].ZL_Schrift_Name;

                    if ((myData[zeile].ZL_Schrift_Style & (1 << 0)) != 0)
                    {
                        currentCell.Font.Bold = true;
                    }
                    else
                    {
                        currentCell.Font.Bold = false;
                    }

                    if ((myData[zeile].ZL_Schrift_Style & (1 << 1)) != 0)
                    {
                        currentCell.Font.Italic = true;
                    }
                    else
                    {
                        currentCell.Font.Italic = false;
                    }

                    if ((myData[zeile].ZL_Schrift_Style & (1 << 2)) != 0)
                    {
                        currentCell.Font.Underline = true;
                    }
                    else
                    {
                        currentCell.Font.Underline = false;
                    }

                    if ((myData[zeile].ZL_Schrift_Style & (1 << 3)) != 0)
                    {
                        currentCell.Font.Strikethrough = true;
                    }
                    else
                    {
                        currentCell.Font.Strikethrough = false;
                    }
                    // currentCell.Font.FontStyle = myData[zeile].ZL_Schrift_Style;
                    currentCell.Font.Color = ConvertDelphiColorToExcelColor(myData[zeile].ZL_Textfarbe);

                    //currentCell.Font.TintAndShade = myData[zeile].ZL_Transparenz;
                }
            }

        }
        private void FormatFormulaB(Excel.Range cell, string Mandanten_ID, string Jahr_ID, string UStruktur_ID, string Datenart_ID)
        {
            Mandanten_ID = GetMandantenIDFromName(Mandanten_ID);
            string host = ReadRegistry();

            string ur1 = $"http://{host}/api/data/balance/" + Mandanten_ID + "/" + Jahr_ID + "/" + UStruktur_ID + "/" + Datenart_ID;


            var awaiter = CallURL(ur1);

            var myData = JsonConvert.DeserializeObject<List<Passiva>>(awaiter.Result);

            //for (int spalte = 0; spalte < 16; spalte++)
            {
                for (int zeile = 0; zeile < myData.Count; zeile++)
                {
                    Excel.Range currentCell = cell.Offset[zeile + 1, 0].Resize[1, 8];
                    currentCell.NumberFormat = "#,##";
                    currentCell.Interior.Color = ConvertDelphiColorToExcelColor(myData[zeile].ZL_Zeilenfarbe);
                    currentCell.Font.Size = myData[zeile].ZL_Schrift_Groesse;
                    currentCell.Font.Name = myData[zeile].ZL_Schrift_Name;

                    if ((myData[zeile].ZL_Schrift_Style & (1 << 0)) != 0)
                    {
                        currentCell.Font.Bold = true;
                    }
                    else
                    {
                        currentCell.Font.Bold = false;
                    }

                    if ((myData[zeile].ZL_Schrift_Style & (1 << 1)) != 0)
                    {
                        currentCell.Font.Italic = true;
                    }
                    else
                    {
                        currentCell.Font.Italic = false;
                    }

                    if ((myData[zeile].ZL_Schrift_Style & (1 << 2)) != 0)
                    {
                        currentCell.Font.Underline = true;
                    }
                    else
                    {
                        currentCell.Font.Underline = false;
                    }

                    if ((myData[zeile].ZL_Schrift_Style & (1 << 3)) != 0)
                    {
                        currentCell.Font.Strikethrough = true;
                    }
                    else
                    {
                        currentCell.Font.Strikethrough = false;
                    }

                    currentCell.Font.Color = ConvertDelphiColorToExcelColor(myData[zeile].ZL_Textfarbe);
                }
            }

        }

        public static string ExtractCellReference(string input)
        {

            // Prüfen, ob ein '!' im String vorhanden ist.
            int exclamationIndex = input.IndexOf('!');
            if (exclamationIndex != -1)
            {
                // Wenn ja, extrahiere alles nach dem '!' als Zellenbezug.
                return input.Substring(exclamationIndex + 1);
            }
            else
            {
                // Wenn nicht, gehe davon aus, dass der gesamte String der Zellenbezug ist.
                return input;
            }
        }
        public static int[] ConvertCellToCoordinates(string cell)
        {        // Trenne den String in Buchstaben und Zahlen

            //string columnPart = new string(cell.Where(c => char.IsLetter(c)).ToArray());
            string columnPart = new string(ExtractCellReference(cell).Where(c => char.IsLetter(c)).ToArray());
            string rowPart = new string(ExtractCellReference(cell).Where(c => char.IsDigit(c)).ToArray());

            int row = int.Parse(rowPart);

            int column = ColumnToNumber(columnPart);

            return new[] { row, column };
        }

        public static int ColumnToNumber(string column)
        {
            int number = 0;

            int multiplier = 1;

            for (int i = column.Length - 1; i >= 0; i--)

            {
                number += (column[i] - 'A' + 1) * multiplier;

                multiplier *= 26;
            }

            return number;

        }
        private void ReadAllCellsAsNumbers(Excel.Worksheet worksheet)
        {
            double cellValue = 0;
            Excel.Range allCells = worksheet.UsedRange; // Alle verwendeten Zellen im Arbeitsblatt

            foreach (Excel.Range cell in allCells)
            {
                //double cellValue = 0;

                try
                {
                    cellValue = Convert.ToDouble(cell.get_Value());
                    // Hier kannst du den Wert cellValue verwenden, z.B., weiterverarbeiten oder speichern
                    Console.WriteLine("Gelesener Wert als Zahl: " + cellValue);
                }
                catch (Exception ex)
                {
                    // Behandle Ausnahmen, die auftreten können, wenn der Wert nicht extrahiert werden kann
                    // Hier könntest du eine Fehlerbehandlung hinzufügen
                    Console.WriteLine("Fehler beim Extrahieren des Werts aus der Zelle: " + ex.Message);
                }
            }

        }

        public static int ConvertDelphiColorToExcelColor(int delphiColor)
        {

            int r = (delphiColor >> 16) & 0xFF;

            int g = (delphiColor >> 8) & 0xFF;

            int b = delphiColor & 0xFF;

            return (r << 16) | (g << 8) | b;

        }

        private void FormatFormulaB(Excel.Range cell)
        {
            for (int spalte = 0; spalte < 8; spalte++)
            {
                for (int zeile = 0; zeile < 184; zeile++)
                {
                    Excel.Range currentCell = cell.Offset[zeile, spalte];
                    currentCell.Interior.Color = Excel.XlRgbColor.rgbYellow;
                }
            }
        }
    }
}










