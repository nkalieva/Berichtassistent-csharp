using ExcelDna.Integration;
using Microsoft.CodeAnalysis;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading;
using System.Threading.Tasks;
//using Newtonsoft;
//using Newtonsoft.Json;
//using Newtonsoft.Json.Linq;
using System.Net.Http;
using System.Net;
using System.IO;
using System.Xml.Linq;
using System.Globalization;
using System.Linq.Expressions;
using System.Text.Json;
using System.Net.Sockets;
using Microsoft.Office.Interop.Excel;
using Microsoft.Win32;
using System.Text.RegularExpressions;
using Excel = Microsoft.Office.Interop.Excel;
using DocumentFormat.OpenXml.Spreadsheet;
using Microsoft.ReportingServices.ReportProcessing.ReportObjectModel;
using Microsoft.Graph.Models;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Bibliography;
using OfficeOpenXml;
using OfficeOpenXml.Style;
using System.Drawing;
using System.Runtime.InteropServices;
using Color = System.Drawing.Color;
using Font = System.Drawing.Font;
using System.Runtime.ConstrainedExecution;
using ExcelDna.Integration.CustomUI;
using System.Text;


public static class CustomFunctions
{
    // private RegistryKey key;
    private const string RegistrySubKey = "BA2";
    private const string RegistryValueServer = "BA2_Server";
    private const string RegistryValuePort = "BA2_Port";
    private static string Mandanten_Name;

    public static async Task<string> CallURL(string url)
    {
        HttpClient client = new HttpClient();
        ServicePointManager.SecurityProtocol = SecurityProtocolType.Tls;
        client.DefaultRequestHeaders.Accept.Clear();
        var response = client.GetStringAsync(url);
        return await response;
    }

    //currentCell.Offset[2, 1].Interior.Color = Excel.XlRgbColor.rgbSilver; (Farbe in button Click)
    //currentCell.Offset[0, 0].Interior.Color = Excel.XlRgbColor.rgbGray;
  
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

    private static int GetMandantenIDFromName(string Mandanten_Name)
    {
        int Mandanten_ID = 0;
        if (!string.IsNullOrEmpty(Mandanten_Name))
        {
            string mandantenIdNurZahlen = new string(Mandanten_Name.TakeWhile(char.IsDigit).ToArray());
            Int32.TryParse(mandantenIdNurZahlen, out Mandanten_ID);
        }
        return Mandanten_ID;
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

    private static string GetColumnName(int columnNumber)
    {
        // Konvertiere den Spaltenindex in den entsprechenden Spaltenbuchstaben (z. B. 1 -> A, 2 -> B, ...)
        int dividend = columnNumber;
        string columnName = String.Empty;

        while (dividend > 0)
        {
            int modulo = (dividend - 1) % 26;
            columnName = Convert.ToChar(65 + modulo) + columnName;
            dividend = (dividend - modulo) / 26;
        }
        return columnName;
    }

    [ExcelFunction(Description = "Abfrage der Jahreszahlen für einen bestimmten Mandanten")]
    private static object Stamm_UStruktur(string host, int Mandanten_ID, int Jahr_ID)

    {
        object[,] ergebnis = null;
        try
        {
            string ur1 = $"http://{host}/api/data/clients/" + Mandanten_ID + "/ustruktur";
            var awaiter = CallURL(ur1);

            var myData = JsonSerializer.Deserialize<List<UStruktur>>(awaiter.Result);

            // Initialisiere das Ergebnis-Array basierend auf der Anzahl der passenden Einträge
            int count = myData.Count(item => item.Jahr_ID == Jahr_ID);
            ergebnis = new object[count + 1, 1]; // +1 für die Kopfzeile "ID und Name"

            ergebnis[0, 0] = "ID und Name";

            int rowIndex = 1; // Starte ab Zeile 1 für die Daten
            foreach (var item in myData)
            {
                if (item.Jahr_ID == Jahr_ID)
                {
                    ergebnis[rowIndex, 0] = $"{item.U_Struktur_ID} - {item.U_Name}";
                    rowIndex++;
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
   
    public class UStruktur
    {
        public int U_Struktur_ID { get; set; }
        public int Jahr_ID { get; set; }
        public string U_Nummer { get; set; }
        public string U_Name { get; set; }
        public int U_Struktur_ID_Prev { get; set; }
        public int Level { get; set; }
        public int U_Struktur_ID_Vorjahr { get; set; }
        public int Unterste_Ebene { get; set; }
        public int selected { get; set; }
        public int CGU_Art_ID { get; set; }
        public int Datenart_ID { get; set; }
        public int Steuer_ignorieren { get; set; }
        public int Anteil { get; set; }
        public string U_Name_2 { get; set; }
        public int Farbschema { get; set; }
        public int Synchron_ID { get; set; }
        public int Konsolidierungsstruktur { get; set; }
        public int U_Struktur_ID_Ausgleich { get; set; }
    }

    [ExcelFunction(Description = "Abfrage der Mandanten-ID und -Bezeichnung in einer Tabelle")]
    private static object Liste_Stamm_Mandanten_1(string host)
    {
        object[,] ergebnis;

        try
        {
            string ur1 = $"http://{host}/api/data/clients";
            var awaiter = CallURL(ur1);

            var myData = JsonSerializer.Deserialize<List<Mandanten_Liste>>(awaiter.Result);
            //var myData = JsonConvert.DeserializeObject<List<Mandanten_Liste>>(awaiter.Result);

            ergebnis = new object[myData.Count + 1, 2];
            //Array.Sort(ergebnis, (x, y) => string.Compare((string)x[1], (string)y[1]));

            ergebnis[0, 0] = "ID";
            ergebnis[0, 1] = "Name";

            myData = myData.OrderBy(x => x.Name).ToList(); //wurde nach Namen sortiert
                                                                     // var orderedArray = myData.OrderBy(x => x.Name).ToArray();
                                                                     //Array 
            for (int i = 0; i < myData.Count; i++)
            // for (int i = 0; i < orderedArray.Length; i++)
            {

                ergebnis[i + 1, 0] = myData[i].Mandant_ID;
                ergebnis[i + 1, 1] = myData[i].Name;

            }
         
            return ergebnis;
        }

        catch (Exception ex)
        {
            ergebnis = new object[1, 2];
            ergebnis[0, 0] = "Fehler:";
            Type exceptionType = ex.GetType();
            ergebnis[0, 1] = "Parameter nicht erkannt. Bitte geben Sie einen gültigen Servernamen, Port und Mandanten-ID ein. ";
            return ergebnis; // oder eine leere Liste oder einen anderen Wert, der zurückgegeben werden soll
        }
    }
    public class Mandanten_Liste
    {
        public long Mandant_ID { get; set; }
        public string Name { get; set; }
        public int Hauptwaehrung { get; set; }
        public int IF_ID { get; set; }
        public int Versatz { get; set; }
        public string Monat1 { get; set; }
        public string Monat2 { get; set; }
        public string Monat3 { get; set; }
        public string Monat4 { get; set; }
        public string Monat5 { get; set; }
        public string Monat6 { get; set; }
        public string Monat7 { get; set; }
        public string Monat8 { get; set; }
        public string Monat9 { get; set; }
        public string Monat10 { get; set; }
        public string Monat11 { get; set; }
        public string Monat12 { get; set; }
        public int letzte_Update_ID { get; set; }
        public string Module { get; set; }
        public int Berechnung_aus { get; set; }
        public int Runden_an { get; set; }
        public int Einstellung { get; set; }
        public string Excel_Exportpfad { get; set; }
        public int Pfad_verwenden { get; set; }
        public int Pfad_verw_alternat { get; set; }
        public int Schnittstellen_ID { get; set; }
        public int Mandanten_Art { get; set; }
        public int Letztes_GJ { get; set; }
        public int Letzter_MA { get; set; }
        public int DB_Size { get; set; }
        public string Hauptwaehrung_Name { get; set; }
        public int Erstes_GJ_SaldoAP { get; set; }
        public int Erstes_GJ_US_Synchron { get; set; }
        public string Erstes_Datum_ZT_Steuer_Grenze { get; set; }
        public int Erstes_GJ_Co { get; set; }
        public string Anschrift { get; set; }
        public string Systempasswort { get; set; }
        public string Bemerkung { get; set; }
    }

    [ExcelFunction(Description = "Abfrage der Mandanten-ID und -Bezeichnung in einer Tabelle")]
    private static object Liste_Stamm_Mandanten(string host, int layoutOption = 0)
    {
        object[,] ergebnis;

        try
        {
            string ur1 = $"http://{host}/api/data/clients";
            var awaiter = CallURL(ur1);

            var myData = JsonSerializer.Deserialize<List<Mandanten_Liste_1>>(awaiter.Result);
            //var myData = JsonConvert.DeserializeObject<List<Mandanten_Liste>>(awaiter.Result);

            ergebnis = new object[myData.Count + 1, 1];
            //Array.Sort(ergebnis, (x, y) => string.Compare((string)x[1], (string)y[1]));

            //ergebnis[0, 0] = "ID";
            //ergebnis[0, 1] = "Name";
            ergebnis[0, 0] = "ID und Name";

            myData = myData.OrderBy(x => x.Name).ToList();
            for (int i = 0; i < myData.Count; i++)
            // for (int i = 0; i < orderedArray.Length; i++)
            {
                //ergebnis[i + 1, 0] = myData[i].Mandant_ID;
                //ergebnis[i + 1, 1] = myData[i].Name;
                //ergebnis[i + 1, 2] = ergebnis[i + 1, 0] + " - " + ergebnis[i + 1, 1];
                ergebnis[i + 1, 0] = myData[i].Mandant_ID + " - " + myData[i].Name;

            }
            return ergebnis;
        }

        catch (Exception ex)
        {
            ergebnis = new object[1, 2];
            ergebnis[0, 0] = "Fehler:";
            Type exceptionType = ex.GetType();
            ergebnis[0, 1] = "Parameter nicht erkannt. Bitte geben Sie einen gültigen Servernamen, Port und Mandanten-ID ein. ";
            return ergebnis; // oder eine leere Liste oder einen anderen Wert, der zurückgegeben werden soll
        }
    }
    public class Mandanten_Liste_1
    {
        public long Mandant_ID { get; set; }
        public string Name { get; set; }
        public int Hauptwaehrung { get; set; }
        public int IF_ID { get; set; }
        public int Versatz { get; set; }
        public string Monat1 { get; set; }
        public string Monat2 { get; set; }
        public string Monat3 { get; set; }
        public string Monat4 { get; set; }
        public string Monat5 { get; set; }
        public string Monat6 { get; set; }
        public string Monat7 { get; set; }
        public string Monat8 { get; set; }
        public string Monat9 { get; set; }
        public string Monat10 { get; set; }
        public string Monat11 { get; set; }
        public string Monat12 { get; set; }
        public int letzte_Update_ID { get; set; }
        public string Module { get; set; }
        public int Berechnung_aus { get; set; }
        public int Runden_an { get; set; }
        public int Einstellung { get; set; }
        public string Excel_Exportpfad { get; set; }
        public int Pfad_verwenden { get; set; }
        public int Pfad_verw_alternat { get; set; }
        public int Schnittstellen_ID { get; set; }
        public int Mandanten_Art { get; set; }
        public int Letztes_GJ { get; set; }
        public int Letzter_MA { get; set; }
        public int DB_Size { get; set; }
        public string Hauptwaehrung_Name { get; set; }
        public int Erstes_GJ_SaldoAP { get; set; }
        public int Erstes_GJ_US_Synchron { get; set; }
        public string Erstes_Datum_ZT_Steuer_Grenze { get; set; }
        public int Erstes_GJ_Co { get; set; }
        public string Anschrift { get; set; }
        public string Systempasswort { get; set; }
        public string Bemerkung { get; set; }
        public int Layout { get; internal set; }
        public string Formatierung { get; internal set; }
    }

    [ExcelFunction(Description = "Abfrage der Jahreszahlen für einen bestimmten Mandanten")]
    private static object Liste_Stamm_Geschaeftsjahre(string host, int Mandanten_ID)

    {
        object[,] ergebnis = null;
        try
        {
            string ur1 = $"http://{host}/api/data/clients/";
            var awaiter = CallURL(ur1);
           
            var myData = JsonSerializer.Deserialize<List<Geschaeftsjahre1>>(awaiter.Result);
         

            //object[,] ergebnis = null;

            for (int i = 0; i < myData.Count; i++)
            {
            

                if (myData[i].Mandant_ID == Mandanten_ID)
                {

                    int yearsCount = myData[i].Letztes_GJ - myData[i].Erstes_GJ_Co + 1;
                    ergebnis = new object[yearsCount + 1, 1]; // +1 for the header

                    ergebnis[0, 0] = "Jahre";

                    for (int j = myData[i].Erstes_GJ_Co; j <= myData[i].Letztes_GJ; j++)
                    {
                        ergebnis[j - myData[i].Erstes_GJ_Co + 1, 0] = j;
                    }
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
    public class Geschaeftsjahre1
    {
        public long Mandant_ID { get; set; }
        public string Name { get; set; }
        public int Hauptwaehrung { get; set; }
        public int IF_ID { get; set; }
        public int Versatz { get; set; }
        public string Monat1 { get; set; }
        public string Monat2 { get; set; }
        public string Monat3 { get; set; }
        public string Monat4 { get; set; }
        public string Monat5 { get; set; }
        public string Monat6 { get; set; }
        public string Monat7 { get; set; }
        public string Monat8 { get; set; }
        public string Monat9 { get; set; }
        public string Monat10 { get; set; }
        public string Monat11 { get; set; }
        public string Monat12 { get; set; }
        public int letzte_Update_ID { get; set; }
        public string Module { get; set; }
        public int Berechnung_aus { get; set; }
        public int Runden_an { get; set; }
        public int Einstellung { get; set; }
        public string Excel_Exportpfad { get; set; }
        public int Pfad_verwenden { get; set; }
        public int Pfad_verw_alternat { get; set; }
        public int Schnittstellen_ID { get; set; }
        public int Mandanten_Art { get; set; }
        public int Letztes_GJ { get; set; }
        public int Letzter_MA { get; set; }
        public int DB_Size { get; set; }
        public string Hauptwaehrung_Name { get; set; }
        public int Erstes_GJ_SaldoAP { get; set; }
        public int Erstes_GJ_US_Synchron { get; set; }
        public string Erstes_Datum_ZT_Steuer_Grenze { get; set; }
        public int Erstes_GJ_Co { get; set; }
        public string Anschrift { get; set; }
        public string Systempasswort { get; set; }
        public string Bemerkung { get; set; }
    }

    [ExcelFunction(Description = "Abfrage der Jahreszahlen für einen bestimmten Mandanten")]
    private static object Liste_Stamm_Geschaeftsjahre_Zusammen(string host, int Mandanten_ID)

    {
        object[,] ergebnis = null;
        try
        {
            string ur1 = $"http://{host}/api/data/clients/";
            var awaiter = CallURL(ur1);

            var myData = JsonSerializer.Deserialize<List<Geschaeftsjahre_1>>(awaiter.Result);

            //object[,] ergebnis = null;

            for (int i = 0; i < myData.Count; i++)
            {

                if (myData[i].Mandant_ID == Mandanten_ID)
                {

                    ergebnis = new object[myData[i].Letztes_GJ - myData[i].Erstes_GJ_Co + 1, 1];

                    for (int j = myData[i].Erstes_GJ_Co; j <= myData[i].Letztes_GJ; j++)
                    {
                        ergebnis[j - myData[i].Erstes_GJ_Co, 0] = j;

                    }
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
    public class Geschaeftsjahre_1
    {
        public long Mandant_ID { get; set; }
        public string Name { get; set; }
        public int Hauptwaehrung { get; set; }
        public int IF_ID { get; set; }
        public int Versatz { get; set; }
        public string Monat1 { get; set; }
        public string Monat2 { get; set; }
        public string Monat3 { get; set; }
        public string Monat4 { get; set; }
        public string Monat5 { get; set; }
        public string Monat6 { get; set; }
        public string Monat7 { get; set; }
        public string Monat8 { get; set; }
        public string Monat9 { get; set; }
        public string Monat10 { get; set; }
        public string Monat11 { get; set; }
        public string Monat12 { get; set; }
        public int letzte_Update_ID { get; set; }
        public string Module { get; set; }
        public int Berechnung_aus { get; set; }
        public int Runden_an { get; set; }
        public int Einstellung { get; set; }
        public string Excel_Exportpfad { get; set; }
        public int Pfad_verwenden { get; set; }
        public int Pfad_verw_alternat { get; set; }
        public int Schnittstellen_ID { get; set; }
        public int Mandanten_Art { get; set; }
        public int Letztes_GJ { get; set; }
        public int Letzter_MA { get; set; }
        public int DB_Size { get; set; }
        public string Hauptwaehrung_Name { get; set; }
        public int Erstes_GJ_SaldoAP { get; set; }
        public int Erstes_GJ_US_Synchron { get; set; }
        public string Erstes_Datum_ZT_Steuer_Grenze { get; set; }
        public int Erstes_GJ_Co { get; set; }
        public string Anschrift { get; set; }
        public string Systempasswort { get; set; }
        public string Bemerkung { get; set; }
    }

    [ExcelFunction(Description = "Abfrage der Unternehmensstruktur für einen bestimmten Mandanten und ein bestimmtes Geschäftsjahr")]
    private static string Liste_Stamm_Unternehmensstruktur(string host, int Mandanten_ID, int Jahreszahl)
    {
        return $"Diese Funktion ist aktuell noch nicht implementiert";
    }

    [ExcelFunction(Description = "Abfrage des Aufbaus der GuV für einen bestimmten Mandanten")]
    private static object Liste_Aufbau_GuV(string host, int Mandanten_ID)

    {
        object[,] ergebnis;
        try
        {
            string ur1 = $"http://{host}/api/data/clients/" + Mandanten_ID + "/guv";
            var awaiter = CallURL(ur1);

            var myData = JsonSerializer.Deserialize<List<Aufbau_GuV>>(awaiter.Result);
            ergebnis = new object[myData.Count + 1, 3];

            ergebnis[0, 0] = "ID";
            ergebnis[0, 1] = "Reihenfolge";
            ergebnis[0, 2] = "Name";

            myData = myData.OrderBy(x => x.Reihenfolge).ToList(); //nach Reihenfolge sortieren 

            for (int i = 0; i < myData.Count; i++)
            {
                ergebnis[i + 1, 0] = myData[i].Pos_GuV_ID;
                ergebnis[i + 1, 1] = myData[i].Reihenfolge;
                ergebnis[i + 1, 2] = myData[i].Position_Name;
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
    public class Aufbau_GuV
    {
        public int Pos_GuV_ID { get; set; }
        public string Position_Name { get; set; }
        public int Vorzeichen { get; set; }
        public int Planungsart_ID { get; set; }
        public int MwSt_ID { get; set; }
        public int Bilanz_ID_aktiv { get; set; }
        public int Bilanz_ID_passiv { get; set; }
        public int Reihenfolge { get; set; }
        public int Statistik { get; set; }
        public int Systemzeile { get; set; }
        public int Zeilenlayout_ID { get; set; }
        public int Quotierung_Ignorieren { get; set; }
        public int Saldierungstyp { get; set; }
        public int FixVariabel { get; set; }
        public int Einzug { get; set; }
        public int BA2VE { get; set; }
    }

   

    [ExcelFunction(Description = "Abfrage des Aufbaus der Aktiva für einen bestimmten Mandanten")]
    private static object Liste_Aufbau_Aktiva(string host, int Mandanten_ID)
    {
        object[,] ergebnis;
        try
        {
            string ur1 = $"http://{host}/api/data/clients/" + Mandanten_ID + "/aktiva";
            var awaiter = CallURL(ur1);

            var myData = JsonSerializer.Deserialize<List<Aufbau_Aktiva>>(awaiter.Result);
            ergebnis = new object[myData.Count + 1, 3];

            ergebnis[0, 0] = "ID";
            ergebnis[0, 1] = "Reihenfolge";
            ergebnis[0, 2] = "Name";

            myData = myData.OrderBy(x => x.Reihenfolge).ToList();

            for (int i = 0; i < myData.Count; i++)
            {
                ergebnis[i + 1, 0] = myData[i].Pos_Bilanz_ID;
                ergebnis[i + 1, 1] = myData[i].Reihenfolge;
                ergebnis[i + 1, 2] = myData[i].Position_Name;
            }
        }
        catch (Exception ex)
        {
            ergebnis = new object[1, 2];
            ergebnis[0, 0] = "Fehler:";
            Type exceptionType = ex.GetType();
            ergebnis[0, 1] = "Parameter nicht erkannt. Bitte geben Sie einen gültigen Servernamen, Port und Mandanten-ID ein. ";
        }
        return ergebnis;
    }
    public class Aufbau_Aktiva
    {
        public int Pos_Bilanz_ID { get; set; }
        public string Position_Name { get; set; }
        public int Planungsart_ID { get; set; }
        public string Gliederung { get; set; }
        public int Reihenfolge { get; set; }
        public int Statistik { get; set; }
        public int Anlagespiegel { get; set; }
        public int Systemzeile_Aktiva { get; set; }
        public int Zeilenlayout_ID { get; set; }
        public int Quotierung_Ignorieren { get; set; }
        public int FixVariabel { get; set; }
        public int AufloesungUeberOP { get; set; }
        public int InBankAB { get; set; }
        public int Einzug { get; set; }
    }

    [ExcelFunction(Description = "Abfrage des Aufbaus der Passiva für einen bestimmten Mandanten")]
    private static object Liste_Aufbau_Passiva(string host, int Mandanten_ID)

    {
        object[,] ergebnis;
        try
        {
            string ur1 = $"http://{host}/api/data/clients/" + Mandanten_ID + "/passiva";
            var awaiter = CallURL(ur1);

            var myData = JsonSerializer.Deserialize<List<Aufbau_Passiva>>(awaiter.Result);
            ergebnis = new object[myData.Count + 1, 3];

            ergebnis[0, 0] = "ID";
            ergebnis[0, 1] = "Reihenfolge";
            ergebnis[0, 2] = "Name";

            myData = myData.OrderBy(x => x.Reihenfolge).ToList();  //nach Reihenfolge sortieren 

            for (int i = 0; i < myData.Count; i++)
            {

                ergebnis[i + 1, 0] = myData[i].Pos_Bilanz_ID;
                ergebnis[i + 1, 1] = myData[i].Reihenfolge;
                ergebnis[i + 1, 2] = myData[i].Position_Name;

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
    public class Aufbau_Passiva
    {
        public int Pos_Bilanz_ID { get; set; }
        public string Position_Name { get; set; }
        public int Planungsart_ID { get; set; }
        public string Gliederung { get; set; }
        public int Reihenfolge { get; set; }
        public int Statistik { get; set; }
        public int Systemzeile_Passiva { get; set; }
        public int Zeilenlayout_ID { get; set; }
        public int Quotierung_Ignorieren { get; set; }
        public int FixVariabel { get; set; }
        public int AufloesungUeberOP { get; set; }
        public int InBankAB { get; set; }
        public int Einzug { get; set; }
    }



    [ExcelFunction(Description = "Abfrage der Berichte für einen bestimmten Mandanten")]
    private static string Liste_Aufbau_Berichte(string host, int Mandanten_ID)
    {
        return $"Diese Funktion ist aktuell noch nicht implementiert";
    }


    [ExcelFunction(Description = "Abfrage der Berichte für einen bestimmten Mandanten")]
    private static string Liste_Aufbau_Bericht(string host, int Mandanten_ID, int Berichts_ID)
    {
        return $"Diese Funktion ist aktuell noch nicht implementiert";
    }

    [ExcelFunction(Description = "Abfrage der GuV-Daten für einen bestimmten Mandanten, ein bestimmtes Jahr, eine bestimmte Unternehmensstruktur und eine bestimmte Datenart")]
    private static object Liste_Daten_GuV(string host, int Mandanten_ID, int Jahr_ID, int UStruktur_ID, int selectedDatenartID)

    {
        object[,] ergebnis;
        try
        {
            string ur1 = $"http://{host}/api/data/income/" + Mandanten_ID + "/" + Jahr_ID + "/" + UStruktur_ID + "/" + selectedDatenartID;
            var awaiter = CallURL(ur1);

            var myData = JsonSerializer.Deserialize<List<GuV_Liste>>(awaiter.Result);

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
                    ergebnis[i + 1, 2] = myData[i].Monat1;       // ToString("N2"); formatierung
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

    [ExcelFunction(Description = "Abfrage der GuV-Daten für einen bestimmten Mandanten, ein bestimmtes Jahr, eine bestimmte Unternehmensstruktur und eine bestimmte Datenart")]
    private static object Liste_Daten_GuV_Test(string host, int Mandanten_ID, int Jahr_ID, int UStruktur_ID, int selectedDatenartID)

    {
        object[,] ergebnis;
        try
        {
            string ur1 = $"http://{host}/api/data/income/" + Mandanten_ID + "/" + Jahr_ID + "/" + UStruktur_ID + "/" + selectedDatenartID;
            var awaiter = CallURL(ur1);

            var myData = JsonSerializer.Deserialize<List<GuV_Liste_Test>>(awaiter.Result);

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
                    ergebnis[i + 1, 2] = myData[i].Monat1;       // ToString("N2"); formatierung
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
    public class GuV_Liste_Test
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


    [ExcelFunction(Description = "Abfrage der Passiva-Daten für einen bestimmten Mandanten, ein bestimmtes Jahr, eine bestimmte Unternehmensstruktur und eine bestimmte Datenart")]
    private static object Liste_Daten_Passiva(string host, int Mandanten, int Jahr, int UStruktur, int Datenart)
    {
        object[,] ergebnis;
        try
        {
            string ur1 = $"http://{host}/api/data/balance/" + Mandanten + "/" + Jahr + "/" + UStruktur + "/" + Datenart;
            var awaiter = CallURL(ur1);

            var myData = JsonSerializer.Deserialize<List<Passiva>>(awaiter.Result);

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

    [ExcelFunction(Description = "Abfrage der Berichtsdaten für einen bestimmten Mandanten, ein bestimmtes Jahr, eine bestimmte Unternehmensstruktur, eine bestimmte Datenart und einen bestimmten Bericht")]
    private static string Liste_Daten_Bericht(string host, int Mandanten_ID, int Jahr_ID, int Unternehmensstruktur_ID, int Datenart_ID, int Berichts_ID)
    {
        return $"Diese Funktion ist aktuell noch nicht implementiert";
    }

    [ExcelFunction(Description = "Abfrage der Passiva-Daten für einen bestimmten Mandanten, ein bestimmtes Jahr, eine bestimmte Unternehmensstruktur und eine bestimmte Datenart")]
    private static object Liste_Daten_Aktiva(string host, int Mandanten, int Jahr, int UStruktur, int Datenart)

    {
        object[,] ergebnis;
        try
        {
            string ur1 = $"http://{host}/api/data/balance/" + Mandanten + "/" + Jahr + "/" + UStruktur + "/" + Datenart;
            var awaiter = CallURL(ur1);

            var myData = JsonSerializer.Deserialize<List<Aktiva>>(awaiter.Result);

            ergebnis = new object[myData.Count + 1, 17];

            ergebnis[0, 0] = "ID";
            ergebnis[0, 1] = "Reihenfolge";
            ergebnis[0, 2] = "Position_Name";
            ergebnis[0, 3] = "Anfangsbestand";
            ergebnis[0, 4] = "Monat1";
            ergebnis[0, 5] = "Monat2";
            ergebnis[0, 6] = "Monat3";
            ergebnis[0, 7] = "Monat4";
            ergebnis[0, 8] = "Monat5";
            ergebnis[0, 9] = "Monat6";
            ergebnis[0, 10] = "Monat7";
            ergebnis[0, 11] = "Monat8";
            ergebnis[0, 12] = "Monat9";
            ergebnis[0, 13] = "Monat10";
            ergebnis[0, 14] = "Monat11";
            ergebnis[0, 15] = "Monat12";
            ergebnis[0, 16] = "Endbestand";

            for (int i = 0; i < myData.Count; i++)
            {
                {
                    ergebnis[i + 1, 0] = myData[i].Pos_Bilanz_ID;
                    ergebnis[i + 1, 1] = myData[i].Reihenfolge;
                    ergebnis[i + 1, 2] = myData[i].Position_Name;
                    ergebnis[i + 1, 3] = myData[i].Anfangsbestand;
                    ergebnis[i + 1, 4] = myData[i].Monat1_S;
                    ergebnis[i + 1, 5] = myData[i].Monat2_S;
                    ergebnis[i + 1, 6] = myData[i].Monat3_S;
                    ergebnis[i + 1, 7] = myData[i].Monat4_S;
                    ergebnis[i + 1, 8] = myData[i].Monat5_S;
                    ergebnis[i + 1, 9] = myData[i].Monat6_S;
                    ergebnis[i + 1, 10] = myData[i].Monat7_S;
                    ergebnis[i + 1, 11] = myData[i].Monat8_S;
                    ergebnis[i + 1, 12] = myData[i].Monat9_S;
                    ergebnis[i + 1, 13] = myData[i].Monat10_S;
                    ergebnis[i + 1, 14] = myData[i].Monat11_S;
                    ergebnis[i + 1, 15] = myData[i].Monat12_S;
                    ergebnis[i + 1, 16] = myData[i].Endbestand;
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
    public class Aktiva
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

    [ExcelFunction(Description = "Gibt den Zellnamen der aktiven Zelle zurück")]
    public static string GetActiveCellName()
    {
        try
        {
            ExcelReference caller = XlCall.Excel(XlCall.xlfCaller) as ExcelReference;
            if (caller != null)
            {
                int rowNumber = caller.RowFirst + 1;
                int columnNumber = caller.ColumnFirst + 1;

                string columnName = GetColumnName(columnNumber);

                return columnName + rowNumber;
            }
        }
        catch
        {
            // Fehlerbehandlung hier (optional)
        }

        return "Fehler";
    }


    public class Passiva02
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

    
    [ExcelFunction(Description = "Abfrage der GuV-Daten für einen bestimmten Mandanten, ein bestimmtes Jahr, eine bestimmte Unternehmensstruktur und eine bestimmte Datenart")]
    public static object Daten_GuV(string host, int Mandanten_ID, int Jahr_ID, int UStruktur_ID, int Datenart_ID)

    {
        object[,] ergebnis;
        try
        {
            string ur1 = $"http://{host}/api/data/income/" + Mandanten_ID + "/" + Jahr_ID + "/" + UStruktur_ID + "/" + Datenart_ID;
            var awaiter = CallURL(ur1);

            var myData = JsonSerializer.Deserialize<List<GuV_layout>>(awaiter.Result);

            ergebnis = new object[myData.Count + 1, 9];

            ergebnis[0, 0] = "Reihenfolge";
            ergebnis[0, 1] = "Name";
            ergebnis[0, 2] = "Pos_GuV_ID";
            ergebnis[0, 3] = "ZL_Zeilenfarbe";
            ergebnis[0, 4] = "ZL_Transparenz";
            ergebnis[0, 5] = "ZL_Textfarbe";
            ergebnis[0, 6] = "ZL_Schrift_Style";
            ergebnis[0, 7] = "ZL_Schrift_Name";
            ergebnis[0, 8] = "ZL_Schrift_Groesse";

            {
                for (int i = 0; i < myData.Count; i++)
                {
                    ergebnis[i + 1, 0] = myData[i].Reihenfolge;
                    ergebnis[i + 1, 1] = myData[i].Position_Name;
                    ergebnis[i + 1, 2] = myData[i].Pos_GuV_ID;
                    ergebnis[i + 1, 3] = myData[i].ZL_Zeilenfarbe;
                    ergebnis[i + 1, 4] = myData[i].ZL_Transparenz;
                    ergebnis[i + 1, 5] = myData[i].ZL_Textfarbe;
                    ergebnis[i + 1, 6] = myData[i].ZL_Schrift_Style;
                    ergebnis[i + 1, 7] = myData[i].ZL_Schrift_Name;
                    ergebnis[i + 1, 8] = myData[i].ZL_Schrift_Groesse;
                }
            }

        }
        catch (Exception ex)
        {
            ergebnis = new object[1, 2];
            ergebnis[0, 0] = "Fehler:";
            Type exceptionType = ex.GetType();
            ergebnis[0, 1] = "Parameter nicht erkannt. Bitte geben Sie einen gültigen Servernamen, Port und Mandanten-ID ein. ";
        }
        return ergebnis;
    }
    public class GuV_layout
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

    [ExcelFunction(Description = "Abfrage der Passiva-Daten für einen bestimmten Mandanten, ein bestimmtes Jahr, eine bestimmte Unternehmensstruktur und eine bestimmte Datenart")]
    public static object Daten_Bilanz(string host, int Mandanten, int Jahr, int UStruktur, int Datenart)
    {
        object[,] ergebnis;
        try
        {
            string ur1 = $"http://{host}/api/data/balance/" + Mandanten + "/" + Jahr + "/" + UStruktur + "/" + Datenart;
            var awaiter = CallURL(ur1);

            var myData = JsonSerializer.Deserialize<List<Bilanz_Daten>>(awaiter.Result);

            ergebnis = new object[myData.Count + 1, 9];

            ergebnis[0, 0] = "Reihenfolge";
            ergebnis[0, 1] = "Name";
            ergebnis[0, 2] = "Pos_Bilanz_ID";
            ergebnis[0, 3] = "ZL_Zeilenfarbe";
            ergebnis[0, 4] = "ZL_Transparenz";
            ergebnis[0, 5] = "ZL_Textfarbe";
            ergebnis[0, 6] = "ZL_Schrift_Style";
            ergebnis[0, 7] = "ZL_Schrift_Name";
            ergebnis[0, 8] = "ZL_Schrift_Groesse";

            {
                for (int i = 0; i < myData.Count; i++)
                {
                    ergebnis[i + 1, 0] = myData[i].Reihenfolge;
                    ergebnis[i + 1, 1] = myData[i].Position_Name;
                    ergebnis[i + 1, 2] = myData[i].Pos_Bilanz_ID;
                    ergebnis[i + 1, 3] = myData[i].ZL_Zeilenfarbe;
                    ergebnis[i + 1, 4] = myData[i].ZL_Transparenz;
                    ergebnis[i + 1, 5] = myData[i].ZL_Textfarbe;
                    ergebnis[i + 1, 6] = myData[i].ZL_Schrift_Style;
                    ergebnis[i + 1, 7] = myData[i].ZL_Schrift_Name;
                    ergebnis[i + 1, 8] = myData[i].ZL_Schrift_Groesse;

                }
            }
        }
        catch (Exception ex)
        {
            ergebnis = new object[1, 2];
            ergebnis[0, 0] = "Fehler:";
            Type exceptionType = ex.GetType();
            ergebnis[0, 1] = "Parameter nicht erkannt. Bitte geben Sie einen gültigen Servernamen, Port und Mandanten-ID ein. ";
        }
        return ergebnis;
    }

    public class Bilanz_Daten
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
    public static object Stamm_UStruktur(string Mandanten_ID_Name, int Jahr_ID)
    {
        string host = ReadRegistry();
        int Mandanten_ID = GetMandantenIDFromName(Mandanten_ID_Name);
        object[,] ergebnis = (object[,])Stamm_UStruktur(host, Mandanten_ID, Jahr_ID);
        return ergebnis;
    }
    private static object Liste_Stamm_Mandanten_1()
    {
        string host = ReadRegistry();
        object[,] ergebnis = (object[,])Liste_Stamm_Mandanten_1(host);
        return ergebnis;
    }
    public static object Liste_Stamm_Mandanten()
    {
        string host = ReadRegistry();
        object[,] ergebnis = (object[,])Liste_Stamm_Mandanten(host);
        return ergebnis;
    }

    public static object Liste_Stamm_Geschaeftsjahre(string Mandanten_ID_Name)
    {
        string host = ReadRegistry();
        int Mandanten_ID = GetMandantenIDFromName(Mandanten_ID_Name);
        object[,] ergebnis = (object[,])Liste_Stamm_Geschaeftsjahre(host, Mandanten_ID);
        return ergebnis;
    }

    public static object Liste_Stamm_Geschaeftsjahre_Zusammen(string Mandanten_ID_Name)
    {
        string host = ReadRegistry();
        int Mandanten_ID = GetMandantenIDFromName(Mandanten_ID_Name);
        object[,] ergebnis = (object[,])Liste_Stamm_Geschaeftsjahre(host, Mandanten_ID);
        return ergebnis;
    }

    public static object[,] Liste_Stamm_Unternehmensstruktur()
    {
        throw new NotImplementedException();  // wurde von Liste_Stamm_Unternehmensstruktur generiert. weil es keine Inhalt gibt.
    }

    public static object Liste_Aufbau_GuV(string Mandanten_ID_Name)
    {
        string host = ReadRegistry();
        int Mandanten_ID = GetMandantenIDFromName(Mandanten_ID_Name);
        object[,] ergebnis = (object[,])Liste_Aufbau_GuV(host, Mandanten_ID);
        return ergebnis;
    }
    public static object Liste_Aufbau_Aktiva(string Mandanten_ID_Name)
    {
        string host = ReadRegistry();
        int Mandanten_ID = GetMandantenIDFromName(Mandanten_ID_Name);
        object[,] ergebnis = (object[,])Liste_Aufbau_Aktiva(host, Mandanten_ID);
        return ergebnis;
    }

    public static object Liste_Aufbau_Passiva(string Mandanten_ID_Name)
    {
        string host = ReadRegistry();
        int Mandanten_ID = GetMandantenIDFromName(Mandanten_ID_Name);
        object[,] ergebnis = (object[,])Liste_Aufbau_Passiva(host, Mandanten_ID);
        return ergebnis;
    }

    public static object Liste_Aufbau_Berichte(string Mandanten_ID_Name)
    {
        string host = ReadRegistry();
        int Mandanten_ID = GetMandantenIDFromName(Mandanten_ID_Name);
        object[,] ergebnis = (object[,])Liste_Aufbau_Berichte(host);
        return ergebnis;
    }

    public static object Liste_Aufbau_Bericht(string Mandanten_ID_Name, int Berichts_ID)
    {
        string host = ReadRegistry();
        int Mandanten_ID = GetMandantenIDFromName(Mandanten_ID_Name);
        object[,] ergebnis = (object[,])Liste_Aufbau_Bericht();
        return ergebnis;
    }

    private static object[,] Liste_Aufbau_Bericht()
    {
        throw new NotImplementedException(); // wurde von Liste_Aufbau_Bericht generiert. weil es keine Inhalt gibt.
    }

    public static object Liste_Daten_GuV_Test(string Mandanten_ID_Name, int Jahr_ID, string UStruktur_ID_Name, string Datenart_ID)
    {
        string host = ReadRegistry();
        int Mandanten_ID = GetMandantenIDFromName(Mandanten_ID_Name);
        int UStruktur_ID = GetUStrukturIDFromName(UStruktur_ID_Name);
        int selectedDatenartID;

        // Überprüfen, ob Datenart_ID ein Text ist und die entsprechenden Zuordnungen vornehmen
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
            object[,] ergebnis = (object[,])Liste_Daten_GuV(host, Mandanten_ID, Jahr_ID, UStruktur_ID, selectedDatenartID);
            return ergebnis;
        }
        else
        {
            throw new ArgumentNullException("Datenart_ID", "Datenart_ID darf nicht leer sein.");
        }
    }

    public static object Liste_Daten_GuV(string Mandanten_ID_Name, int Jahr_ID, string UStruktur_ID_Name, string Datenart_ID)
    {
        string host = ReadRegistry();
        int Mandanten_ID = GetMandantenIDFromName(Mandanten_ID_Name);
        int UStruktur_ID = GetUStrukturIDFromName(UStruktur_ID_Name);
        int selectedDatenartID;

        // Überprüfen, ob Datenart_ID ein Text ist und die entsprechenden Zuordnungen vornehmen
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
            object[,] ergebnis = (object[,])Liste_Daten_GuV(host, Mandanten_ID, Jahr_ID, UStruktur_ID, selectedDatenartID);
            return ergebnis;
        }
        else
        {
            throw new ArgumentNullException("Datenart_ID", "Datenart_ID darf nicht leer sein.");
        }
    }

    public static object Liste_Daten_Passiva(string Mandanten_ID_Name, int Jahr, int UStruktur, int Datenart)
    {
        string host = ReadRegistry();
        int Mandanten_ID = GetMandantenIDFromName(Mandanten_ID_Name);
        object[,] ergebnis = (object[,])Liste_Daten_Passiva(host, Mandanten_ID, Jahr, UStruktur, Datenart);
        return ergebnis;
    }

    public static object Liste_Daten_Bericht(string Mandanten, int Jahr, int UStruktur, int Datenart)
    {
        string host = ReadRegistry();

        object[,] ergebnis = (object[,])Liste_Daten_Bericht(host, Mandanten, Jahr, UStruktur, Datenart);
        return ergebnis;
    }

    private static object[,] Liste_Daten_Bericht(string host, string Mandanten_ID_Name, int jahr, int uStruktur, int datenart)
    {
        throw new NotImplementedException();   // wurde von Liste_Daten_Bericht generiert. weil es keine Inhalt gibt.
    }

    public static object Liste_Daten_Aktiva(string Mandanten_ID_Name, int Jahr, int UStruktur, int Datenart)
    {
        string host = ReadRegistry();
        int Mandanten_ID = GetMandantenIDFromName(Mandanten_ID_Name);
        object[,] ergebnis = (object[,])Liste_Daten_Aktiva(host, Mandanten_ID, Jahr, UStruktur, Datenart);
        return ergebnis;
    }
}




