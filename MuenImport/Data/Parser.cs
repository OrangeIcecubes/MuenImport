//using LinqToExcel;
using MuenImport.Data.DataObjects;
using MuenImport.Gui.Panels;
using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Xml;
using exl = Microsoft.Office.Interop.Excel;

namespace MuenImport.Data
{
    /// <summary>
    /// Stellt Methoden zur Verfügung, die das parsen der Exceldatei in eine interne Objekt-Struktur (Haltung)
    /// und das Erstellen eines XMLDocument erledigen.
    /// </summary>
    public class Parser
    {
        private exl.Worksheet ws;

        /// <summary>
        /// Öffnet eine Exceldatei (MüP-Liste) und parst bestimmte Eigenschaften daraus
        /// in eine Liste aus Objekten der Klasse Data.Haltung.
        /// </summary>
        /// <param name="excelFileName">Pfad zur Exceldatei (MüP-Liste), die umgewandelt werden soll.</param>
        /// <returns>Eine Liste von Haltungen und deren relevanter Eigenschaften. siehe Class Data.Haltung</returns>
        public List<Haltung> load(string excelFileName)
        {

            if (excelFileName == null) return null;

            exl.Workbook wb = openTemplateWorkbook(excelFileName);
            ws = openSheet(wb, "MÜP-Liste");

            List<Haltung> haltungen = new List<Haltung>();




            exl.Range range = ws.UsedRange;
            var selektion =
                from exl.Range h in range.Rows
                where Convert.ToString(h.Cells[11].Value2).Length == 16
                select h;

            foreach (var i in selektion)
            {
                //    if (i["Haltung"].ToString().Length == 16)
                //    {

                Haltung hh = new Haltung();
                hh.Haltungsnummer = i.Cells[11].Value2;
                hh.NameMaßnahme = i.Cells[10].Value2;
                hh.KostenEpText = Convert.ToString(i.Cells[27].Value2);
                hh.KostenFpText = Convert.ToString(i.Cells[29].Value2);

                hh.Fertigstellung = Convert.ToString(i.Cells[50].Value2);


                if (i.Cells.Count > 55)
                {
                    hh.OidHaltung = Convert.ToInt32(i.Cells[56].Value2);
                    hh.OidInspektion = Convert.ToInt32(i.Cells[57].Value2);
                }

                                             
                if (hh.Fertigstellung == null)
                {
                    hh.Fertigstellung = "9999";
                }

                if (hh.KostenEpText == "")
                {
                    hh.KostenEpText = "0";
                }

                if (hh.KostenFpText == "")
                {
                    hh.KostenFpText = "0";
                }

                //        hh.Kosten = Convert.ToDecimal(hh.KostenEpText) + Convert.ToDecimal(hh.KostenFpText);

                haltungen.Add(hh);
                //    }
            }

            int c = haltungen.Count();
            return haltungen;
        }

        /// <summary>
        /// Erstellt aus der zuvor erstellten Haltungsliste eine XML-Datei.
        /// </summary>
        /// <param name="haltungen">Liste der Haltungen und deren Eigenschaften.</param>
        /// <param name="m">Das MainPanel zur Darstellung des Verarbeitungsergebnisses.</param>
        /// <returns>Das erstellte XmlDocument.</returns>
        public XmlDocument createXml(List<Haltung> haltungen, MainPanel m)
        {

            if (haltungen == null) return null;

            int anzahlMaßnahmen = 0;
            int anzahlHaltungen = 0;
            string aTyp = m.txAuftragTyp.Text;


            XmlDocument doc = new XmlDocument();
            XmlNode decl = doc.CreateXmlDeclaration("1.0", "utf-8", "yes");
            doc.AppendChild(decl);

            XmlElement xAuftraege = doc.CreateElement("Auftraege");
            doc.AppendChild(xAuftraege);

            var au =
                from h in haltungen
                group h by h.NameMaßnahme into g
                select new { Auftrag = g.Key, Haltungsgruppe = g, };

            anzahlMaßnahmen = au.Count();

            foreach (var a in au)
            {

                XmlElement xAuftrag = doc.CreateElement("Auftrag");
                xAuftraege.AppendChild(xAuftrag);


                XmlElement xAuftragName = doc.CreateElement("AuftragName");
                XmlElement xAuftragTyp = doc.CreateElement("AuftragTyp");
                XmlElement xBudget = doc.CreateElement("Budget");
                XmlElement xBaujahr = doc.CreateElement("Baujahr");
                XmlElement xHaltungen = doc.CreateElement("Haltungen");


                xAuftragName.InnerText = a.Auftrag;
                if (aTyp.Equals(""))
                {
                    aTyp = "unbekannt";
                }
                xAuftragTyp.InnerText = aTyp;
                xBudget.InnerText = Convert.ToString(a.Haltungsgruppe.Sum(bu => bu.Kosten));

                string bj = a.Haltungsgruppe.Min(bu => bu.Fertigstellung);
                if (bj != null && !bj.Equals(""))
                {
                    xBaujahr.InnerText = bj;
                }
                else
                {
                    xBaujahr.InnerText = "9999";
                }

                xAuftrag.AppendChild(xAuftragName);
                xAuftrag.AppendChild(xAuftragTyp);
                xAuftrag.AppendChild(xBudget);
                xAuftrag.AppendChild(xBaujahr);
                xAuftrag.AppendChild(xHaltungen);

                anzahlHaltungen += a.Haltungsgruppe.Count();

                foreach (var hh in a.Haltungsgruppe)
                {
                    XmlElement xHaltung = doc.CreateElement("Haltung");
                    xHaltungen.AppendChild(xHaltung);

                    XmlElement xHaltungsnummer = doc.CreateElement("Haltungsnummer");
                    xHaltungsnummer.InnerText = hh.Haltungsnummer;
                    xHaltung.AppendChild(xHaltungsnummer);

                    if (hh.OidHaltung > 0)
                    {
                        XmlElement xOidHaltung = doc.CreateElement("OidHaltung");
                        xOidHaltung.InnerText = hh.OidHaltung.ToString();
                        xHaltung.AppendChild(xOidHaltung);
                    }

                    if (hh.OidInspektion > 0)
                    {
                        XmlElement xOidInspektion = doc.CreateElement("OidInspektion");
                        xOidInspektion.InnerText = hh.OidInspektion.ToString();
                        xHaltung.AppendChild(xOidInspektion);
                    }
                }
            }

            m.TB1.Text = "Info: Anzahl Maßnahmen: " + anzahlMaßnahmen.ToString() + " - Anzahl Haltungen: " + anzahlHaltungen.ToString();

            return (doc);
        }
        /// <summary>
        /// Öffnet die Excel-Datei.
        /// </summary>
        /// <returns></returns>
        private exl.Workbook openTemplateWorkbook(string path)
        {
            exl.Application ap = new exl.Application();
            //ap.dev
            ap.DisplayAlerts = false;
            exl.Workbook wb = null;

            try
            {
                wb = ap.Workbooks.Open(path, null, true);
            }
            catch (Exception e)
            {

                Console.WriteLine(e.Message);
            }

            return wb;
        }

        /// <summary>
        /// Öffnet das Wotrksheet 
        /// </summary>
        /// <param name="wb"></param>
        /// <param name="p"></param>
        /// <returns></returns>
        private exl.Worksheet openSheet(exl.Workbook wb, string p)
        {
            exl.Worksheet ws = null;
            ws = wb.Sheets[p];

            return ws;
        }
    }
}