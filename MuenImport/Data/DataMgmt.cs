using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Text;
using System.Text.RegularExpressions;
using MuenImport.Gui;
using MuenImport.Gui.Panels;
using System.Xml;
using MuenImport.Data;
using MuenImport.Data.DataObjects;

namespace MuenImport.Data {
    /// <summary>
    /// Klasse: DataNgmt.
    /// Regelt zentralisiert IO-Funktionen und Funktionen der Geschäftslogik.
    /// </summary>
    public class DataMgmt {
        /// <summary>
        /// Das Gui-Management.
        /// </summary>
        private GuiMgmt gm;       
        /// <summary>
        /// Liste der Haltungsobjekte.
        /// </summary>
        private List<Haltung> haltungen;
        /// <summary>
        /// Klasse mit Parsing und XML-Erstellungsroutinen.
        /// </summary>
        private Parser parse;


        /// <summary>
        /// Initialisierung des Data-Management.
        /// </summary>
        /// <param name="gm">Das Gui-Management.</param>
        public void init( GuiMgmt g ) {
            gm = g;
            parse = new Parser();
            Console.WriteLine( "Initialisierung fertig" );
        }

        /// <summary>
        /// Aufruf der OpenFile Dateibehandlung.
        /// </summary>
        /// <returns>Den Dateipfad.</returns>
        internal String loadMuenFile() {
            string fn = gm.openDialog();
            return fn;
        }

        /// <summary>
        /// Aufruf der Erstellroutine für die XML-Datei.
        /// </summary>
        /// <param name="fn">Der Filename.</param>
        /// <returns>Das erstellte XMLDocument.</returns>
        internal XmlDocument createXml( string fn ) {
            haltungen = parse.load( fn );
            XmlDocument doc = parse.createXml( haltungen, gm.Mp );
            return doc;
        }

        /// <summary>
        /// Aufruf der SaveFile Dateibehandlung.
        /// </summary>
        /// <param name="doc">Das zu speichernde XMLDocument.</param>
        internal void saveMuenFile( XmlDocument doc ) {
            string f = gm.saveDialog();
            if ( f != null && !f.Equals( "" ) && doc != null ) {
                doc.Save( f );
            }
        }
    }
}

