using MuenImport.Gui.Panels;
using System;
using System.Diagnostics;
using System.IO;
using System.Windows;
using System.Windows.Controls;

namespace MuenImport.Gui {
    public class GuiMgmt {


        private Window appWindow;
        public Window AppWindow {
            get {
                return appWindow;
            }
            set {
                appWindow = value;
            }
        }

        private MainPanel mp;
        public MainPanel Mp {
            get {
                return mp;
            }
            set {
                mp = value;
            }
        }

        private string zielverzeichnis;



        /// <summary>
        /// Erzeugt ein neues GuiMgmt.
        /// </summary>
        /// <param name="v">Das Viewmodel der Anwendung/des Hauptfensters.</param>
        public GuiMgmt(  ) {
            appWindow = Application.Current.Windows[0];
        }

        internal void init( MainWindow mainWindow ) {
            setupSubPanels( mainWindow );
        }

        /// <summary>
        /// Gui wird hier zusammengebaut.
        /// </summary>
        /// <param name="mainWindow"></param>
        /// <param name="mc"></param>
        private void setupSubPanels( MainWindow mainWindow ) {

            mp = new MainPanel();
            mp.SetValue( Grid.RowProperty, 1 );
            mainWindow.mainGrid.Children.Add( mp );
            
        }

        /// <summary>
        /// Öffnet einen Dateidialog liefert einen Dateinamen für die Sicherung zurück..
        /// </summary>
        /// <returns></returns>
        public string saveDialog() {
            string f = "";

            Microsoft.Win32.SaveFileDialog dlg = new Microsoft.Win32.SaveFileDialog();
            dlg.DefaultExt = ".ein";
            dlg.Filter = "XML Dateien (.xml)|*.xml";

            Nullable<bool> result = dlg.ShowDialog();
            if ( result == true ) {

                f = dlg.FileName;
                zielverzeichnis = Path.GetDirectoryName( dlg.FileName );
            }
            else {
                return null;
            }
            return f;
        }




        /// <summary>
        /// Öffnet einen Dateidialog liefert alle ausgewählten Dateinamen zurück.
        /// </summary>
        /// <returns></returns>
        internal string openDialog() {
            // Create OpenFileDialog
            string filename = null;
            Microsoft.Win32.OpenFileDialog dlg = new Microsoft.Win32.OpenFileDialog();

            // Set filter for file extension and default file extension
            dlg.DefaultExt = ".xlsx";
            dlg.Filter = "Excel Dateien (.xls*)|*.xls*";
            dlg.Multiselect = false;

            // Display OpenFileDialog by calling ShowDialog method
            Nullable<bool> result = dlg.ShowDialog();

            // Get the selected file name and display in a TextBox
            if ( result == true ) {

                filename = dlg.FileName;
                zielverzeichnis = Path.GetDirectoryName( dlg.FileName );
            }
            return ( filename );
        }





        /// <summary>
        /// Den Explorer im bereitstehenden Verzeichnis öffnen.
        /// </summary>
        /// <param name="path"></param>
        public void OpenExplorer() {
            if ( Directory.Exists( zielverzeichnis ) )
                Process.Start( "explorer.exe", zielverzeichnis );
        }
    }
}
