using MuenImport.Data;
using MuenImport.Gui;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Navigation;
using System.Windows.Shapes;
using System.Xml;

namespace MuenImport {
    /// <summary>
    /// Interaktionslogik für MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window {

        private DataMgmt dm;
        private GuiMgmt gm;
        private Parser parse;

        public MainWindow() {
            InitializeComponent();
            dm = new DataMgmt();
            gm = new GuiMgmt();
            gm.init( this );
            dm.init( gm);
            parse = new Parser();
        }


        private void openFile_Click_1( object sender, RoutedEventArgs e ) {

            if ( gm.Mp.txAuftragTyp.Text.Equals("") ) {
                DialogMessage msg = new DialogMessage();
                msg.MessageText.Text = "Das Feld Auftragstyp muss ausgefüllt sein!";
                msg.Owner = this;
                msg.ShowDialog();
                return; 
            }
            
            string f = dm.loadMuenFile();
            XmlDocument doc = dm.createXml( f );
            dm.saveMuenFile( doc );
        }
    }
}
