using System.Windows;
using System.Windows.Controls;
using System.Windows.Input;

namespace MuenImport.Gui
{
    /// <summary>
    /// Interaktionslogik für DialogMessage.xaml
    /// </summary>
    public partial class DialogMessage : Window {
        public DialogMessage() {
            InitializeComponent();
        }

        private void okButton_Click(object sender, RoutedEventArgs e) {
            this.DialogResult = true;
        }
    }
}
