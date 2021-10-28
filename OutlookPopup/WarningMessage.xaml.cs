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
using System.Windows.Interop;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Navigation;
using System.Windows.Shapes;
using Outlook = Microsoft.Office.Interop.Outlook;
namespace OutlookPopup
{
    /// <summary>
    /// Interaction logic for WarningMessage.xaml
    /// </summary>
    public partial class WarningMessage : Window
    {

        public string ExternalRecpMessageBody { get; set; }
        public string ExternalRecpMessageTitle { get; set; }
        public string SendBtnText { get; set; }

        public string DSendBtnText { get; set; }
        
        public WarningMessage()
        {
            InitializeComponent();
            this.DataContext = this;
            this.Activate();
            DontSend.Focusable = true;
            DontSend.Focus();
            ExternalRecpMessageBody = Globals.ThisAddIn.regValues.ExternalRecpMessageBody;
            ExternalRecpMessageTitle = Globals.ThisAddIn.regValues.ExternalRecpMessageTitle;
            SendBtnText = Globals.ThisAddIn.regValues.SendButttonText;
            DSendBtnText = Globals.ThisAddIn.regValues.DSendButtonText;
            
        }

        private void Button_Click(object sender, RoutedEventArgs e)
        {
            this.Close();
            Globals.ThisAddIn.hasToSend = true;
        }

        private void DontSend_Click(object sender, RoutedEventArgs e)
        {
            this.Close();
            Globals.ThisAddIn.hasToSend = false;
        }

        private void Window_Activated(object sender, EventArgs e)
        {
            DontSend.Focus();
        }
    }
}
