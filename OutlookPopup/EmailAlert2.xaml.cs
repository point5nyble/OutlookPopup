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

namespace OutlookPopup
{
    /// <summary>
    /// Interaction logic for EmailAlert2.xaml
    /// </summary>
    public partial class EmailAlert2 : Window
    {
        public string AttachmentMessage { get; set; }
        public string AttachmentTitle { get; set; }
        public string SendBtnText { get; set; }
        public string DSendBtnText { get; set; }

        public EmailAlert2()
        {
            InitializeComponent();
            AttachmentMessage = Globals.ThisAddIn.regValues.AttachmentMessageBody;
            AttachmentTitle = Globals.ThisAddIn.regValues.AttachmentMessageTitle;
            SendBtnText = Globals.ThisAddIn.regValues.SendButttonText;
            DSendBtnText = Globals.ThisAddIn.regValues.DSendButtonText;
            this.DataContext = this;
        }

        private void DontSend_Click(object sender, RoutedEventArgs e)
        {
            this.Close();
            Globals.ThisAddIn.hasToSend = false;
        }

        private void Button_Click(object sender, RoutedEventArgs e)
        {
            this.Close();
            //Globals.ThisAddIn.Application.ItemSend -= new Microsoft.Office.Interop.Outlook.ApplicationEvents_11_ItemSendEventHandler(Globals.ThisAddIn.Item_Send);
            //Globals.ThisAddIn.mailItem.Send();
            //Globals.ThisAddIn.Application.ItemSend += new Microsoft.Office.Interop.Outlook.ApplicationEvents_11_ItemSendEventHandler(Globals.ThisAddIn.Item_Send);
            Globals.ThisAddIn.hasToSend = true;
        }
    }
}
