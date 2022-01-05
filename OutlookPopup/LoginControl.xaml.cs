using Newtonsoft.Json;
using System;
using System.Collections.Generic;
using System.IO;
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
    /// Interaction logic for LoginControl.xaml
    /// </summary>
    public partial class LoginControl : Window
    {
        public LoginControl()
        {
            InitializeComponent();
            InitializeAsync();
        }

        public async void InitializeAsync()
        {
            await webView.EnsureCoreWebView2Async(null);


        }
        private async void CoreWebView2_WebResourceResponseReceived(object sender, Microsoft.Web.WebView2.Core.CoreWebView2WebResourceResponseReceivedEventArgs e)
        {
            if (e.Request.Method == "POST")
            {
                Stream stream = await e.Response.GetContentAsync();
                TextReader tr = new StreamReader(stream);
                User uDetails = JsonConvert.DeserializeObject<User>(tr.ReadToEnd());

                OutlookPopup.Properties.Settings.Default.emailId= uDetails.userCredentials.email;
                string token = OutlookPopup.Properties.Settings.Default.token= uDetails.token;

                this.Close();
            }

        }
    }
}
