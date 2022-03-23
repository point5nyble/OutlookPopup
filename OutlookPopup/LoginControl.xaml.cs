using Microsoft.Web.WebView2.Core;
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
           
        }

        public async Task InitializeAsync()
        {
            
            var env = await CoreWebView2Environment.CreateAsync(null, "C:\\temp");
            await webView.EnsureCoreWebView2Async(env);
            var licenseURL = OutlookPopup.Properties.Settings.Default.ServerAddress;
            webView.CoreWebView2.Navigate(licenseURL);
            webView.CoreWebView2.WebResourceResponseReceived += CoreWebView2_WebResourceResponseReceived;

        }

        

        private async void CoreWebView2_WebResourceResponseReceived(object sender, Microsoft.Web.WebView2.Core.CoreWebView2WebResourceResponseReceivedEventArgs e)
        {
            if (e.Request.Method == "POST")
            {
                Stream stream = await e.Response.GetContentAsync();
                TextReader tr = new StreamReader(stream);
                User uDetails = JsonConvert.DeserializeObject<User>(tr.ReadToEnd());
                if (uDetails.userCredentials==null)
                {

                }
                else
                {
                    OutlookPopup.Properties.Settings.Default.emailId = uDetails.userCredentials.email;
                    OutlookPopup.Properties.Settings.Default.token = uDetails.token;
                    OutlookPopup.Properties.Settings.Default.Save();
                    this.Close();
                }
                
                
            }

        }

        private async void Window_Loaded(object sender, RoutedEventArgs e)
        {
            await InitializeAsync();
        }
    }
}
