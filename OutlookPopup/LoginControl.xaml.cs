using Microsoft.Web.WebView2.Core;
using Microsoft.Web.WebView2.Wpf;
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

            //webView.Source = new Uri( "https://www.microsoft.com");
            webView.Initialized += WebView_Initialized;
            //webView.BringIntoView();
            //InitializeAsync(webView, "");
        }

        private async void  WebView_Initialized(object sender, EventArgs e)
        {
            await InitializeAsync();
           // await InitializeAsync(webView, "");
        }

        //private async void WebView_Initialized(object sender, EventArgs e)
        //{

        //}
        public async Task InitializeAsync(WebView2 wv, string webCacheDir = "")
        {
            CoreWebView2EnvironmentOptions options = null;
            string tempWebCacheDir = string.Empty;
            CoreWebView2Environment webView2Environment = null;

            //set value
            tempWebCacheDir = webCacheDir;

            if (String.IsNullOrEmpty(tempWebCacheDir))
            {
                //get fully-qualified path to user's temp folder
                tempWebCacheDir = System.IO.Path.GetTempPath();

                tempWebCacheDir = System.IO.Path.Combine(tempWebCacheDir, System.Guid.NewGuid().ToString("N"));
            }

            //use with WebView2 FixedVersionRuntime
            webView2Environment = await CoreWebView2Environment.CreateAsync(@".\Microsoft.WebView2.FixedVersionRuntime.88.0.705.81.x86", tempWebCacheDir, options);

            //webView2Environment = await CoreWebView2Environment.CreateAsync(@"C:\Program Files (x86)\Microsoft\Edge Dev\Application\90.0.810.1", tempWebCacheDir, options);
            //webView2Environment = await CoreWebView2Environment.CreateAsync(null, tempWebCacheDir, options);

            //wait for CoreWebView2 initialization
            await wv.EnsureCoreWebView2Async(webView2Environment);

        }

        public async Task InitializeAsync()
        {
            string folder = @"C:\temp";//Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData);
            var env = await CoreWebView2Environment.CreateAsync(null, folder);
            await webView.EnsureCoreWebView2Async();

            webView.CoreWebView2.Navigate("https://www.microsoft.com");
            webView.CoreWebView2.WebResourceResponseReceived += CoreWebView2_WebResourceResponseReceived;

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
