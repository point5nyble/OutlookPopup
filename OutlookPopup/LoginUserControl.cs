using Newtonsoft.Json;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace OutlookPopup
{
    public partial class LoginUserControl : UserControl
    {
        public LoginUserControl()
        {
            InitializeComponent();
            //InitializeAsync();
        }
        public async void InitializeAsync()
        {
            await webViewLogin.EnsureCoreWebView2Async(null);

        }

        private async void CoreWebView2_WebResourceResponseReceived(object sender, Microsoft.Web.WebView2.Core.CoreWebView2WebResourceResponseReceivedEventArgs e)
        {
            if (e.Request.Method == "POST")
            {
                Stream stream = await e.Response.GetContentAsync();
                TextReader tr = new StreamReader(stream);
                User uDetails = JsonConvert.DeserializeObject<User>(tr.ReadToEnd());

                OutlookPopup.Properties.Settings.Default.emailId = uDetails.userCredentials.email;
                string token = OutlookPopup.Properties.Settings.Default.token = uDetails.token;

                Globals.ThisAddIn.CustomTaskPanes.RemoveAt(0);
            }

        }
    }
}
