using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Net;
using System.Net.Http;
using System.Runtime.InteropServices;
using System.Text;
using System.Text.Json;
using System.Threading.Tasks;
using System.Windows;

namespace OutlookPopup
{
    public class LicenseService
    {
     

        public static async Task<LicenseStatus> IsLicenseValidAsync(ClientInfo info, string token)
        {

            var json = JsonSerializer.Serialize(info);
            var data = new StringContent(json, Encoding.UTF8, "application/json");

            var url = OutlookPopup.Properties.Settings.Default.ServerAddress + "api/checkLicense";
            using (var client = new HttpClient())
            {
                client.DefaultRequestHeaders.Add("x-access-token", token);
                var response = await client.PostAsync(url, data);
                string result = response.Content.ReadAsStringAsync().Result;
                if (result=="No valid user found")
                {

                    return new LicenseStatus { Message = "No License Assigned, Please contact your administrator.", IsValid = false };
                    
                }
                else if(result.Contains("License is expired"))
                {
                    return new LicenseStatus { Message = "License expired, Please contact your administrator.", IsValid = false };
                    
                }
            }
            return new LicenseStatus { Message = "Valid License", IsValid = true };
            
        }

        internal static Task<bool> HasOfflineLimitReachedAsync()
        {
            int Desc;
            //Write to file
            string filePath = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData),
                "OfflineTracker.dat");

            InternetGetConnectedState(out Desc, 0);
            if (Desc==0)
            {
                
                string content = string.Empty;
                if (!File.Exists(filePath))
                {
                    File.Create(filePath);
                }
                using (StreamReader outFile=new StreamReader(filePath))
                {
                    content = outFile.ReadLine();
                }
                if (content.Length==30)
                {
                    return Task.FromResult(true);
                }
                else
                {
                    using (StreamWriter outputFile = new StreamWriter(filePath))
                    {
                        outputFile.WriteLineAsync(content+'F');
                    }
                    return Task.FromResult(false);
                }
                
            }
            else
            {
                if (File.Exists(filePath))
                {
                    File.Delete(filePath);
                }
                return Task.FromResult(true);
            }
        }

        internal static async Task<bool> IsTokenValid(string email, string token)
        {
            var json = JsonSerializer.Serialize(new ClientInfo { EmailId=email          
            
            });
            var data = new StringContent(json, Encoding.UTF8, "application/json");
                
            var url = OutlookPopup.Properties.Settings.Default.ServerAddress + "api/checkLicense";
            System.Net.ServicePointManager.SecurityProtocol = SecurityProtocolType.Tls | SecurityProtocolType.Tls11 | SecurityProtocolType.Tls12;
            using (var client = new HttpClient())
            {
                client.DefaultRequestHeaders.Add("x-access-token", token);
                var response = await client.PostAsync(url, data);
                
                string result = response.Content.ReadAsStringAsync().Result;
                if (result == "Invalid Token")
                {
                    return false;
                }
                else
                    return true;
            }
            
        }

        [DllImport("wininet.dll")]
        private extern static bool InternetGetConnectedState(out int Description, int ReservedValue);
    }
}
