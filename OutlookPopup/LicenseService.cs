using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading.Tasks;

namespace OutlookPopup
{
    public class LicenseService
    {
        
        public void getLicense(ClientInfo client)
        {

        }

        public static Task<bool> IsLicenseValidAsync(string emailId)
        {

            return Task.FromResult(true);
        }

        public void UpdateClientInfo(ClientInfo client)
        {

        }

        public void AddClientInfo(ClientInfo client)
        {

        }

        internal static Task<bool> HasOfflineLimitReached()
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
                return Task.FromResult(false);
            }
        }

        [DllImport("wininet.dll")]
        private extern static bool InternetGetConnectedState(out int Description, int ReservedValue);
    }
}
