using Microsoft.Win32;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace OutlookPopup
{
    public class RegValues
    {

        private static readonly log4net.ILog log =
                log4net.LogManager.GetLogger(System.Reflection.MethodBase.GetCurrentMethod().DeclaringType);

        public void createRegistryKeys()
        {
            Microsoft.Win32.RegistryKey key;
            key = Microsoft.Win32.Registry.LocalMachine.CreateSubKey("Software\\Zepto\\OutlookMailPrompt");
            key.SetValue("AttachmentPromptEnabled",1);
            key.SetValue("ExternalRecpPromptEnabled", 1);
            key.SetValue("AttachmentMessageBody", "This email will be sent to an external party. Please ensure sensitive and/or confidential information has been encrypted as per Great Eastern IS Policy.");
            key.SetValue("AttachmentMessageTitle", "Attention!!");
            key.SetValue("ExternalRecpMessageTitle", "Warning!! External Recepient(s).");
            key.SetValue("ExternalRecpMessageBody", "This email will be sent to an external party.!&#10;&#13;Please validate the intended recipient(s) ensure this email complies with Section 133(1) of Financial Services Act 2013 and Section 9(1) of Personal Data Protection Act 2010.!&#10;&#13;!&#10;&#13;Do you want to proceed ? ");
            key.SetValue("SendButttonText", "Send");
            key.SetValue("DSendButtonText", "Don't Send");
            key.SetValue("AcceptedDomains", "zepto.com.my,outlook.com,rhingle.com");

            key.Close();
            readRegistryKeys();
        }

        public void readRegistryKeys()
        {

           // Microsoft.Win32.RegistryKey key = Microsoft.Win32.Registry.LocalMachine.OpenSubKey("Software\\CTC\\OutlookMailPrompt");
             try
             {

                using (var view32 = RegistryKey.OpenBaseKey(RegistryHive.LocalMachine, RegistryView.Default))
                {
                    using (var key = view32.OpenSubKey("Software\\Zepto\\OutlookMailPrompt", false))
                    {
                        // actually accessing Wow6432Node 
                        //if it does exist, retrieve the stored values  
                   
                        if (key != null)
                        {
                            //var val = key.GetValue("AttachmentPromptEnabled");
                            log.Info("Registry key found,Reading the values" + key.ToString());
                            AttachmentPromptEnabled = Convert.ToInt32(key.GetValue("AttachmentPromptEnabled"));
                            ExternalRecpPromptEnabled = Convert.ToInt32(key.GetValue("ExternalRecpPromptEnabled"));
                            AttachmentMessageTitle = key.GetValue("AttachmentMessageTitle").ToString();
                            AttachmentMessageBody = key.GetValue("AttachmentMessageBody").ToString();
                            ExternalRecpMessageTitle = key.GetValue("ExternalRecpMessageTitle").ToString();
                            ExternalRecpMessageBody = key.GetValue("ExternalRecpMessageBody").ToString();
                            SendButttonText = key.GetValue("SendButttonText").ToString();
                            DSendButtonText = key.GetValue("DSendButttonText").ToString();
                            AcceptedDomains = key.GetValue("AcceptedDomainList").ToString();
                            //EndpointURL = key.GetValue("EndpointURL").ToString();
                            LicenseId = key.GetValue("SoftwareUniqueIdentifier").ToString();
                        }
                        else
                        {                           
                            log.Info("Registry Values not found, Software\\Zepto\\OutlookMailPromp");
                            //createRegistryKeys();
                        }
                   
                    }
                }
                   
             }
             catch (Exception ex)
             {
                 log.Error("Exception while reading registry values.");
                 //throw ex;
             }
   
        }
        public  int AttachmentPromptEnabled { get; set; }
        public  int ExternalRecpPromptEnabled { get; set; }
        public  string AttachmentMessageTitle { get; set; }
        public  string AttachmentMessageBody { get; set; }
        public  string ExternalRecpMessageTitle { get; set; }

        public   string ExternalRecpMessageBody { get; set; }

        public  string SendButttonText { get; set; }
        public  string DSendButtonText { get; set; }

        public string AcceptedDomains { get; set; }

        public string EndpointURL { get; set; }

        public string LicenseId { get; set; }
    }
}
