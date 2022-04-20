using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Xml.Linq;
using Outlook = Microsoft.Office.Interop.Outlook;
using Office = Microsoft.Office.Core;
using System.Net;
using System.DirectoryServices;
using log4net;
using System.Windows.Interop;
using System.Threading.Tasks;
using System.Threading;
using System.Windows;
//using Microsoft.Exchange.WebServices.Data;
namespace OutlookPopup
{
    public partial class ThisAddIn
    {

        
        public RegValues regValues = new RegValues();
       
        private static readonly log4net.ILog log = 
                        log4net.LogManager.GetLogger(System.Reflection.MethodBase.GetCurrentMethod().DeclaringType);
        
        private  void ThisAddIn_Startup(object sender, System.EventArgs e)
        {
            //call license service
            IsLicenseActiveEx();            
            this.Application.ItemSend += new Outlook.ApplicationEvents_11_ItemSendEventHandler(Item_Send);
            log4net.Config.XmlConfigurator.Configure();
            log.Info("Plugin Loaded Successfully");

        }

       
        
        bool isTokenValid = false;
        public LicenseStatus licenseStatus { get;private set; }
        bool hasOfflineLimitReached;

        private async void IsLicenseActiveEx()
        {
            string email = Properties.Settings.Default.emailId;
            string token = Properties.Settings.Default.token;
            
            isTokenValid = await LicenseService.IsTokenValid(email, token);
            if (!isTokenValid)
            {
                var respose = await LicenseService.Login("jabez@zepto.international", "Jabez@123");
                if (respose.success)
                {
                    Properties.Settings.Default.token = respose.token;
                    Properties.Settings.Default.Save();
                }
            }
            // call service now to check on license
            token = Properties.Settings.Default.token;
            
            //readRegistry
            if (regValues.SendButttonText == null)
                regValues.readRegistryKeys();
            string licenseid = regValues.LicenseId;
            if (licenseid==null)
            {
                licenseStatus = new LicenseStatus { IsValid = false, Message = "No Valid License assigned, registry missing." };
                return;
            }
            string loggedinUserEmail = GetCurrentUserEmail();
            var licenseResponse = await LicenseService.IsValidLicenseID(loggedinUserEmail,licenseid, token);
            if (!licenseResponse)
            {
                licenseStatus = new LicenseStatus { IsValid = false, Message = "No Valid License assigned, please contact admin." };
                hasOfflineLimitReached = await LicenseService.HasOfflineLimitReachedAsync();
            }
            else
            {
                licenseStatus = new LicenseStatus { IsValid = true, Message = "Valid License Found." };               
            }
                
        }

        private string GetOutlookVersion()
        {
            return Globals.ThisAddIn.Application.Version;
        }

        private string GetMachineOS()
        {
            OperatingSystem os = Environment.OSVersion;

            if (os.Version.Major > 6)
            {
                return "Win 10";
            }
            else if (os.Version.Minor == 2)
            {
                return "Win 8/8.1";
            }
            else if (os.Version.Minor == 1)
            {
                return "Win 7";
            }
            else
                return "Lower than Win 7";

        }

        private string GetCurrentUserEmail()
        {
            log.Info("Get the Logged in user's Email id.");
            Outlook.AddressEntry addrEntry =
                Application.Session.CurrentUser.AddressEntry;
            if (addrEntry.Type == "EX")
            {
                Outlook.ExchangeUser currentUser =
                    Application.Session.CurrentUser.
                    AddressEntry.GetExchangeUser();
                if (currentUser != null)
                {
                    log.Info($"Logged in user's Email id:{currentUser.PrimarySmtpAddress}");
                    return currentUser.PrimarySmtpAddress;
                    //sb.AppendLine("Title: "
                    //    + currentUser.JobTitle);
                    //sb.AppendLine("Department: "
                    //    + currentUser.Department);
                    //sb.AppendLine("Location: "
                    //    + currentUser.OfficeLocation);
                    //sb.AppendLine("Business phone: "
                    //    + currentUser.BusinessTelephoneNumber);
                    //sb.AppendLine("Mobile phone: "
                    //    + currentUser.MobileTelephoneNumber);
                    
                }
                else
                {
                    log.Info($"Failed to get the Logged in user's Email id.Random email id will be configured as a placeholder.");
                    return new Guid().ToString() +"@Zepto.international";
                }
                    
            }
            else
            {
                string email= Application.Session.CurrentUser.AddressEntry.Address;
                log.Info($"User is no an Exchange user, user's Email id:{email}");
                return email;
            }
                
        }

        private void ThisAddIn_Shutdown(object sender, System.EventArgs e)
        {
            this.Application.ItemSend -= new Outlook.ApplicationEvents_11_ItemSendEventHandler(Item_Send);
        }

        public bool hasToSend=false;

        //private LoginUserControl myUserControl1;
        //private Microsoft.Office.Tools.CustomTaskPane myCustomTaskPane;

        const string PR_SMTP_ADDRESS = "http://schemas.microsoft.com/mapi/proptag/0x39FE001E";
        public void Item_Send(object Item, ref bool Cancel)
        {
            HookUnhookItemSendEvent();
            if (regValues.SendButttonText==null)
                regValues.readRegistryKeys();

            if (regValues.ExternalRecpPromptEnabled == 1)
            {
                if (Item is Outlook.MailItem)
                {
                    Outlook.PropertyAccessor pa;
                    Outlook.MailItem mailItem;
                    mailItem = Item as Outlook.MailItem;
                    foreach (Outlook.Recipient recp in mailItem.Recipients)
                    {

                        Outlook.AddressEntry addEntry = recp.AddressEntry;
                        pa = recp.PropertyAccessor;
                        //if (addEntry.AddressEntryUserType != Outlook.OlAddressEntryUserType.olExchangeUserAddressEntry)
                        try
                        {
                            //var domain = addEntry.Address.Split('@');
                            string domain = (string)pa.GetProperty(PR_SMTP_ADDRESS);

                            if (!regValues.AcceptedDomains.Contains(domain.Split('@')[1]))
                            {
                                //initialize logging

                                log.Info("Plugin was Loaded Successfully and Recipient list contains external  user.");

                                log.Info("Registry Values read");

                                log.Info("External User found,Warning Window should be shown");
                                WarningMessage window = new WarningMessage();

                                //Set Popup as child of the active window of Outlook
                                Outlook.Inspector activeWindow = Globals.ThisAddIn.Application.ActiveWindow() as Outlook.Inspector;
                                if (activeWindow != null)
                                {
                                    IntPtr outlookHwnd = new OfficeWin32Window(activeWindow).Handle;
                                    WindowInteropHelper wih = new WindowInteropHelper(window);
                                    wih.Owner = outlookHwnd;
                                }
                                else
                                {
                                    Outlook.Explorer activeExplorer = Globals.ThisAddIn.Application.ActiveWindow() as Outlook.Explorer;
                                    if (activeExplorer != null)
                                    {
                                        IntPtr outlookHwnd = new OfficeWin32Window(activeExplorer).Handle;
                                        WindowInteropHelper wih = new WindowInteropHelper(window);
                                        wih.Owner = outlookHwnd;
                                    }

                                }
                                window.ShowActivated = true;

                                window.ShowDialog();

                                if (hasToSend)
                                {
                                    //Cnaned on 16tdec2019
                                    //showAttachmentPopup(mailItem.Attachments);
                                }

                                if (!hasToSend)
                                {
                                    Cancel = true;
                                }


                                break;
                            }
                            else
                                log.Info(domain + " User is not an external user.");
                        }
                        catch (Exception ex)
                        {

                            log.Fatal(ex.Message + "Some Registry Values/Keys missing.");
                        }

                    }

                }
                else if (Item is Outlook.MeetingItem)
                {
                    Outlook.PropertyAccessor pa;
                    Outlook.MeetingItem aptItem;
                    aptItem = Item as Outlook.MeetingItem;
                    foreach (Outlook.Recipient recp in aptItem.Recipients)
                    {

                        Outlook.AddressEntry addEntry = recp.AddressEntry;
                        pa = recp.PropertyAccessor;
                        //if (addEntry.AddressEntryUserType != Outlook.OlAddressEntryUserType.olExchangeUserAddressEntry)
                        try
                        {
                            //var domain = addEntry.Address.Split('@');
                            string domain = (string)pa.GetProperty(PR_SMTP_ADDRESS);

                            if (!regValues.AcceptedDomains.Contains(domain.Split('@')[1]))
                            {
                                //initialize logging

                                log.Info("Plugin was Loaded Successfully and Recipient list contains external  user.");

                                log.Info("Registry Values read");

                                log.Info("External User found,Warning Window should be shown");
                                WarningMessage window = new WarningMessage();
                                //Alert window = new Alert();
                                //Set Popup as child of the active window of Outlook
                                Outlook.Inspector activeWindow = Globals.ThisAddIn.Application.ActiveWindow() as Outlook.Inspector;
                                if (activeWindow != null)
                                {
                                    IntPtr outlookHwnd = new OfficeWin32Window(activeWindow).Handle;
                                    WindowInteropHelper wih = new WindowInteropHelper(window);
                                    wih.Owner = outlookHwnd;
                                }
                                else
                                {
                                    Outlook.Explorer activeExplorer = Globals.ThisAddIn.Application.ActiveWindow() as Outlook.Explorer;
                                    if (activeExplorer != null)
                                    {
                                        IntPtr outlookHwnd = new OfficeWin32Window(activeExplorer).Handle;
                                        WindowInteropHelper wih = new WindowInteropHelper(window);
                                        wih.Owner = outlookHwnd;
                                    }

                                }
                                window.ShowDialog();
                                window.Activate();

                                if (!hasToSend)
                                {
                                    Cancel = true;
                                }

                                break;
                            }
                            else
                                log.Info(domain + " User is not an external user.");
                        }
                        catch (Exception ex)
                        {
                            log.Fatal(ex.Message + "Some Registry Values/Keys missing.");
                        }

                    }
                }
            }
            else
                Cancel = true;
        }
        private void HookUnhookItemSendEvent()
        {
                        
            log.Info("Item Send event hooked");
           
            if (licenseStatus == null)
            {
                IsLicenseActiveEx();
            }
            
            if (licenseStatus.IsValid)
            {
                log.Info("Valid License, Item send event will continue hooked");
                //this.Application.ItemSend += new Outlook.ApplicationEvents_11_ItemSendEventHandler(Item_Send);
                hasToSend = true;
            }
            else
            {
                if (!hasOfflineLimitReached)
                {
                    log.Info("InValid License but within Offline Limit, Item send event will continue hooked");
                    //this.Application.ItemSend += new Outlook.ApplicationEvents_11_ItemSendEventHandler(Item_Send);
                    hasToSend = true;
                }
                else
                {
                    log.Info("Invalid License, Item send event unhooked");
                    this.Application.ItemSend -= new Outlook.ApplicationEvents_11_ItemSendEventHandler(Item_Send);
                    MessageBox.Show(licenseStatus.Message+ "\n\nSent emails will no longer be monitored.", "Outlook Popup");
                    hasToSend = false;

                }

            }
           
        }
        void showAttachmentPopup(Outlook.Attachments attchments )
        {
            bool showPopup = false;
            Outlook.Inspector openWindow = Globals.ThisAddIn.Application.ActiveInspector() as Outlook.Inspector;

            if (openWindow != null)
            {
                if (attchments != null)
                {
                    if (attchments.Count > 0 && Globals.ThisAddIn.regValues.AttachmentPromptEnabled == 1)
                    {
                        hasToSend = false;
                        foreach (Microsoft.Office.Interop.Outlook.Attachment item in attchments)
                        {
                            string value =(string) item.PropertyAccessor.GetProperty("http://schemas.microsoft.com/mapi/proptag/0x3712001E");

                            if ("" == value)
                            {
                                showPopup = true;
                                break;
                            }
                            else
                                showPopup = false;

                        }


                        if (showPopup)
                        {
                            EmailAlert2 alert = new EmailAlert2();

                            //Set Popup as child of the active window of Outlook
                            Outlook.Inspector activeWindow = Globals.ThisAddIn.Application.ActiveWindow() as Outlook.Inspector;
                            if (activeWindow != null)
                            {

                                IntPtr outlookHwnd = new OfficeWin32Window(activeWindow).Handle;
                                WindowInteropHelper wih = new WindowInteropHelper(alert);
                                wih.Owner = outlookHwnd;
                            }

                            alert.ShowDialog();
                        }
                        else
                        {
                            Globals.ThisAddIn.hasToSend = true;
                        }
                    }
                    else
                    {
                        
                        Globals.ThisAddIn.hasToSend = true;
                    }
                }
                
            }
            else
            {
                #region Explorer

                Outlook.Explorer activeExplorer = Globals.ThisAddIn.Application.ActiveWindow() as Outlook.Explorer;
                if (activeExplorer != null)
                {
                    if (attchments != null)
                    {
                        if (attchments.Count > 0 && Globals.ThisAddIn.regValues.AttachmentPromptEnabled == 1)
                        {
                            hasToSend = false;
                            foreach (Microsoft.Office.Interop.Outlook.Attachment item in attchments)
                            {
                                string value =(string) item.PropertyAccessor.GetProperty("http://schemas.microsoft.com/mapi/proptag/0x3712001E");

                                if ("" == value)
                                {
                                    showPopup = true;
                                    break;
                                }
                                else
                                    showPopup = false;

                            }


                            if (showPopup)
                            {
                                EmailAlert2 alert = new EmailAlert2();

                                IntPtr outlookHwnd = new OfficeWin32Window(activeExplorer).Handle;
                                WindowInteropHelper wih = new WindowInteropHelper(alert);
                                wih.Owner = outlookHwnd;
                                alert.ShowDialog();
                            }
                            else
                            {
                                Globals.ThisAddIn.hasToSend = true;
                            }
                        }
                        else
                        {
                            Globals.ThisAddIn.hasToSend = true;
                        }
                    }

                }
                #endregion
            }
        }

        #region VSTO generated code

        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InternalStartup()
        {
            this.Startup += new System.EventHandler(ThisAddIn_Startup);
            this.Shutdown += new System.EventHandler(ThisAddIn_Shutdown);
        }
        
        #endregion
    }
}
