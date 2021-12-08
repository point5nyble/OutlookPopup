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
//using Microsoft.Exchange.WebServices.Data;
namespace OutlookPopup
{
    public partial class ThisAddIn
    {

        //ExchangeService service = new ExchangeService(ExchangeVersion.Exchange2010_SP1);

        public RegValues regValues = new RegValues();
       // private static readonly log4net.ILog log = log4net.LogManager.GetLogger(typeof(ThisAddIn));
        private static readonly log4net.ILog log = 
                        log4net.LogManager.GetLogger(System.Reflection.MethodBase.GetCurrentMethod().DeclaringType);
        
        private void ThisAddIn_Startup(object sender, System.EventArgs e)
        {
            //call license service
            //IsLicenseActive();
            this.Application.ItemSend += new Outlook.ApplicationEvents_11_ItemSendEventHandler(Item_Send);
            log4net.Config.XmlConfigurator.Configure();
            log.Info("Plugin Loaded Successfully");
        }

        private async void IsLicenseActive()
        {
            string emailId = Globals.ThisAddIn.Application.Session.CurrentUser.Address;
            bool isActive = await LicenseService.IsLicenseValidAsync(emailId);
            bool hasOfflineLimitReached = await LicenseService.HasOfflineLimitReached();
            if (isActive)
            {
                this.Application.ItemSend += new Outlook.ApplicationEvents_11_ItemSendEventHandler(Item_Send);
            }
            else
            {
                if (!hasOfflineLimitReached)
                {
                    this.Application.ItemSend += new Outlook.ApplicationEvents_11_ItemSendEventHandler(Item_Send);
                }
            }
        }

        private void ThisAddIn_Shutdown(object sender, System.EventArgs e)
        {
            this.Application.ItemSend -= new Outlook.ApplicationEvents_11_ItemSendEventHandler(Item_Send);
        }

        //private void InitiateService()
        //{
        //    // Set specific credentials.
        //    service.Credentials = new NetworkCredential("aplifeisgreat\\pvmexsvcr02", "QaZx8we$");

        //    // Look up the user's EWS endpoint by using Autodiscover.
        //    service.Url = new Uri("https://exchange.greateasternlife.com/EWS/Exchange.asmx");
        //}
        
        public bool hasToSend=false;

        //public static DirectoryEntry GetDirectoryEntry()
        //{
        //    try
        //    {
        //        DirectoryEntry entryRoot = new DirectoryEntry("LDAP://zeptoc.onmicrosoft.com");
        //        string Domain = (string)entryRoot.Properties["defaultNamingContext"][0];

        //        DirectoryEntry de = new DirectoryEntry();

        //        de.Path = "LDAP://" + Domain;
        //        de.AuthenticationType = AuthenticationTypes.Secure;

        //        return de;
        //    }
        //    catch
        //    {
        //        return null;
        //    }

        //}


        //private void acceptedDomainList()
        //{

        //    try
        //    {
        //        DirectoryEntry rdRootDSE = GetDirectoryEntry();
        //        //rdRootDSE.Path = "LDAP://CN=henrylim,DC=APCPR03A001,DC=prod,DC=outlook,DC=com";
        //        DirectorySearcher cfConfigPartitionSearch = new DirectorySearcher(rdRootDSE);
        //        cfConfigPartitionSearch.Filter = "(objectClass=msExchAcceptedDomain)";
        //        cfConfigPartitionSearch.SearchScope = SearchScope.Subtree;
        //        SearchResultCollection srSearchResults = cfConfigPartitionSearch.FindAll();
        //        foreach (SearchResult srSearchResult in srSearchResults)
        //        {
        //            DirectoryEntry acDomain = srSearchResult.GetDirectoryEntry();
        //            Console.WriteLine("Domain : " + acDomain.Properties["msExchAcceptedDomainName"].Value.ToString());

        //        }
        //    }
        //    catch (Exception ex )
        //    {
                
        //        throw ex;
        //    }
            
        //}

        const string PR_SMTP_ADDRESS = "http://schemas.microsoft.com/mapi/proptag/0x39FE001E";
        public void Item_Send(object Item, ref bool Cancel)
        {

            log.Info("Item Send event hooked");
            
            if  (regValues.SendButttonText==null)
                regValues.readRegistryKeys();

            if (regValues.ExternalRecpPromptEnabled==1)
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
                            var domain = pa.GetProperty(PR_SMTP_ADDRESS);

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
                                    if (activeExplorer!=null)
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
                else if(Item is Outlook.MeetingItem)
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
                            var domain = pa.GetProperty(PR_SMTP_ADDRESS);

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
                    
            }
           
            
        }

        void showAttachmentPopup(Outlook.Attachments attchments )
        {
            bool showPopup = false;
            Outlook.Inspector openWindow = Globals.ThisAddIn.Application.ActiveInspector() as Outlook.Inspector;

            if (openWindow != null)
            {
                //#region inspector
                //if (openWindow.CurrentItem is Outlook.MailItem)
                //{
                //    Outlook.MailItem mItem = (Outlook.MailItem)openWindow.CurrentItem;
                //    attchments = mItem.Attachments;
                //}
                //else if (openWindow.CurrentItem is Outlook.MeetingItem)
                //{
                //    Outlook.MeetingItem mItem = (Outlook.MeetingItem)openWindow.CurrentItem;
                //    attchments = mItem.Attachments;
                //}
                //else if (openWindow.CurrentItem is Outlook.AppointmentItem)
                //{
                //    Outlook.AppointmentItem mItem = (Outlook.AppointmentItem)openWindow.CurrentItem;
                //    attchments = mItem.Attachments;
                //}

                if (attchments != null)
                {
                    if (attchments.Count > 0 && Globals.ThisAddIn.regValues.AttachmentPromptEnabled == 1)
                    {
                        hasToSend = false;
                        foreach (Microsoft.Office.Interop.Outlook.Attachment item in attchments)
                        {
                            string value = item.PropertyAccessor.GetProperty("http://schemas.microsoft.com/mapi/proptag/0x3712001E");

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
                                string value = item.PropertyAccessor.GetProperty("http://schemas.microsoft.com/mapi/proptag/0x3712001E");

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
