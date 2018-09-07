using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Xml.Linq;
using Outlook = Microsoft.Office.Interop.Outlook;
using Office = Microsoft.Office.Core;
using System.Windows.Forms;
using Microsoft.Graph;
using System.Diagnostics;

namespace OutlookEventsAddIn
{
    public partial class ThisAddIn
    {
        Outlook.NameSpace outlookNameSpace;
        //Outlook.MAPIFolder inbox;
        //Outlook.Items items;

        private void ThisAddIn_Startup(object sender, System.EventArgs e)
        {
            outlookNameSpace = this.Application.GetNamespace("MAPI");
            GetCurrentUserInfo();
            //registerInboxItemAdd();
            registerSentItemEvent();

            //registerUnreadMails();
            //AuthenticateGraphAPI();
        }

        private void ThisAddIn_Shutdown(object sender, System.EventArgs e)
        {
            // Note: Outlook no longer raises this event. If you have code that 
            //    must run when Outlook shuts down, see https://go.microsoft.com/fwlink/?LinkId=506785
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

        private void GetCurrentUserInfo()
        {
            Outlook.AddressEntry addrEntry =
                Application.Session.CurrentUser.AddressEntry;
            if (addrEntry.Type == "EX")
            {
                Outlook.ExchangeUser currentUser =
                    Application.Session.CurrentUser.
                    AddressEntry.GetExchangeUser();
                if (currentUser != null)
                {
                    StringBuilder sb = new StringBuilder();
                    sb.AppendLine("Name: "
                        + currentUser.Name);
                    sb.AppendLine("STMP address: "
                        + currentUser.PrimarySmtpAddress);
                    sb.AppendLine("Title: "
                        + currentUser.JobTitle);
                    sb.AppendLine("Department: "
                        + currentUser.Department);
                    sb.AppendLine("Location: "
                        + currentUser.OfficeLocation);
                    sb.AppendLine("Business phone: "
                        + currentUser.BusinessTelephoneNumber);
                    sb.AppendLine("Mobile phone: "
                        + currentUser.MobileTelephoneNumber);
                    Debug.WriteLine(sb.ToString());
                }
            }
        }

        private void AuthenticateGraphAPI()
        {
            GraphServiceClient graphClient = AuthenticationHelper.GetAuthenticatedClient();
            if (graphClient != null)
            {
                Debug.WriteLine("GraphClient available");
            }
        }
        private void registerUnreadMails()
        {
            Outlook.MAPIFolder inbox =
                this.Application.ActiveExplorer().Session.GetDefaultFolder
                (Outlook.OlDefaultFolders.olFolderInbox);

            Outlook.Items unreadItems = inbox.
                Items.Restrict("[Unread]=true");

            Console.WriteLine(unreadItems.Count);
            //MessageBox.Show(
            //    string.Format("Unread items in Inbox = {0}", unreadItems.Count));
        }

        private void registerSentItemEvent()
        {
            Outlook.MAPIFolder sentMail = outlookNameSpace.GetDefaultFolder(
                    Outlook.OlDefaultFolders.olFolderSentMail);

            Outlook.Items items = sentMail.Items;
            items.ItemAdd += new Outlook.ItemsEvents_ItemAddEventHandler(mailItemSend);
            //((Outlook.ItemEvents_10_Event)items).Send += new Outlook.ItemEvents_10_SendEventHandler(mailItemSendHandler);
            
        }

        private void mailItemSendHandler(ref bool isSend)
        {
            Console.WriteLine(isSend);
        }

        void mailItemSend(object Item)
        {
            Outlook.MailItem mailItem = (Outlook.MailItem)Item;
            string entryId = mailItem.EntryID;
            Console.WriteLine("test");
        }

        private void registerInboxItemAdd()
        {
            Outlook.MAPIFolder inbox = outlookNameSpace.GetDefaultFolder(
                    Outlook.OlDefaultFolders.olFolderInbox);

            Outlook.Items items = inbox.Items;
            items.ItemAdd +=
                new Outlook.ItemsEvents_ItemAddEventHandler(mailItemRecievedHandler);
        }

        private void mailItemRecievedHandler(object Item)
        {
            string filter = "Test Mail";
            if (Item != null && Item is Outlook.MailItem)
            {
                Outlook.MailItem mailItem = (Outlook.MailItem) Item;

                if (mailItem.MessageClass == "IPM.Note" &&
                            mailItem.Subject.ToUpper().Contains(filter.ToUpper()))
                {
                    //mail.Move(outlookNameSpace.GetDefaultFolder(
                    //    Outlook.OlDefaultFolders.olFolderJunk));
                    
                    string messageId = getHeaderValueForKey(mailItem, "message-id");
                    //string messageId1 = getMessageIdnew(mailItem);
                    string refernceMessageId = getHeaderValueForKey(mailItem, "references");

                }
                
            }
        }

        private string getHeaderValueForKey(Outlook.MailItem mailItem, string key)
        {
            const string PR_MAIL_HEADER_TAG = "http://schemas.microsoft.com/mapi/proptag/0x007D001E";
            Outlook.PropertyAccessor oPropAccessor = mailItem.PropertyAccessor;
            string headers = (string)oPropAccessor.GetProperty(PR_MAIL_HEADER_TAG);
            //Debug.WriteLine(headers);
            int keyIndex = -1;
            //becasue the key word "message-id",usually is "Message-ID",but sometimes is "Message-Id" or "message-Id"
            string lowercaseHeaders = headers.ToLower();
            if (lowercaseHeaders.Contains(key))
            {
                keyIndex = lowercaseHeaders.IndexOf(key + ":");
            }
            else
            {
                return null;
            }
            int startIndex = headers.IndexOf(@"<", keyIndex);
            int endIndex = headers.IndexOf(@">", startIndex);
            return headers.Substring(startIndex, endIndex - startIndex + 1);
        }

        public string getMessageIdnew(Outlook.MailItem mailItem)
        {
            const string CdoPR_INTERNET_MESSAGE_ID = "http://schemas.microsoft.com/mapi/proptag/0x1035001F";
            Outlook.PropertyAccessor oPropAccessor = mailItem.PropertyAccessor;
            return (string)oPropAccessor.GetProperty(CdoPR_INTERNET_MESSAGE_ID);
        }
    }
}
