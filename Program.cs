using Microsoft.Office.Interop.Outlook;
using static System.Net.Mime.MediaTypeNames;
using System.Reflection;
using System;
using Application = Microsoft.Office.Interop.Outlook.Application;
namespace OutlookConnection
{
    class Program
    {
        static void Main(string[] args)
        {
           
            // Create an instance of the Outlook application
            Microsoft.Office.Interop.Outlook.Application outlookApp = new Application();

            // Get the MAPI namespace
            NameSpace outlookNamespace = outlookApp.GetNamespace("MAPI");

            outlookNamespace.SendAndReceive(true);
            int t2 = outlookNamespace.Accounts.Count;
            string h1 = outlookNamespace.CurrentProfileName;
            // Log in to Outlook
            outlookNamespace.Logon(h1, Missing.Value, false, true);

            // Access the default Inbox folder
            MAPIFolder inboxFolder = outlookNamespace.Folders["Личные папки"];


            MAPIFolder folder = inboxFolder.Folders["Входящие"].Folders["support"]; //.GetFirst();

            try
            {

                // Iterate through the items in the Inbox folder
                foreach (object item in folder.Items)
                {
                    if (item is MailItem)
                    {
                        Console.WriteLine("=====================================================================================");
                        Console.WriteLine("=====================================================================================");
                        Console.WriteLine("=====================================================================================");
                        MailItem mailItem = (MailItem)item;
                        string hhhh = mailItem.ConversationIndex;
                        //string hhhh1 = mailItem.ConversationID;
                        if (mailItem.Recipients.Count > 0)
                        {
                            for (int y = 0; y < mailItem.Recipients.Count; y++)
                            {
                                string jj = mailItem.Recipients[y + 1].Address;
                                Console.WriteLine("Recipients" + y.ToString() + ": " + jj);
                            }
                        }
                        Console.WriteLine("To: " + mailItem.To + "\n Subject: " + mailItem.Subject + "\n recieved: " + mailItem.ReceivedTime + " created:" + mailItem.CreationTime + " attachments: " + mailItem.Attachments.Count);
                        Console.WriteLine("=====================================================================================");
                        if (mailItem.Body != null)
                            Console.WriteLine(" ++++ body: " + mailItem.Body.Normalize().ToString().Trim());
                        Console.WriteLine("=====================================================================================");
                        if (mailItem.Attachments.Count > 0)
                        {
                            for (int y = 0; y < mailItem.Attachments.Count; y++)
                            {
                                try
                                {
                                    Console.WriteLine(" ----- attachment" + y.ToString() + " fname: " + mailItem.Attachments[y + 1].FileName);
                                }
                                catch { }
                            }
                        }
                    }
                }

            }
            catch (System.Exception ex)
            {
                int gg = 1;
            }

            // Log off from Outlook
            outlookNamespace.Logoff();
            Console.ReadKey();
        }
    }
}
