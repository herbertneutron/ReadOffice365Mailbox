using System;
using System.IO;
using Microsoft.Exchange.WebServices.Data;
using System.Data.SqlClient;
using System.Collections.Generic;

namespace ReadOffice365Mailbox
{
    class Program
    {
        private const string SaveFolder = "D:\\Emails\\";
        private const string MappingUrl = "http://x.y.z./files/";

        static void Main(string[] args)
        {
            Console.WriteLine("Welcome");
            Console.WriteLine($"Default folder to save attachments is {SaveFolder}. Would you like to change it? (Y/N)");
            string input = Console.ReadLine().ToLower();
            if(input=="y")
            {
                Console.WriteLine("Please enter custom folder path.Make sure you have permissions to write in that folder");
                string path = Console.ReadLine();
                if(!Directory.Exists(path))
                    Console.WriteLine("Invalid path. Program will continue using default folder");
            }

            ExchangeService _service;

            try
            {
                Console.WriteLine("Registering Exchange connection");

                _service = new ExchangeService
                {
                    Credentials = new WebCredentials("mailbox@server.com", "password")
                };
            }
            catch
            {
                Console.WriteLine("new ExchangeService failed. Press enter to exit:");
                return;
            }

            // This is the office365 webservice URL
            _service.Url = new Uri("https://outlook.office365.com/EWS/Exchange.asmx");

            // Prepare seperate class for writing email to the database
            try
            {
                Write2DB db = new Write2DB();

                Console.WriteLine("Reading mail");

                // Read 100 mails
                foreach (EmailMessage email in _service.FindItems(WellKnownFolderName.Inbox, new ItemView(100)))
                {
                    if (email.HasAttachments)
                    {
                        List<string> savedFiles = DownloadAttachments(_service, email.Id, SaveFolder);
                    }
                    db.Save(email);
                }

                Console.WriteLine("Exiting");
            }
            catch (Exception e)
            {
                Console.WriteLine("An error has occured. \n:" + e.Message);
            }
        }
        static List<string> DownloadAttachments(ExchangeService service, ItemId itemId, string saveFolder)
        {
            EmailMessage message = EmailMessage.Bind(service, itemId, new PropertySet(ItemSchema.Attachments));
            List<string> paths = new List<string>();
            foreach (Attachment attachment in message.Attachments)
            {
                if (attachment is FileAttachment)
                {
                    FileAttachment fileAttachment = attachment as FileAttachment;
                    // Load the attachment into a file.
                    // This call results in a GetAttachment call to EWS.
                    try
                    {
                        string path = Path.Combine(saveFolder, itemId.UniqueId.ToString(), fileAttachment.Name);
                        fileAttachment.Load(path);
                        Console.WriteLine("File attachment saved at : " + path);
                        paths.Add(path);
                    }
                    catch (Exception ex)
                    {
                        Console.WriteLine("Unable to save attachment. Exception : "+ex.Message);
                    }
                }
                else if (attachment is ItemAttachment)
                {
                    ItemAttachment itemAttachment = attachment as ItemAttachment;
                    itemAttachment.Load(ItemSchema.MimeContent);
                    string path = Path.Combine(saveFolder, itemId.UniqueId.ToString(), itemAttachment.Item.Subject + ".eml");

                    // Write the bytes of the attachment into a file.
                    File.WriteAllBytes(path, itemAttachment.Item.MimeContent.Content);
                    Console.WriteLine("Email attachment saved at : " + path);
                    paths.Add(path);
                }
            }
            return paths;
        }
    }
    class Write2DB
    {
        SqlConnection conn = null;

        public Write2DB()
        {
            Console.WriteLine("Connecting to SQL Server");
            try
            {
                conn = new SqlConnection("Server=<SQLServer>;DataBase=<database>;uid=<uid>;pwd=<pwd>;Connection Timeout=1");
               // conn.Open();
                Console.WriteLine("Connected");
            }
            catch (System.Data.SqlClient.SqlException e)
            {
                throw (e);
            }
        }

        public void Save(EmailMessage email,List<string> files)
        {
            email.Load(new PropertySet(BasePropertySet.FirstClassProperties, ItemSchema.TextBody));

            SqlCommand cmd = new SqlCommand("dbo.usp_servicedesk_savemail", conn)
            {
                CommandType = System.Data.CommandType.StoredProcedure,
                CommandTimeout = 1500
            };

            string recipients = "";

            foreach (EmailAddress emailAddress in email.CcRecipients)
            {
                if (recipients != "")
                    recipients += ";";

                recipients += emailAddress.Address.ToString();
            }
            cmd.Parameters.AddWithValue("@message_id", email.InternetMessageId);
            cmd.Parameters.AddWithValue("@from", email.From.Address);
            cmd.Parameters.AddWithValue("@body", email.Body.ToString());
            cmd.Parameters.AddWithValue("@cc", recipients);
            cmd.Parameters.AddWithValue("@subject", email.Subject);
            cmd.Parameters.AddWithValue("@received_time", email.DateTimeReceived.ToUniversalTime().ToString());

            recipients = "";
            foreach (EmailAddress emailAddress in email.ToRecipients)
            {
                if (recipients != "")
                    recipients += ";";

                recipients += emailAddress.Address.ToString();
            }
            cmd.Parameters.AddWithValue("@to", recipients);

            // Execute the procedure
            cmd.ExecuteNonQuery();
        }
        ~Write2DB()
        {
            Console.WriteLine("Disconnecting from SQLServer");
        }
    }
}
