using MailKit;
using MailKit.Net.Imap;
using MailKit.Search;
using MimeKit;
using Newtonsoft.Json;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Threading;

namespace CheckingMails
{
    internal class Program1
    {
        static void Main(string[] args)
        {
            try
            {
                Program1 program = new Program1();
                program.Connect();

                while (true)
                {
                    program.CheckEmails();
                    Thread.Sleep(TimeSpan.FromMinutes(5));
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine("ERROR: " + ex.Message);
            }
        }

        public string MailAddress { get; set; } = "example yandex mail";
        public string Password { get; set; } = "the password";
        public int Port { get; set; } = 993;
        public string Host { get; set; } = "imap.yandex.com";
        public bool Ssl { get; set; } = true;
        ImapClient client = new ImapClient();
        public string FilePath { get; set; } = @"C:\Users\berka\Desktop\eMailFiles\eMails.txt";
        public string AttachmentPath { get; set; } = @"C:\Users\berka\Desktop\eMailFiles\attachments";

        public class Attachment
        {
            public string FileName { get; set; }
        }

        public class Message
        {
            public ulong Id { get; set; }
            public string Sender { get; set; }
            public DateTime Time { get; set; }
            public string Body { get; set; }
            public List<Attachment> Attachments { get; set; } = new List<Attachment>();
        }

        public void Connect()
        {
            try
            {
                client.Connect(Host, Port, Ssl);
                client.Authenticate(MailAddress, Password);
                Console.WriteLine("Connected successfully.");
            }
            catch (Exception ex)
            {
                Console.WriteLine("Connection failed. " + ex.Message);
            }
        }

        public void CheckEmails()
        {
            try
            {
                client.Inbox.Open(FolderAccess.ReadOnly);
                var uids = client.Inbox.Search(SearchQuery.NotSeen);

                List<Message> messages = new List<Message>();
                HashSet<ulong> seenIds = new HashSet<ulong>();

                if (File.Exists(FilePath))
                {
                    string fileContent = File.ReadAllText(FilePath);
                    messages = JsonConvert.DeserializeObject<List<Message>>(fileContent);
                    seenIds = new HashSet<ulong>(messages.Select(m => m.Id));
                }

                List<Message> newMessages = new List<Message>();

                foreach (var uid in uids)
                {
                    var message = client.Inbox.GetMessage(uid);
                    var message1 = new Message
                    {
                        Id = uid.Id,
                        Sender = message.From.ToString(),
                        Time = message.Date.DateTime,
                        Body = message.HtmlBody,
                    };

                    foreach (var attachment in message.Attachments)
                    {
                        if (attachment is MimePart part) // Multipurpose Internet Mail Extensions
                        {
                            var fileName = part.FileName;
                            var filePath = Path.Combine(AttachmentPath, fileName);

                            // Ek içeriğini dosya olarak kaydet
                            Directory.CreateDirectory(AttachmentPath);
                            using (var fileStream = File.Create(filePath))
                            {
                                part.Content.DecodeTo(fileStream);
                            }

                            message1.Attachments.Add(new Attachment
                            {
                                FileName = fileName,
                            });
                        }
                    }

                    if (!seenIds.Contains(message1.Id))
                    {
                        newMessages.Add(message1);
                        seenIds.Add(message1.Id);
                        messages.Add(message1);
                    }
                }

                if (newMessages.Count > 0)
                {
                    string newJson = JsonConvert.SerializeObject(messages, Formatting.Indented);
                    using (StreamWriter sw = new StreamWriter(FilePath, false))
                    {
                        sw.WriteLine(newJson);
                    }
                    Console.WriteLine("{0} new messages added", newMessages.Count);
                }
                else
                {
                    Console.WriteLine("No new messages.");
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine("Error: " + ex.Message);
                Console.WriteLine(ex.StackTrace);
            }
        }
    }
}
