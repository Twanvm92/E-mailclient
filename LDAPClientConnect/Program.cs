using System;
using System.Collections;
using System.Collections.Generic;
using System.DirectoryServices;
using System.Linq;
using MailKit;
using MailKit.Net.Imap;
using MailKit.Net.Smtp;
using MimeKit;


namespace LDAPClientConnect
{
	class Program
	{
		static void Main(string[] args)
		{
			
			if (args.Length != 7)
			{
				Console.WriteLine("Usage: <LDAP path> <email username> <email password> <IMAP host> <SMTP host> <IMAP port> <SMTP port>");
				Console.ReadKey();
                Environment.Exit(1);
            }
       
			var ldapPath = args[0];
			var username = args[1];
			var password = args[2];
			var imapHost = args[3];
			var smtpHost = args[4];
			var imapPort = int.Parse(args[5]);
			var smtpPort = int.Parse(args[6]);

            DirectoryEntry entry =
            new DirectoryEntry(ldapPath, "", "", AuthenticationTypes.None);

            DirectorySearcher search = new DirectorySearcher(entry);
            search.Filter = "(objectClass=person)";
            search.PropertiesToLoad.Add("mail");
            search.PropertiesToLoad.Add("cn");

            SearchResultCollection results = search.FindAll();

			Console.WriteLine("Retrieving all users in directory..");
			Console.WriteLine("");

            foreach (SearchResult searchResult in results)
            {
                foreach (DictionaryEntry property in searchResult.Properties)
                {
                    Console.Write(property.Key + ": ");
                    foreach (var val in ((ResultPropertyValueCollection)property.Value))
                    {
                        Console.Write(val + "; ");
                    }
                    Console.WriteLine("");
                }
            }

            using (var imapClient = new ImapClient())
            {
                // For demo-purposes, accept all SSL certificates
                imapClient.ServerCertificateValidationCallback = (s, c, h, e) => true;

                imapClient.Connect(imapHost, imapPort, true);

                imapClient.Authenticate(username, password);

                // The Inbox folder is always available on all IMAP servers...
                var inbox = imapClient.Inbox;
                inbox.Open(FolderAccess.ReadOnly);

                Console.WriteLine("");
                Console.WriteLine("Retrieving all messages from your emailaddress..");
                Console.WriteLine("Total messages: {0}", inbox.Count);
                Console.WriteLine("Recent messages: {0}", inbox.Recent);
                Console.WriteLine("");
                Console.WriteLine("Retrieving the 10 most recent messages..");
                Console.WriteLine("");

                Dictionary<string, string[]> emailDictionary = new Dictionary<string, string[]>();
                var index = Math.Max(inbox.Count - 10, 0);
                for (int i = index; i < inbox.Count; i++)
                {
                    var message = inbox.GetMessage(i);
                    var internetAddressList = message.From;
                    var sender = "sender unknown";

                    if (internetAddressList.Count > 0)
                    {
                        sender = internetAddressList.ToString();
                        Char charRange = '"';
                        int startIndex = sender.IndexOf(charRange);
                        int endIndex = sender.LastIndexOf(charRange);
                        int length = endIndex - startIndex + 1;
          
                        var address = internetAddressList.Mailboxes.ElementAt(0).Address;
                        emailDictionary.Add(message.MessageId, new[] {sender.Substring(startIndex, length), address});
                    }
                    Console.WriteLine("MessageId: {0}, Sender: {1}, Subject: {2}", message.MessageId, sender, message.Subject);
                }

                imapClient.Disconnect(true);

                Console.WriteLine("");
                Console.WriteLine("Choose one of the above senders to reply to by typing in the MessageId..");

                string answer;
                string[] value = {};
                while ((answer = Console.ReadLine()) != null)
                {
                    if (emailDictionary.TryGetValue(answer, out value))
                    {
                        Console.WriteLine("found an emailaddress: {0}", value[1]);
                        Console.WriteLine("");
                        break;
                    }

                    Console.WriteLine("Could not find an emailaddress with this MessageId: {0}", answer);
                    Console.WriteLine("Choose one of the above senders to send an email to by typing in the MessageId..");

                }

                var mimeMessage = new MimeMessage();
                mimeMessage.From.Add(new MailboxAddress("Twan van Maastricht", "twanvm92@gmail.com"));
                mimeMessage.To.Add(new MailboxAddress(value[0], value[1]));
				
				Console.WriteLine("Type in a subject for your email");
                mimeMessage.Subject = Console.ReadLine();

                Console.WriteLine("");
                Console.WriteLine("Type in a body for your message");
				var body = Console.ReadLine();

                mimeMessage.Body = new TextPart("plain")
                {
                    Text = body
                };

                using (var smtpClient = new SmtpClient())
                {
                    // For demo-purposes, accept all SSL certificates (in case the server supports STARTTLS)
                    smtpClient.ServerCertificateValidationCallback = (s, c, h, e) => true;

                    smtpClient.Connect(smtpHost, smtpPort, false);

                    // Note: only needed if the SMTP server requires authentication
                    smtpClient.Authenticate(username, password);

                    smtpClient.Send(mimeMessage);
                    
                    smtpClient.Disconnect(true);
                    Console.WriteLine("Message was send, use a key to quit...");
                    Console.ReadKey();
                }

            }


        }
	}
}
