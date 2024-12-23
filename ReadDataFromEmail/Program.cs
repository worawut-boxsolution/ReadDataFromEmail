// Setting Email with Azure https://www.codewrecks.com/post/security/accessing-office-365-imap-with-oauth2/

using System;
using MailKit.Net.Imap;
using MailKit;
using MimeKit;
using MailKit.Search;
using MailKit.Security;
class Program
{
    static void Main(string[] args)
    {
        // Define your email server settings
        //string email = "Worawut.cha@outlook.com";
        //string password = "Worawut02";
        //string imapServer = "outlook.office365.com";

        string email = "worawutlchaingam@gmail.com";
        string password = "W@ssw0rd0123";
        string imapServer = "imap.gmail.com";
        int port = 993;

        // Create an IMAP client
        using (var client = new ImapClient())
        {
            try
            {
                // Connect to the IMAP server
                client.Connect(imapServer, port, true);
                //client.AuthenticationMechanisms.Remove("XOAUTH2");
                //client.AuthenticationMechanisms.Remove("NTLM");

                // Authenticate with the server
                client.Authenticate(email, password);

                // Select the inbox folder
                client.Inbox.Open(FolderAccess.ReadOnly);

                // Fetch the first 10 messages
                for (int i = 0; i < Math.Min(10, client.Inbox.Count); i++)
                {
                    var message = client.Inbox.GetMessage(i);
                    Console.WriteLine($"Subject: {message.Subject}");
                    Console.WriteLine($"From: {message.From}");
                    Console.WriteLine($"Date: {message.Date}");
                    Console.WriteLine($"Body: {message.TextBody}");
                    Console.WriteLine(new string('-', 50));
                }

                // Disconnect from the server
                client.Disconnect(true);
            }
            catch (AuthenticationException ex) { Console.WriteLine($"Authentication failed: {ex.Message}"); }
            catch (ImapCommandException ex) { Console.WriteLine($"IMAP command failed: {ex.Message}"); }
            catch (Exception ex) { Console.WriteLine($"An error occurred: {ex.Message}"); }
            //catch (Exception ex)
            //{
            //    Console.WriteLine($"An error occurred: {ex.Message}");
            //}
        }
    }
}


//public class Program
//{
//    public static async Task Main(string[] args)
//    {
//        string email = "worawut.cha@outlook.com";
//        string password = "Worawut01";
//        string host = "outlook.live.com";
//        int port = 993; // Use port 993 for IMAP over SSL
//        string savePath = "./";

//        var emailHandling = new EmailHandling(email, password);
//        await emailHandling.DownloadAttachmentAsync(host, port, savePath);
//    }
//}

//public class EmailHandling
//{
//    private readonly string email;
//    private readonly string password;

//    public EmailHandling(string email, string password)
//    {
//        this.email = email;
//        this.password = password;
//    }

//    public async Task DownloadAttachmentAsync(string host, int port, string savePath)
//    {
//        try
//        {
//            using (var client = new ImapClient())
//            {
//                using (var cancel = new CancellationTokenSource())
//                {
//                    await client.ConnectAsync(host, port, SecureSocketOptions.SslOnConnect, cancel.Token);

//                    // Remove unnecessary authentication mechanisms
//                    client.AuthenticationMechanisms.Remove("XOAUTH2");

//                    await client.AuthenticateAsync(email, password, cancel.Token);

//                    var inbox = client.Inbox;
//                    await inbox.OpenAsync(FolderAccess.ReadOnly, cancel.Token);

//                    var uniqueIds = await inbox.SearchAsync(SearchQuery.NotSeen, cancel.Token);

//                    foreach (var uniqueId in uniqueIds.OrderByDescending(id => id.Id).Take(3))
//                    {
//                        var message = await inbox.GetMessageAsync(uniqueId, cancel.Token);

//                        // Check if the email is from a specific address
//                        if (message.From.Mailboxes.Any(m => m.Address.Equals("noreply@dhl.com", StringComparison.OrdinalIgnoreCase)))
//                        {
//                            foreach (var attachment in message.Attachments)
//                            {
//                                if (attachment is MimePart part && part.FileName.EndsWith(".zip", StringComparison.OrdinalIgnoreCase))
//                                {
//                                    var fileName = Path.Combine(savePath, part.FileName);

//                                    using (var stream = File.Create(fileName))
//                                    {
//                                        await part.Content.DecodeToAsync(stream, cancel.Token);
//                                    }

//                                    Console.WriteLine($"Attachment saved: {fileName}");
//                                }
//                            }
//                        }
//                    }

//                    await client.DisconnectAsync(true, cancel.Token);
//                }
//            }
//        }
//        catch (AuthenticationException ex)
//        {
//            Console.WriteLine("Authentication failed.");
//            Console.WriteLine(ex.Message);
//        }
//        catch (Exception ex)
//        {
//            Console.WriteLine("An error occurred while downloading the attachment.");
//            Console.WriteLine(ex.Message);
//        }
//    }
//}