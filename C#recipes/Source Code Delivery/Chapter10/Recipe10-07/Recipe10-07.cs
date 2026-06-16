using System;
using System.Net;
using System.Net.Mail;

namespace Apress.VisualCSharpRecipes.Chapter10
{
    class Recipe10_07
    {
        public static void Main(string[] args)
        {
            // Create and configure the SmtpClient that will send the mail.
            // Specify the host name of the SMTP server and the port used 
            // to send mail.
            SmtpClient client = new SmtpClient("mail.somecompany.com", 25);

            // Configure the SmtpClient with the credentials used to connect
            // to the SMTP server.
            client.Credentials =
                new NetworkCredential("user@somecompany.com", "password");

            // Create the MailMessage to represent the e-mail being sent.
            using (MailMessage msg = new MailMessage())
            {
                // Configure the e-mail sender and subject.
                msg.From = new MailAddress("author@visual-csharp-recipes.com");
                msg.Subject = "Greetings from Visual C# Recipes";

                // Configure the e-mail body.
                msg.Body = "This is a message from Recipe 10-07 of" +
                    " Visual C# Recipes. Attached is the source file " +
                    " and the binary for the recipe.";

                // Attach the files to the e-mail message and set their MIME type.
                msg.Attachments.Add(
                    new Attachment(@"..\..\Recipe10-07.cs", "text/plain"));
                msg.Attachments.Add(
                    new Attachment(@".\Recipe10-07.exe", 
                    "application/octet-stream"));

                // Iterate through the set of recipients specified on the 
                // command line. Add all addresses with the correct structure as 
                // recipients.
                foreach (string str in args)
                {
                    // Create a MailAddress from each value on the command line 
                    // and add it to the set of recipients.
                    try
                    {
                        msg.To.Add(new MailAddress(str));
                    }
                    catch (FormatException ex)
                    {
                        // Proceed to the next specified recipient.
                        Console.WriteLine("{0}: Error -- {1}", str, ex.Message);
                        continue;
                    }
                }

                // Send the message.
                client.Send(msg);
            }

            // Wait to continue.
            Console.WriteLine(Environment.NewLine);
            Console.WriteLine("Main method complete. Press Enter");
            Console.ReadLine();
        }
    }
}
