
using System;
using System.Net;
using System.Net.Mail;
using System.Text;
using System.Threading;

namespace EmailExample
{
    class Program
    {
        static void Main(string[] args)
        {
            Console.ForegroundColor = ConsoleColor.Red;
            Console.WriteLine("Email sender");
            Console.WriteLine("Use carefully");
            Console.ResetColor();

            Console.WriteLine("Type your email - It must be an outlook.com");
            string email = Console.ReadLine();
            Console.WriteLine("Type your email password");
            string password = Console.ReadLine();

            // Reset console color
            Console.ResetColor();

            Console.WriteLine("Type the destination email");
            string dest_email = Console.ReadLine();

            Console.WriteLine("Type the email title or leave empty for a default title");
            string eTitle = Console.ReadLine();

            // If the email title is empty, set the title to "Email Title"
            if (eTitle.Length == 0)
            {
                eTitle = "Email Title";
            }

            Console.WriteLine("Type the email body or leave empty for a default body");
            string eBody = Console.ReadLine();

            // If the email body is empty, set the body to "Email Body"
            if (eBody.Length == 0)
            {
                eBody = "Email Body";
            }

            // Asks the user to input the number of times they want to send the email
            Console.WriteLine("Now type how many times you want to send the email");
            int times = int.Parse(Console.ReadLine());

            // Loop that sends the email the specified number of times
            for (int i = 0; i < times; i++)
            {
                //Makes the program sleeps for 700 milisecons
                //NOTE: if you lower this value, your email may not be sent 
                //because it will be target as spam
                Thread.Sleep(700);

                //Create a thread to send emails faster. 
                //If you do not want to send faster, you can increase the value of 'Thread.Sleep()'
                //or remove the threads.
                Thread thread = new Thread(() =>
                {
                    try
                    {
                        Console.WriteLine($"email {i} sent successfully!");
                        // Create a new MailMessage object
                        MailMessage mail = new MailMessage();
                        // Create a new SmtpClient object using the Outlook.com SMTP server
                        SmtpClient SmtpServer = new SmtpClient("smtp-mail.outlook.com");

                        // Set the email address of the sender
                        mail.From = new MailAddress(email);
                        // Add the destination email address
                        mail.To.Add(dest_email);

                        //Use the line below to send a image

                        //mail.Attachments.Add(new Attachment("your image path"));

                        // Set the subject of the email
                        mail.Subject = eTitle;
                        // Set the body of the email
                        mail.Body = eBody;

                        // Set the port to use for the SMTP connection
                        SmtpServer.Port = 587;
                        // Set the email and password credentials to use for authentication
                        SmtpServer.Credentials = new NetworkCredential(email, password);
                        // Enable SSL encryption for the SMTP connection
                        SmtpServer.EnableSsl = true;

                        // Send the email
                        SmtpServer.Send(mail);
                        Console.WriteLine("email sent successfully");
                    }
                    //Catch the errors
                    catch (Exception ex)
                    {
                        //If the email could not be found
                        if (ex is SmtpFailedRecipientException)
                        {
                            Console.WriteLine("the email could not be found");
                        }
                        //if any other error occurs
                        else
                        {
                            Console.WriteLine("an error occurred: " + ex.Message);
                        }
                    }
                });
                thread.Start();
            }

            Console.WriteLine("Press any key to exit");
            Console.ReadKey();
        }
    }
}
