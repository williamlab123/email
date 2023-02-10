using System;
using System.Net;
using System.Net.Mail;
using System.Text;

namespace EmailExample
{
    class Program
    {
        static void Main(string[] args)
        {

            Console.ForegroundColor = ConsoleColor.Red;
            System.Console.WriteLine("Email sender");
            System.Console.WriteLine("Use carefully");
            Console.ResetColor();

            System.Console.WriteLine("Type your email - It must be an outlook.com");
            string email = Console.ReadLine();
            System.Console.WriteLine("Type your email password");
            string password = Console.ReadLine();
          
   

           
            // Reset da cor do console
            Console.ResetColor();

            System.Console.WriteLine("Type the destination email");
            string dest_email = Console.ReadLine();


            System.Console.WriteLine("Type the email title or leave empty for a default title");
            string eTitle = Console.ReadLine();

            if (eTitle.Length == 0)
            {
                eTitle = "Email Title";
            }

            System.Console.WriteLine("Type the email body or leave empty for a default body");
            string eBody = Console.ReadLine();

            if (eBody.Length == 0)
            {
                eBody = "Email Body";
            }



            System.Console.WriteLine("Now type how many times you want to send the email");
            int times = int.Parse(Console.ReadLine());

            for (int i = 0; i < times; i++)
            {
                try
                {
                    MailMessage mail = new MailMessage();
                    SmtpClient SmtpServer = new SmtpClient("smtp-mail.outlook.com");
              

                    mail.From = new MailAddress(email);
                    mail.To.Add(dest_email);
                    mail.Subject = eTitle;
                    mail.Body = eBody;

                    SmtpServer.Port = 587;
                    SmtpServer.Credentials = new System.Net.NetworkCredential(email, password);
                    SmtpServer.EnableSsl = true;

                    SmtpServer.Send(mail);
                    Console.WriteLine("email sent successfully");
                }
                catch (Exception ex)
                {
                    if (ex is SmtpFailedRecipientException)
                    {
                        Console.WriteLine("the email could not be found");
                           System.Environment.Exit(0);
                    }
                    else
                    {
                        Console.WriteLine("an error occurred: " + ex.Message);
                        System.Environment.Exit(0);
                    }
                }

            }
               System.Environment.Exit(0);

        }
        
    }
}