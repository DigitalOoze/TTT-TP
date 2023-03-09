using System;
using System.Collections.Generic;
using System.Linq;
using System.Net.Mail;
using System.Net;
using System.Text;
using System.Threading.Tasks;
using System.Net.Http;
using MailKit.Net.Imap;
using Microsoft.Extensions.Configuration;

namespace YellowPagesScrapper
{
    public static class MailSender
    {
        public static void SendEmail(string toEmail, string subject, string body)
        {

            var builder = new ConfigurationBuilder()
                .SetBasePath(AppDomain.CurrentDomain.BaseDirectory)
                .AddJsonFile("config.json", optional: true, reloadOnChange: true);

            var configuration = builder.Build();

            // Set the SMTP server details
            string smtpServer = "ragys.serveriai.lt";
            int smtpPort = 587; // or 465 for SSL/TLS
            string smtpUsername = configuration["User"];
            string smtpPassword = configuration["Password"];

            // Set the email details
            string fromEmail = "Papuna@digitalooze.xyz";

            // Create a new SmtpClient instance
            SmtpClient smtpClient = new SmtpClient(smtpServer, smtpPort);
            smtpClient.EnableSsl = true;
            smtpClient.Timeout = 30000;
            smtpClient.DeliveryMethod = SmtpDeliveryMethod.Network;
            smtpClient.UseDefaultCredentials = false;
            smtpClient.Credentials = new NetworkCredential(smtpUsername, smtpPassword);

            // Create a new MailMessage instance
            MailMessage mailMessage = new MailMessage();
            mailMessage.From = new MailAddress(fromEmail);
            mailMessage.To.Add(new MailAddress(toEmail));
            mailMessage.Subject = subject;
            mailMessage.Body = body;

            try
            {
                smtpClient.Send(mailMessage);
                Console.WriteLine("Email sent successfully!");
            }
            catch (Exception ex)
            {
                Console.WriteLine("Failed to send email: " + ex.Message);
            }

            // Dispose of the SmtpClient and MailMessage objects
            smtpClient.Dispose();
            mailMessage.Dispose();
        }
    }
}