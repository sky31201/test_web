using MailKit.Net.Smtp;
using MimeKit;
using NLog;
using System;
using System.Collections.Generic;
using System.Configuration;
using System.IO;
using System.Linq;
using System.Security.Cryptography.X509Certificates;
using System.Text;
using System.Threading.Tasks;
using System.Web.Hosting;

namespace MohwEmail.Resender
{
    class Program
    {
        static string _smtpHost = ConfigurationManager.AppSettings["SmtpHost"];
        static int _smtpPort = Int32.Parse(ConfigurationManager.AppSettings["SmtpPort"]);
        static string _mohwAcct = ConfigurationManager.AppSettings["MOHWAcct"];
        static string _mohwPwd = ConfigurationManager.AppSettings["MOHWEmailSecret"];

        public static Logger logger = LogManager.GetCurrentClassLogger();
        readonly public static string failureOutboxPath = "D:\\Projects\\mohwemail\\FileStorage\\FailureOutbox";


        static void Main(string[] args)
        {

            if (GetUnreadEmailCount(failureOutboxPath)== 0)
            {
                return;
            }

            System.Net.ServicePointManager.ServerCertificateValidationCallback =
                    new System.Net.Security.RemoteCertificateValidationCallback(ValidateServerCertificate);

            using (SmtpClient client = new SmtpClient())
            {
                client.CheckCertificateRevocation = false;
                client.ServerCertificateValidationCallback = (s, c, h, e) => true;

                logger.Info($"Establishing connections to {_smtpHost}:{_smtpPort} ");
                Console.WriteLine($"Establishing connections to {_smtpHost}:{_smtpPort} ");
                // 建立連線
                client.Connect(_smtpHost, _smtpPort, MailKit.Security.SecureSocketOptions.Auto);

                logger.Info($"Logining in with {_mohwAcct}...");
                Console.WriteLine($"Logining in with {_mohwAcct}...");
                client.Authenticate(_mohwAcct, _mohwPwd);

                Console.WriteLine("Sending mails..");
                foreach (string file in Directory.GetFiles(failureOutboxPath, "*.eml"))
                {
                    try
                    {
                        MimeMessage mimeMessage = MimeMessage.Load(file);
                        client.SendAsync(mimeMessage);
                        File.Delete(file);
                    }
                    catch (Exception)
                    {
                        logger.Warn($"Failed to send {file}");
                    }

                }

                // 中斷連線
                client.Disconnect(true);
            }

            logger.Info("SendMail End.");
            Console.ReadKey();
        }

        private static int GetUnreadEmailCount(string path)
        {
            return Directory.GetFiles(path, "*.eml", SearchOption.TopDirectoryOnly).Length;
        }

        public static bool ValidateServerCertificate(Object sender, X509Certificate certificate, X509Chain chain, System.Net.Security.SslPolicyErrors sslPolicyErrors)
        {
            return true;
        }
    }
}
