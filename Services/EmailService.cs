using MailKit.Net.Smtp;
using MailKit.Security;
using MimeKit;
using MimeKit.Text;
using SimpleSendMailAPI.Dto;
using ClosedXML.Excel;

namespace SimpleSendMailAPI.Services
{
    public interface IEmailService
    {
        void SendEmail(EmailDto request);
        void SendEmailToAllEmailFromExcel(string subject);
    }

    public class EmailService : IEmailService
    {
        private readonly IConfiguration _config;
        public EmailService(IConfiguration config)
        {
            _config = config;
        }

        public void SendEmail(EmailDto request)
        {
            MimeMessage email = new MimeMessage();
            string htmlMessage = "";
            string? dtoHtmlMessage = "";
            if (request.Body == "string" || request.Body == "")
            {
                htmlMessage = File.ReadAllText(@"Documents\message.html");
            }
            else
            {
                dtoHtmlMessage = request.Body;
            }
            email.From.Add(MailboxAddress.Parse(_config.GetSection("EmailUsername").Value));
            email.To.Add(MailboxAddress.Parse(request.To));
            email.Subject = request.Subject;
            email.Body = new TextPart(TextFormat.Html) { Text = htmlMessage + dtoHtmlMessage };

            SendMailBySmtp(email);
        }

        public void SendMailBySmtp(MimeMessage email)
        {
            using var smtp = new SmtpClient();
            smtp.Connect(_config.GetSection("EmailHost").Value, 587, SecureSocketOptions.StartTls);
            smtp.Authenticate(_config.GetSection("EmailUsername").Value, _config.GetSection("EmailPassword").Value);
            smtp.Send(email);
            smtp.Disconnect(true);
        }

        public void SendEmailToAllEmailFromExcel(string subject)
        {
            List<string> mailAddressList = ReadMailAddressFromExcelFile();
            string htmlMessage = File.ReadAllText(@"Documents\message.html");
            foreach (string mailAdress in mailAddressList)
            {
                MimeMessage email = new MimeMessage();
                email.From.Add(MailboxAddress.Parse(_config.GetSection("EmailUsername").Value));
                email.To.Add(MailboxAddress.Parse(mailAdress));
                email.Subject = subject;
                email.Body = new TextPart(TextFormat.Html) { Text = htmlMessage };

                SendMailBySmtp(email);
            }

        }

        private List<string> ReadMailAddressFromExcelFile()
        {
            List<string> mailAddressList = new List<string>();
            string filePath = @"Documents\ListEmail.xlsx";
            using var workbook = new XLWorkbook(filePath);
            var ws = workbook.Worksheet(1);
            for (int i = 1; i <= 100; i++)
            {
                var data = ws.Cell($"A{i}").GetValue<string>();
                if (data != "")
                {
                    mailAddressList.Add(data.ToString());
                }
            }
            return mailAddressList;
        }
    }
}