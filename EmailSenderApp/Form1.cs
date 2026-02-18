using MimeKit;
using System;
using System.IO;
using System.Linq;
using System.Net.Mail;
using System.Net.Mime;
using MailKit.Net.Smtp;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace EmailSenderApp
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            string csvPath = @"C:\Users\user\Documents\emails.csv";
            string cvPath = @"C:\Users\user\Documents\xxxxx_xxxxx_CV.docx";

            // Gmail login - use App Password if 2FA enabled
            string gmailUser = "xxxxxxl@gmail.com";
            string gmailPass = "xxxxxxxxxxxxxx";

            // Read all emails from CSV (skip header)
            var emails = File.ReadAllLines(csvPath);
            for (int i = 0; i < emails.Length; i++)
            {
                string recipient = emails[i].Trim();
                if (string.IsNullOrEmpty(recipient)) continue;

                var message = new MimeMessage();
                message.From.Add(new MailboxAddress("xxxxxxxxxxxx", gmailUser));
                message.To.Add(new MailboxAddress("", recipient));
                message.Subject = "Application for [xxxxxxxxxx] Position";

                message.Body = new TextPart("plain")
                {
                    Text = @"Hello,

I hope you are doing well.

My name is xxxxxx xxxxxxxx, and I am writing to express my interest in a [xxxxxxxxx] position within your organization. I have solid experience in C# Windows Forms, Oracle, and SQL Server, and I would welcome the opportunity to contribute to your team.

Please find my CV attached for your review. I would be happy to provide any additional information if needed.

Thank you for your time and consideration. I look forward to the possibility of discussing how I can contribute to your company.

Best regards,
xxxxx xxxxxxx
[number phone]
"
                };

                // Attach CV
                var attachment = new MimePart("application", "vnd.openxmlformats-officedocument.wordprocessingml.document")
                {
                    Content = new MimeContent(File.OpenRead(cvPath)),
                    ContentTransferEncoding = ContentEncoding.Base64,

                    FileName = "xxxx_xxxxx_XXXX.docx"
                };

                var multipart = new Multipart("mixed");
                multipart.Add(message.Body);
                multipart.Add(attachment);
                message.Body = multipart;

                // Send email via Gmail SMTP
                using (var client = new MailKit.Net.Smtp.SmtpClient())
                {
                    client.Connect("smtp.gmail.com", 587, MailKit.Security.SecureSocketOptions.StartTls);

                    client.Authenticate(gmailUser, gmailPass);
                    client.Send(message);
                    client.Disconnect(true);
                }

                Console.WriteLine($"Email sent to: {recipient}");
            }

            Console.WriteLine("All emails sent!");
        }
    }
}

