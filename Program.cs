using System;
using System.IO;
using System.Net.Mail;
using System.Windows.Forms;
using Microsoft.CognitiveServices.Speech;
using Microsoft.CognitiveServices.Speech.Audio;
using Microsoft.Extensions.Configuration;
using MailKit.Net.Imap;
using MailKit.Search;
using MailKit;
using MailKit.Net.Smtp;
using MimeKit;
using System.Linq;
using System.Net.Mime;

namespace Scribe
{
    

    public partial class MainForm : Form
    {
        private readonly IConfiguration _configuration;

        public MainForm()
        {
            InitializeComponent();

            // Load the app settings from appsettings.json
            var builder = new ConfigurationBuilder()
                .SetBasePath(Directory.GetCurrentDirectory())
                .AddJsonFile("appsettings.json", optional: true, reloadOnChange: true);

            _configuration = builder.Build();
        }

        class Program
        {
            [STAThread]
            static void Main()
            {
                Application.EnableVisualStyles();
                Application.SetCompatibleTextRenderingDefault(false);
                Application.Run(new MainForm());
            }
        }

        private void ProcessEmailsButton_Click(object sender, EventArgs e)
        {
            // Retrieve new emails from the email server.
            var emails = RetrieveNewEmails();

            // Process each email.
            foreach (var email in emails)
            {
                ProcessEmail(email);
            }

            MessageBox.Show("Email processing complete.");
        }

        private async void ProcessVoicemail(string voicemailPath)
        {
            // Use the Speech SDK to transcribe the voicemail.
            var config = SpeechConfig.FromSubscription(_configuration["SpeechToText:SubscriptionKey"], _configuration["SpeechToText:Region"]);
            using var audioInput = AudioConfig.FromWavFileInput(voicemailPath);
            using var recognizer = new SpeechRecognizer(config, audioInput);
            var result = await recognizer.RecognizeOnceAsync();

            if (result.Reason == ResultReason.RecognizedSpeech)
            {
                Console.WriteLine($"Transcription: {result.Text}");
                // Do something with the transcription, such as send an email with the transcription as the body.
            }
            else
            {
                Console.WriteLine($"Speech recognition failed: {result.Reason}");
            }
        }

        private void ProcessVoicemailButton_Click(object sender, EventArgs e)
        {
            // Get the paths of new voicemail files.
            var voicemailPaths = GetNewVoicemailPaths();

            // Process each voicemail.
            foreach (var voicemailPath in voicemailPaths)
            {
                ProcessVoicemail(voicemailPath);
            }

            MessageBox.Show("Voicemail processing complete.");
        }

        private MailMessage MimeMessageToMailMessage(MimeMessage mimeMessage)
        {
            var mailMessage = new MailMessage();
            mailMessage.From = new MailAddress(mimeMessage.From[0].ToString());

            if (mimeMessage.To.Count > 0)
            {
                mailMessage.To.Add(new MailAddress(mimeMessage.To[0].ToString()));
            }

            mailMessage.Subject = mimeMessage.Subject;

            if (mimeMessage.Attachments != null)
            {
                foreach (var attachment in mimeMessage.Attachments)
                {
                    var mimePart = attachment as MimePart;
                    if (mimePart != null)
                    {
                        mailMessage.Attachments.Add(new Attachment(mimePart.Content.Stream, mimePart.FileName));
                    }
                }
            }

            mailMessage.Body = mimeMessage.TextBody;

            return mailMessage;
        }

        private MailMessage[] RetrieveNewEmails()
        {
            using var client = new ImapClient();
            client.Connect(_configuration["Email:Server"], Convert.ToInt32(_configuration["Email:995"]), Convert.ToBoolean(_configuration["Email:UseSsl"]));
            client.Authenticate(_configuration["Email:nick_sumner@yahoo.com"], _configuration["Email:ygdjyupmgkcnlsoz"]);
            var inbox = client.Inbox;
            inbox.Open(FolderAccess.ReadOnly);
            var searchQuery = SearchQuery.NotSeen;
            var uids = inbox.Search(searchQuery);
            var messages = new MailMessage[uids.Count];
            for (int i = 0; i < uids.Count; i++)
            {
                var message = inbox.GetMessage(uids[i]);
                messages[i] = MimeMessageToMailMessage(message);
            }
            client.Disconnect(true);
            return messages;
        }


        private string[] GetNewVoicemailPaths()
        {
            // Get the paths of new voicemail files from the appropriate directory.
            // For example:
            var voicemailDirectory = _configuration["Voicemail:Directory"];
            return Directory.GetFiles(voicemailDirectory, "*.wav");
        }

        private void InitializeComponent()
        {
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(MainForm));
            button1 = new Button();
            button2 = new Button();
            button3 = new Button();
            button4 = new Button();
            button5 = new Button();
            button6 = new Button();
            SuspendLayout();
            // 
            // button1
            // 
            button1.BackColor = Color.DarkGoldenrod;
            button1.FlatStyle = FlatStyle.Popup;
            button1.ForeColor = Color.White;
            button1.Location = new Point(12, 262);
            button1.Name = "button1";
            button1.Size = new Size(75, 23);
            button1.TabIndex = 0;
            button1.Text = "Prev";
            button1.UseVisualStyleBackColor = false;
            // 
            // button2
            // 
            button2.BackColor = Color.DarkGoldenrod;
            button2.FlatStyle = FlatStyle.Popup;
            button2.ForeColor = Color.White;
            button2.Location = new Point(266, 262);
            button2.Name = "button2";
            button2.Size = new Size(75, 23);
            button2.TabIndex = 1;
            button2.Text = "Play";
            button2.UseVisualStyleBackColor = false;
            button2.Click += button2_Click;
            // 
            // button3
            // 
            button3.BackColor = Color.DarkGoldenrod;
            button3.FlatStyle = FlatStyle.Popup;
            button3.ForeColor = Color.White;
            button3.Location = new Point(473, 262);
            button3.Name = "button3";
            button3.Size = new Size(75, 23);
            button3.TabIndex = 2;
            button3.Text = "Next";
            button3.UseVisualStyleBackColor = false;
            // 
            // button4
            // 
            button4.BackColor = Color.DarkGoldenrod;
            button4.FlatStyle = FlatStyle.Popup;
            button4.ForeColor = Color.White;
            button4.Location = new Point(12, 12);
            button4.Name = "button4";
            button4.Size = new Size(75, 23);
            button4.TabIndex = 3;
            button4.Text = "button4";
            button4.UseVisualStyleBackColor = false;
            // 
            // button5
            // 
            button5.BackColor = Color.DarkGoldenrod;
            button5.FlatStyle = FlatStyle.Popup;
            button5.ForeColor = Color.White;
            button5.Location = new Point(266, 12);
            button5.Name = "button5";
            button5.Size = new Size(75, 23);
            button5.TabIndex = 4;
            button5.Text = "Stop";
            button5.UseVisualStyleBackColor = false;
            // 
            // button6
            // 
            button6.BackColor = Color.DarkGoldenrod;
            button6.FlatStyle = FlatStyle.Popup;
            button6.ForeColor = Color.White;
            button6.Location = new Point(473, 12);
            button6.Name = "button6";
            button6.Size = new Size(75, 23);
            button6.TabIndex = 5;
            button6.Text = "button6";
            button6.UseVisualStyleBackColor = false;
            // 
            // MainForm
            // 
            BackgroundImage = (Image)resources.GetObject("$this.BackgroundImage");
            ClientSize = new Size(560, 297);
            Controls.Add(button6);
            Controls.Add(button5);
            Controls.Add(button4);
            Controls.Add(button3);
            Controls.Add(button2);
            Controls.Add(button1);
            Name = "MainForm";
            ResumeLayout(false);
        }

        private void ProcessEmail(MailMessage email)
        {
            // Extract the voicemail attachment from the email (if present).
            Attachment voicemailAttachment = email.Attachments.FirstOrDefault(a =>
                a.ContentType?.MediaType == "audio/wav" &&
                a.ContentDisposition?.DispositionType == "attachment" &&
                !string.IsNullOrEmpty(a.ContentDisposition.FileName));

            if (voicemailAttachment == null)
            {
                Console.WriteLine("No voicemail attachment found in email.");
                return;
            }

            // Save the attachment to a file.
            var voicemailDirectory = _configuration["Voicemail:Directory"];
            var fileName = voicemailAttachment.ContentDisposition.FileName;
            var filePath = Path.Combine(voicemailDirectory, fileName);
            using (var fileStream = new FileStream(filePath, FileMode.Create))
            {
                voicemailAttachment.ContentStream.CopyTo(fileStream);
            }

            // Process the voicemail.
            ProcessVoicemail(filePath);
        }

        private void button2_Click(object sender, EventArgs e)
        {

        }

        private Button button1;
        private Button button2;
        private Button button3;
        private Button button4;
        private Button button5;
        private Button button6;
    }
}
