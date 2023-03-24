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
    }
}
