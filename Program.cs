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
using static System.Windows.Forms.VisualStyles.VisualStyleElement.ListView;
using System.Media;

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


        class MyProgram
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
            if (string.IsNullOrEmpty(voicemailPath))
            {
                Console.WriteLine("Invalid voicemail path.");
                return;
            }

            if (!File.Exists(voicemailPath))
            {
                Console.WriteLine($"Voicemail file not found: {voicemailPath}");
                return;
            }

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
            client.Connect(_configuration["Email:smtp.mail.yahoo.com"], Convert.ToInt32(_configuration["Email:995"]), Convert.ToBoolean(_configuration["Email:UseSsl"]));
            client.Authenticate(_configuration["Email:nick_sumner@yahoo.com"], _configuration["Email:ygdjyupmgkcnlsoz"]);
            var inbox = client.Inbox;
            inbox.Open(FolderAccess.ReadOnly);
            var searchQuery = SearchQuery.NotSeen;
            var uids = inbox.Search(searchQuery);
            var messages = new MailMessage[uids.Count];
            var port = Convert.ToInt32(_configuration["Email:995"]);
            var useSsl = Convert.ToBoolean(_configuration["Email:UseSsl"]);
            var userName = _configuration["Email:nick_sumner@yahoo.com"];
            var password = _configuration["Email:ygdjyupmgkcnlsoz"];
            var host = _configuration["Email:smtp.mail.yahoo.com"];
            if (string.IsNullOrEmpty(host))
            {
                throw new ArgumentNullException(nameof(host), "The host parameter cannot be null or empty.");
            }
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
            button1 = new Button();
            Play = new Button();
            button3 = new Button();
            ReadScreen = new Button();
            button5 = new Button();
            ProcessVM = new Button();
            label1 = new Label();
            panel1 = new Panel();
            panel2 = new Panel();
            textBox1 = new TextBox();
            panel2.SuspendLayout();
            SuspendLayout();
            // 
            // button1
            // 
            button1.BackColor = Color.Black;
            button1.FlatStyle = FlatStyle.Popup;
            button1.Font = new Font("Impact", 9.75F, FontStyle.Regular, GraphicsUnit.Point);
            button1.ForeColor = Color.FromArgb(255, 128, 0);
            button1.Location = new Point(0, 60);
            button1.Name = "button1";
            button1.Size = new Size(75, 23);
            button1.TabIndex = 0;
            button1.Text = "<<<";
            button1.UseVisualStyleBackColor = false;
            // 
            // Play
            // 
            Play.BackColor = Color.FromArgb(0, 192, 0);
            Play.FlatStyle = FlatStyle.Popup;
            Play.Font = new Font("Impact", 9.75F, FontStyle.Italic, GraphicsUnit.Point);
            Play.ForeColor = Color.White;
            Play.Location = new Point(162, 60);
            Play.Name = "Play";
            Play.Size = new Size(75, 23);
            Play.TabIndex = 1;
            Play.Text = "Play";
            Play.UseVisualStyleBackColor = false;
            Play.Click += button2_Click;
            // 
            // button3
            // 
            button3.BackColor = Color.Black;
            button3.FlatStyle = FlatStyle.Popup;
            button3.Font = new Font("Impact", 9.75F, FontStyle.Regular, GraphicsUnit.Point);
            button3.ForeColor = Color.FromArgb(255, 128, 0);
            button3.Location = new Point(243, 60);
            button3.Name = "button3";
            button3.Size = new Size(75, 23);
            button3.TabIndex = 2;
            button3.Text = ">>>";
            button3.UseVisualStyleBackColor = false;
            // 
            // ReadScreen
            // 
            ReadScreen.BackColor = Color.Black;
            ReadScreen.FlatStyle = FlatStyle.Popup;
            ReadScreen.Font = new Font("Impact", 9.75F, FontStyle.Italic, GraphicsUnit.Point);
            ReadScreen.ForeColor = Color.FromArgb(255, 128, 0);
            ReadScreen.Location = new Point(0, 12);
            ReadScreen.Name = "ReadScreen";
            ReadScreen.Size = new Size(75, 23);
            ReadScreen.TabIndex = 3;
            ReadScreen.Text = "Read";
            ReadScreen.UseVisualStyleBackColor = false;
            // 
            // button5
            // 
            button5.BackColor = Color.Red;
            button5.FlatStyle = FlatStyle.Popup;
            button5.Font = new Font("Impact", 9.75F, FontStyle.Italic, GraphicsUnit.Point);
            button5.ForeColor = Color.White;
            button5.Location = new Point(81, 60);
            button5.Name = "button5";
            button5.Size = new Size(75, 23);
            button5.TabIndex = 4;
            button5.Text = "Stop";
            button5.UseVisualStyleBackColor = false;
            // 
            // ProcessVM
            // 
            ProcessVM.BackColor = Color.Black;
            ProcessVM.FlatStyle = FlatStyle.Popup;
            ProcessVM.Font = new Font("Impact", 9.75F, FontStyle.Italic, GraphicsUnit.Point);
            ProcessVM.ForeColor = Color.FromArgb(255, 128, 0);
            ProcessVM.Location = new Point(243, 12);
            ProcessVM.Name = "ProcessVM";
            ProcessVM.Size = new Size(75, 23);
            ProcessVM.TabIndex = 5;
            ProcessVM.Text = "Refresh";
            ProcessVM.UseVisualStyleBackColor = false;
            ProcessVM.Click += button6_Click;
            // 
            // label1
            // 
            label1.AutoSize = true;
            label1.BackColor = Color.Black;
            label1.Font = new Font("Magneto", 24F, FontStyle.Bold, GraphicsUnit.Point);
            label1.ForeColor = Color.FromArgb(255, 128, 0);
            label1.Location = new Point(91, -3);
            label1.Name = "label1";
            label1.Size = new Size(137, 41);
            label1.TabIndex = 6;
            label1.Text = "Scribe";
            // 
            // panel1
            // 
            panel1.BackColor = Color.Black;
            panel1.Dock = DockStyle.Bottom;
            panel1.Location = new Point(0, 86);
            panel1.Name = "panel1";
            panel1.Size = new Size(318, 12);
            panel1.TabIndex = 7;
            // 
            // panel2
            // 
            panel2.BackColor = Color.Black;
            panel2.Controls.Add(ProcessVM);
            panel2.Controls.Add(ReadScreen);
            panel2.Controls.Add(label1);
            panel2.Dock = DockStyle.Top;
            panel2.Location = new Point(0, 0);
            panel2.Name = "panel2";
            panel2.Size = new Size(318, 35);
            panel2.TabIndex = 8;
            // 
            // textBox1
            // 
            textBox1.BackColor = Color.Black;
            textBox1.BorderStyle = BorderStyle.None;
            textBox1.Location = new Point(0, 41);
            textBox1.Name = "textBox1";
            textBox1.Size = new Size(318, 16);
            textBox1.TabIndex = 9;
            textBox1.TextChanged += textBox1_TextChanged;
            // 
            // MainForm
            // 
            AccessibleDescription = "Scribe (voicemail transcription)";
            AccessibleName = "Scribe";
            BackColor = Color.FromArgb(255, 128, 0);
            ClientSize = new Size(318, 98);
            Controls.Add(textBox1);
            Controls.Add(button5);
            Controls.Add(button3);
            Controls.Add(Play);
            Controls.Add(button1);
            Controls.Add(panel1);
            Controls.Add(panel2);
            Font = new Font("Magneto", 9.75F, FontStyle.Bold, GraphicsUnit.Point);
            ForeColor = Color.White;
            FormBorderStyle = FormBorderStyle.FixedSingle;
            Name = "MainForm";
            Load += MainForm_Load;
            panel2.ResumeLayout(false);
            panel2.PerformLayout();
            ResumeLayout(false);
            PerformLayout();
        }

        public void ProcessEmail(MailMessage email)
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



        private Button button1;
        private Button Play;
        private Button button3;
        private Button ReadScreen;
        private Button button5;
        private Label label1;
        private Panel panel1;
        private Panel panel2;
        private TextBox textBox1;
        private Button ProcessVM;

        private void MainForm_Load(object sender, EventArgs e)
        {

        }

        private void button2_Click(object sender, EventArgs e)
        {
            // Assuming voicemail file path is stored in a string variable called "voicemailLocation"
            string voicemailLocation = @"C:\path\to\voicemail.wav";

            if (!File.Exists(voicemailLocation))
            {
                MessageBox.Show("Voicemail file not found.");
                return;
            }

            // Initialize a new instance of the SoundPlayer class with the voicemail file path
            SoundPlayer player = new SoundPlayer(voicemailLocation);

            try
            {
                // Play the voicemail
                player.Play();
            }
            catch (Exception ex)
            {
                MessageBox.Show("Error playing voicemail: " + ex.Message);
            }
        }

        private void button4_Click(object sender, EventArgs e)
        {
            MessageBox.Show("Voicemail Transcription");
        }

        private void button6_Click(object sender, EventArgs e)
        {
            ProcessEmailsButton_Click(sender, e);
        }

        private void textBox1_TextChanged(object sender, EventArgs e)
        {

        }


        static void PopUp(string transcription)
        {
            // Create a new instance of the form to display the transcription
            Form transcriptionForm = new Form();
            transcriptionForm.Text = "Voicemail Transcription";
            transcriptionForm.Size = new Size(400, 300);

            // Create a label to display the transcription
            Label transcriptionLabel = new Label();
            transcriptionLabel.Text = transcription;
            transcriptionLabel.Dock = DockStyle.Fill;
            transcriptionLabel.TextAlign = ContentAlignment.MiddleCenter;

            // Add the label to the form
            transcriptionForm.Controls.Add(transcriptionLabel);

            // Display the form as a dialog box
            transcriptionForm.ShowDialog();
        }

    }

}
