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
using System.Reflection;
using NAudio.Wave;
using System.Configuration;
using Google.Cloud.Speech.V1;
using Grpc.Core;
using System.Net.NetworkInformation;

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

        private static async Task<bool> PingAsync()
        {
            var hostUrl = "imap.mail.yahoo.com";

            Ping ping = new Ping();

            PingReply result = await ping.SendPingAsync(hostUrl);
            return result.Status == IPStatus.Success;
            Console.WriteLine(result);
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
            var host = _configuration["Email:Host"] ?? "imap.mail.yahoo.com";
            var port = Convert.ToInt32(_configuration["Email:IncomingPort"]);
            var useSsl = Convert.ToBoolean(_configuration["Email:UseSsl"]);
            client.Connect(host, port, useSsl);
            client.Authenticate(_configuration["Email:UserName"], _configuration["Email:Password"]);
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
            Previous = new Button();
            Play = new Button();
            Next = new Button();
            ReadScreen = new Button();
            Stop = new Button();
            ProcessVM = new Button();
            label1 = new Label();
            panel1 = new Panel();
            panel2 = new Panel();
            textBox1 = new TextBox();
            panel2.SuspendLayout();
            SuspendLayout();
            // 
            // Previous
            // 
            Previous.BackColor = Color.Black;
            Previous.FlatStyle = FlatStyle.Popup;
            Previous.Font = new Font("Impact", 9.75F, FontStyle.Regular, GraphicsUnit.Point);
            Previous.ForeColor = Color.FromArgb(255, 128, 0);
            Previous.Location = new Point(0, 60);
            Previous.Name = "Previous";
            Previous.Size = new Size(75, 23);
            Previous.TabIndex = 0;
            Previous.Text = "<<<";
            Previous.UseVisualStyleBackColor = false;
            Previous.Click += button1_Click;
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
            // Next
            // 
            Next.BackColor = Color.Black;
            Next.FlatStyle = FlatStyle.Popup;
            Next.Font = new Font("Impact", 9.75F, FontStyle.Regular, GraphicsUnit.Point);
            Next.ForeColor = Color.FromArgb(255, 128, 0);
            Next.Location = new Point(243, 60);
            Next.Name = "Next";
            Next.Size = new Size(75, 23);
            Next.TabIndex = 2;
            Next.Text = ">>>";
            Next.UseVisualStyleBackColor = false;
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
            ReadScreen.Click += button4_Click;
            // 
            // Stop
            // 
            Stop.BackColor = Color.Red;
            Stop.FlatStyle = FlatStyle.Popup;
            Stop.Font = new Font("Impact", 9.75F, FontStyle.Italic, GraphicsUnit.Point);
            Stop.ForeColor = Color.White;
            Stop.Location = new Point(81, 60);
            Stop.Name = "Stop";
            Stop.Size = new Size(75, 23);
            Stop.TabIndex = 4;
            Stop.Text = "Stop";
            Stop.UseVisualStyleBackColor = false;
            Stop.Click += button5_Click;
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
            textBox1.Location = new Point(0, 38);
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
            Controls.Add(Stop);
            Controls.Add(Next);
            Controls.Add(Play);
            Controls.Add(Previous);
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



        private Button Previous;
        private Button Play;
        private Button Next;
        private Button ReadScreen;
        private Button Stop;
        private Label label1;
        private Panel panel1;
        private Panel panel2;
        private TextBox textBox1;
        private Button ProcessVM;

        private void MainForm_Load(object sender, EventArgs e)
        {

        }

        private List<string> voicemailList = new List<string>();

        private void LoadVoicemails()
        {
            // Assuming voicemails are stored in a directory called "Voicemails"
            string voicemailDirectory = @"C:\path\to\voicemails";

            // Get all files with the .wav extension in the directory
            string[] voicemailFiles = Directory.GetFiles(voicemailDirectory, "*.wav");

            // Add the file paths to the voicemailList
            voicemailList.AddRange(voicemailFiles);
        }

        private int currentVoicemailIndex = 0;

        private void button1_Click(object sender, EventArgs e)
        {
            // Decrement the current voicemail index to go back to the previous voicemail
            currentVoicemailIndex--;

            // Check if the current voicemail index is less than zero, indicating that we've reached the beginning of the list
            if (currentVoicemailIndex < 0)
            {
                MessageBox.Show("This is the first voicemail.");
                currentVoicemailIndex = 0;
                return;
            }

            // Assuming voicemail file path is stored in a string variable called "voicemailLocation"
            string voicemailLocation = voicemailList[currentVoicemailIndex];

            if (!File.Exists(voicemailLocation))
            {
                MessageBox.Show("Voicemail file not found.");
                return;
            }

            PlayMessage(voicemailLocation);
        }

        private void button3_Click(object sender, EventArgs e)
        {
            if (currentVoicemailIndex >= voicemailList.Count - 1)
            {
                MessageBox.Show("No more voicemails.");
                return;
            }

            currentVoicemailIndex++;
            string nextVoicemail = voicemailList[currentVoicemailIndex];
            PlayMessage(nextVoicemail);
        }


        private void button4_Click(object sender, EventArgs e)
        {
            string transcription = "This is a voicemail transcription."; // replace with actual voicemail transcription
            ShowVoicemailTranscription(transcription);
        }

        private SoundPlayer player; // declare the SoundPlayer object at the class level

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
            player = new SoundPlayer(voicemailLocation);

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

        private void button5_Click(object sender, EventArgs e)
        {
            if (player != null)
            {
                player.Stop();
            }
            else
            {
                MessageBox.Show("There is no voicemail currently playing.");
            }
        }

        private void button6_Click(object sender, EventArgs e)
        {
            _ = PingAsync();

            ProcessEmailsButton_Click(sender, e);
        }

        private WaveOutEvent waveOutEvent;
        private AudioFileReader audioFileReader;

        private void PlayMessage(string filePath)
        {
            StopMessage();

            audioFileReader = new AudioFileReader(filePath);
            waveOutEvent = new WaveOutEvent();
            waveOutEvent.Init(audioFileReader);
            waveOutEvent.Play();

            UpdateTextBox();
        }

        private void StopMessage()
        {
            if (waveOutEvent != null)
            {
                waveOutEvent.Stop();
                waveOutEvent.Dispose();
                waveOutEvent = null;
            }

            if (audioFileReader != null)
            {
                audioFileReader.Dispose();
                audioFileReader = null;
            }

            UpdateTextBox();
        }

        private void UpdateTextBox()
        {
            if (audioFileReader != null && waveOutEvent != null)
            {
                var fileName = Path.GetFileName(audioFileReader.FileName);
                var currentPosition = audioFileReader.CurrentTime;
                var totalDuration = audioFileReader.TotalTime;
                var timeRemaining = totalDuration - currentPosition;

                textBox1.Text = $"Playing: {fileName}\nDuration: {totalDuration}\nTime remaining: {timeRemaining}";
            }
            else
            {
                textBox1.Text = "No file playing";
            }
        }

        private void textBox1_TextChanged(object sender, EventArgs e)
        {
            UpdateTextBox();
        }

        private TimeSpan GetVoicemailDuration(string voicemailLocation)
        {
            // Create a new instance of AudioFileReader with the voicemail file path
            using var reader = new AudioFileReader(voicemailLocation);

            // Get the duration of the voicemail
            return reader.TotalTime;
        }


        private void ShowVoicemailTranscription(string transcription)
        {
            // Create a new form
            Form voicemailForm = new Form();

            // Set the title of the form
            voicemailForm.Text = "Voicemail Transcription";

            // Create a label to display the transcription
            Label transcriptionLabel = new Label();
            transcriptionLabel.Text = transcription;
            transcriptionLabel.AutoSize = true;
            transcriptionLabel.Dock = DockStyle.Fill;
            transcriptionLabel.TextAlign = ContentAlignment.MiddleCenter;

            // Add the label to the form
            voicemailForm.Controls.Add(transcriptionLabel);

            // Set the size of the form based on the size of the label
            voicemailForm.ClientSize = new Size(transcriptionLabel.Width + 20, transcriptionLabel.Height + 20);

            // Display the form as a dialog
            voicemailForm.ShowDialog();
        }

    }

}
