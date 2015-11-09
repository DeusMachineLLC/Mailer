using Emailer.Properties;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Net;
using System.Net.Mail;
using System.Net.Security;
using System.Net.Sockets;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Microsoft.VisualBasic;
using System.Management;

namespace Emailer
{
    public partial class DeusEmailer : Form
    {
        // readonly means it can only be set in the constructor
        // i like to name properties (variables) and methods that are going to
        // be private with an underscore prefix to help make the code readable
        private readonly DeusLabel _nameLabel;
        private readonly DeusTextBox _nameField;

        private readonly DeusLabel _emailLabel;
        private readonly DeusTextBox _emailField;

        private readonly DeusLabel _phoneLabel;
        private readonly DeusTextBox _phoneField;

        private readonly DeusLabel _cellLabel;
        private readonly DeusTextBox _cellField;

        private readonly DeusLabel _severityLabel;
        private readonly GroupBox _severityGroupBox;
        private readonly System.Collections.Specialized.StringCollection _severityOptions;

        private readonly DeusLabel _descriptionLabel;
        private readonly DeusTextBox _descriptionField;

        private readonly DeusLabel _noInternetPhone;

        private readonly Button _submitButton;

        private readonly bool _connectedToInternet;

        private RadioButton _checkedButton;

        private readonly PictureBox _brandImage;

        private readonly DeusLabel _onlineSupportLink;

        public DeusEmailer()
        {
            _nameLabel = new DeusLabel(new Point(25, 25), "Your Name*");
            _nameField = new DeusTextBox(new Point(25, 55));
            
            _emailLabel = new DeusLabel(new Point(25, 100), "Your Email Address*");
            _emailField = new DeusTextBox(new Point(25, 130));
            
            _phoneLabel = new DeusLabel(new Point(25, 175), "Phone Extension");
            _phoneField = new DeusTextBox(new Point(25, 205), new Size(100, 28), new Font(TextBox.DefaultFont.FontFamily, 12));
            
            _cellLabel = new DeusLabel(new Point(25, 250), "Cellphone");
            _cellField = new DeusTextBox(new Point(25, 280));
            
            _severityLabel = new DeusLabel(new Point(25, 325), "Issue Severity (click a bubble*)");
            _severityGroupBox = new GroupBox()
            {
                Height = 35
            };
            _severityGroupBox.Location = new Point(25, 355);

            _severityOptions = Settings.Default.IssueSeverityOptions;

            _descriptionLabel = new DeusLabel(new Point(25, 415), "Description of Issue*");
            _descriptionField = new DeusTextBox(new Point(25, 445), new Size(400, 28), new Font(TextBox.DefaultFont.FontFamily, 10))
            {
                Height = 90
            };

            _AddTextControls();

            _AddRadioControls(24, 37, 7);

            _connectedToInternet = _CanPingGoogleMail();
            if (false == _connectedToInternet)
            {
                DialogResult nope = MessageBox.Show("Not Connected To Internet");

                string msg = "Please call " + Settings.Default.NoInternetPhone + " for assistance";
                _noInternetPhone = new DeusLabel(new Point(25, 550), msg);

                this.Controls.Add(_noInternetPhone);
            }
            else
            {
                _submitButton = new Button()
                {
                    Text = "Submit Ticket",
                    Location = new Point(25, 550),
                    Height = 35,
                    Width = 100,
                    ForeColor = Color.Black,
                    BackColor = Color.AliceBlue
                };

                _submitButton.Click += new EventHandler(this.SubmitButtonClicked);

                this.Controls.Add(_submitButton);
            }

            _brandImage = new PictureBox()
            {
                ImageLocation = Settings.Default.BrandImagePath,
                SizeMode = PictureBoxSizeMode.AutoSize,
                Location = new Point(468, 25)
            };
            this.Controls.Add(_brandImage);

            string supportLinkText = Settings.Default.OnlineSupportLink;
            if (!string.IsNullOrEmpty(supportLinkText))
            {
                _onlineSupportLink = new DeusLabel(new Point(25, 650), supportLinkText);
                this.Controls.Add(_onlineSupportLink);
            }

            InitializeComponent();
        }

        private void _AddRadioControls(int radioWidth, int radioOffset, int yOffset)
        {
            int i = 10;
            foreach (string option in _severityOptions)
            {
                RadioButton radio = new RadioButton()
                {
                    Text = option,
                    Location = new Point(i, yOffset),
                    Width = radioWidth
                };
                _severityGroupBox.Controls.Add(radio);
                i += radioOffset;
            }
        }

        private void _AddTextControls()
        {
            this.Controls.Add(_nameLabel);
            this.Controls.Add(_nameField);

            this.Controls.Add(_emailLabel);
            this.Controls.Add(_emailField);

            this.Controls.Add(_phoneLabel);
            this.Controls.Add(_phoneField);

            this.Controls.Add(_cellLabel);
            this.Controls.Add(_cellField);

            this.Controls.Add(_severityLabel);
            this.Controls.Add(_severityGroupBox);

            this.Controls.Add(_descriptionLabel);
            this.Controls.Add(_descriptionField);
        }

        public void SubmitButtonClicked(object sender, EventArgs e)
        {
            if (false == _ValidateForm())
            {
                string error = "Please fill out all fields marked as required (*) and make sure your email address is valid";
                DialogResult missingFields = MessageBox.Show(error);
            }
            else
            {
                MailMessage message = new MailMessage(_emailField.Text, Settings.Default.HelpDeskEmail);

                StringBuilder body = new StringBuilder(Settings.Default.Prepend);

                body.AppendLine("\n\n Name: " + _nameField.Text);
                body.AppendLine("\n\n Email: " + _emailField.Text);

                string phone = string.IsNullOrEmpty(_phoneField.Text) ? "Field Empty" : _phoneField.Text;
                body.AppendLine("\n\n Phone Ext: " + phone);

                string cell = string.IsNullOrEmpty(_cellField.Text) ? "Field Empty" : _cellField.Text;
                body.AppendLine("\n\n Cell: " + cell);

                body.AppendLine("\n\n Issue Severity: " + _checkedButton.Text);
                body.AppendLine("\n\n Description:\n" + _descriptionField.Text);

                ulong memory = new Microsoft.VisualBasic.Devices.ComputerInfo().TotalPhysicalMemory;
                string showMemory;

                if (memory > 1000000000)
                {
                    showMemory = (memory / 1000000000).ToString() + " gigabytes";
                }
                else
                {
                    showMemory = (memory / 1000000).ToString() + " megabytes";
                }

                body.AppendLine("\n\n Total Memory: " + showMemory);

                string bitCount;
                if (Environment.Is64BitOperatingSystem)
                    bitCount = " 64 bit";
                else
                    bitCount = " 32 bit";

                string processorName = "Not Available";
                try
                {

                    ManagementObjectCollection compInfo = new ManagementObjectSearcher("root\\CIMV2", "SELECT * FROM Win32_Processor").Get();
                    foreach (ManagementObject obj in compInfo)
                    {
                        processorName = obj["Name"].ToString();
                    }
                }
                catch { }

                body.AppendLine("\n\n OS Version: " + Environment.OSVersion.ToString() + bitCount);

                body.AppendLine("\n\n Processor: " + processorName);
                body.AppendLine("\n\n Processor Architecture: " + Environment.GetEnvironmentVariable("PROCESSOR_ARCHITECTURE"));
                body.AppendLine("\n\n Number of Processors: " + Environment.ProcessorCount.ToString());
               
                body.AppendLine("\n\n Common Language Runtime Version: " + Environment.Version.ToString());

                message.Body = body.ToString();

                DialogResult dialogResult = MessageBox.Show("Submit Ticket", "Cancel", MessageBoxButtons.YesNo);
                if (dialogResult == DialogResult.Yes)
                {
                    if (false == _SendEmail(message))
                    {
                        DialogResult fail = MessageBox.Show("Ticket Submssion Failed, Unable To Send Mail");
                    }
                    else
                    {
                        DialogResult success = MessageBox.Show("Ticket Submitted Successfully");
                    }
                }
            }
        }

        private bool _SendEmail(MailMessage message)
        {

            SmtpClient client = new SmtpClient(Settings.Default.HelpDeskHost, 587)
            {
                Credentials = new NetworkCredential(Settings.Default.HostUserName, Settings.Default.HostPassword),
                EnableSsl = true
            };

            try
            {
                client.Send(Settings.Default.HelpDeskEmail, Settings.Default.HelpDeskEmail, "New Helpdesk Ticket", message.Body);
            }
            catch (Exception)
            {
                return false;
            }

            return true;
        }

        private bool _ValidateForm()
        {
            if
            (
                string.IsNullOrEmpty(_emailField.Text) ||
                !_emailField.Text.Contains("@") ||
                string.IsNullOrEmpty(_nameField.Text) ||
                string.IsNullOrEmpty(_descriptionField.Text)
            )
                return false;

            // this is probably the hardest line in here Lee, let me know if it doesn't make sense (it's linq)
            _checkedButton = _severityGroupBox.Controls.OfType<RadioButton>().FirstOrDefault(x => x.Checked);
            if (null == _checkedButton)
                return false;

            return true;
        }

        private bool _CanPingGoogleMail()
        {
            try
            {
                using (var client = new TcpClient())
                {
                    var server = "smtp.gmail.com";
                    var port = 465;
                    client.Connect(server, port);
                    using (var stream = client.GetStream())
                    using (var secureStream = new SslStream(stream))
                    {
                        secureStream.AuthenticateAsClient(server);
                        using (var writer = new StreamWriter(secureStream))
                        using (var reader = new StreamReader(secureStream))
                        {
                            writer.WriteLine("EHLO " + server);
                            writer.Flush();

                            string result = reader.ReadLine();

                            return (result.Contains("220 smtp.gmail.com ESMTP")) ? true : false;
                        }
                    }
                }
            }
            catch(Exception)
            {
                return false;
            }
        }

    }
    
    // : means extends in c#
    public class DeusLabel : Label
    {
        // you can overload methods in c# as well as constructors, so I can declare multiple
        // constructors that take different numbers and kinds of parameters
        public DeusLabel(Point point, Font font, string text)
        {
            this.AutoSize = true;
            this.Location = point;
            this.Font = font;
            this.Text = text;
        }

        public DeusLabel(Point point, string text)
        {
            this.AutoSize = true;
            this.Location = point;
            this.Font = new Font(this.Font.FontFamily, 14);
            this.Text = text;
        }
    }

    public class DeusTextBox : TextBox
    {
        public bool required;

        public DeusTextBox(Point point)
        {
            this.Multiline = true;
            this.Size = new Size(300, 28);
            this.Font = new Font(this.Font.FontFamily, 12);

            this.Location = point;
        }

        public DeusTextBox(Point point, Size size, Font font)
        {
            this.Multiline = true;
            this.Size = size;
            this.Font = font;

            this.Location = point;
        }

    }
}