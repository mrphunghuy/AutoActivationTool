using System;
using System.Drawing;
using System.Text;
using System.Windows.Forms;
using System.Runtime.InteropServices;
using System.IO;

namespace AutoActivationTool
{
    [System.ComponentModel.DesignerCategory("")]
    public class TextBoxStreamWriter : TextWriter
    {
        TextBox _output = null;

        public TextBoxStreamWriter(TextBox output)
        {
            _output = output;
        }

        public override void Write(char value)
        {
            base.Write(value);
            _output.AppendText(value.ToString()); // When character data is written, append it to the text box.
        }

        public override Encoding Encoding
        {
            get { return System.Text.Encoding.UTF8; }
        }
    }
    public partial class ActivationForm : Form
    {
        [DllImport("user32.dll")]
        static extern bool HideCaret(IntPtr hWnd);

        private System.Windows.Forms.Label userNameLabel;
        private System.Windows.Forms.Label passwordLabel;
        private System.Windows.Forms.TextBox userNameTextbox;
        private System.Windows.Forms.TextBox passwordTextbox;
        private System.Windows.Forms.Label status;
        private System.Windows.Forms.Button loginButton;
        private System.Windows.Forms.LinkLabel registerLink;
        private System.Windows.Forms.CheckBox rememberCheckbox;
        private System.Windows.Forms.ComboBox keyTypeCombobox;
        private System.Windows.Forms.Button activateButton;
        private TextWriter consoleRedirect;
        private System.Windows.Forms.TextBox consoleTextbox;
        public ActivationForm()
        {
            string version = "1.6.0";
            InitializeComponent(version);
            InitializeKeyTypeCombobox();

            Common.status = this.status;
            OfficeActivator.status = this.status;

            Common.loadAccount(this.userNameTextbox, this.passwordTextbox);
        }
        private void InitializeComponent(string version)
        {
            this.userNameLabel = new System.Windows.Forms.Label();
            this.passwordLabel = new System.Windows.Forms.Label();
            this.userNameTextbox = new System.Windows.Forms.TextBox();
            this.passwordTextbox = new System.Windows.Forms.TextBox();
            this.status = new System.Windows.Forms.Label();
            this.loginButton = new System.Windows.Forms.Button();
            this.registerLink = new System.Windows.Forms.LinkLabel();
            this.rememberCheckbox = new System.Windows.Forms.CheckBox();
            this.keyTypeCombobox = new System.Windows.Forms.ComboBox();
            this.activateButton = new System.Windows.Forms.Button();
            this.consoleTextbox = new System.Windows.Forms.TextBox();
            this.consoleRedirect = new TextBoxStreamWriter(this.consoleTextbox);
            Console.SetOut(consoleRedirect);
            this.SuspendLayout();
            // 
            // label1
            // 
            this.userNameLabel.AutoSize = true;
            this.userNameLabel.Font = new System.Drawing.Font("Arial", 9.75F);
            this.userNameLabel.Location = new System.Drawing.Point(55, 29);
            this.userNameLabel.Name = "label1";
            this.userNameLabel.Size = new System.Drawing.Size(73, 16);
            this.userNameLabel.TabIndex = 0;
            this.userNameLabel.Text = "User Name";
            // 
            // label2
            // 
            this.passwordLabel.AutoSize = true;
            this.passwordLabel.Font = new System.Drawing.Font("Arial", 9.75F);
            this.passwordLabel.Location = new System.Drawing.Point(55, 79);
            this.passwordLabel.Name = "label2";
            this.passwordLabel.Size = new System.Drawing.Size(65, 16);
            this.passwordLabel.TabIndex = 1;
            this.passwordLabel.Text = "Password";
            // 
            // textBox1
            // 
            this.userNameTextbox.Font = new System.Drawing.Font("Arial", 9.75F);
            this.userNameTextbox.Location = new System.Drawing.Point(134, 26);
            this.userNameTextbox.Name = "textBox1";
            this.userNameTextbox.Size = new System.Drawing.Size(186, 22);
            this.userNameTextbox.TabIndex = 2;
            // 
            // textBox2
            // 
            this.passwordTextbox.Font = new System.Drawing.Font("Arial", 9.75F);
            this.passwordTextbox.Location = new System.Drawing.Point(134, 76);
            this.passwordTextbox.Name = "textBox2";
            this.passwordTextbox.PasswordChar = '*';
            this.passwordTextbox.Size = new System.Drawing.Size(186, 22);
            this.passwordTextbox.TabIndex = 3;
            // 
            // label3
            // 
            this.status.AutoSize = false;
            this.status.Font = new System.Drawing.Font("Arial", 9.75F);
            this.status.Location = new System.Drawing.Point(0, 220);
            this.status.Name = "label3";
            this.status.Size = new System.Drawing.Size(400, 16);
            this.status.TabIndex = 4;
            this.status.Text = "Ready";
            this.status.TextAlign = ContentAlignment.MiddleCenter;
            // 
            // button1
            // 
            this.loginButton.Font = new System.Drawing.Font("Arial", 9.75F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.loginButton.Location = new System.Drawing.Point(146, 164);
            this.loginButton.Name = "button1";
            this.loginButton.Size = new System.Drawing.Size(96, 27);
            this.loginButton.TabIndex = 5;
            this.loginButton.Text = "Login";
            this.loginButton.UseVisualStyleBackColor = true;
            this.loginButton.Click += new System.EventHandler(this.OnLoginButton_Click);
            //
            // link label
            //
            this.registerLink.Font = new System.Drawing.Font("Arial", 8.75F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.registerLink.Location = new System.Drawing.Point(316, 230);
            this.registerLink.Name = "button1";
            this.registerLink.Size = new System.Drawing.Size(150, 18);
            this.registerLink.TabIndex = 5;
            this.registerLink.Text = "Register";
            this.registerLink.Click += new System.EventHandler(this.OnRegisterLabel_Click);
            // 
            // checkBox1
            // 
            this.rememberCheckbox.AutoSize = true;
            this.rememberCheckbox.Font = new System.Drawing.Font("Arial", 9.75F);
            this.rememberCheckbox.Location = new System.Drawing.Point(93, 121);
            this.rememberCheckbox.Name = "checkBox1";
            this.rememberCheckbox.Size = new System.Drawing.Size(204, 20);
            this.rememberCheckbox.TabIndex = 7;
            this.rememberCheckbox.Text = "Save user name and password";
            this.rememberCheckbox.UseVisualStyleBackColor = true;
            // 
            // comboBox1
            // 
            this.keyTypeCombobox.Font = new System.Drawing.Font("Arial", 9.75F);
            this.keyTypeCombobox.FormattingEnabled = true;
            this.keyTypeCombobox.Location = new System.Drawing.Point(125, 21);
            this.keyTypeCombobox.Name = "comboBox1";
            this.keyTypeCombobox.Size = new System.Drawing.Size(120, 27);
            this.keyTypeCombobox.TabIndex = 8;
            this.keyTypeCombobox.DropDownStyle = ComboBoxStyle.DropDownList;
            this.keyTypeCombobox.Hide();
            // 
            // button2
            // 
            this.activateButton.Font = new System.Drawing.Font("Arial", 9.75F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.activateButton.Location = new System.Drawing.Point(255, 20);
            this.activateButton.Name = "button2";
            this.activateButton.Size = new System.Drawing.Size(96, 26);
            this.activateButton.TabIndex = 10;
            this.activateButton.Text = "Activate";
            this.activateButton.UseVisualStyleBackColor = true;
            //this.activateButton.ForeColor = Color.Red;
            this.activateButton.Click += OnActivateButton_Click;
            this.activateButton.Hide();
            // 
            // textBox1
            // 
            this.consoleTextbox.Font = new System.Drawing.Font("Consolas", 10.0F);
            this.consoleTextbox.Location = new System.Drawing.Point(5, 60);
            this.consoleTextbox.Name = "textBox1";
            this.consoleTextbox.Size = new System.Drawing.Size(490, 335);
            this.consoleTextbox.TabIndex = 2;
            this.consoleTextbox.Multiline = true;
            this.consoleTextbox.ScrollBars = ScrollBars.Vertical;
            this.consoleTextbox.ReadOnly = true;
            this.consoleTextbox.BackColor = Color.Black;
            this.consoleTextbox.ForeColor = Color.Red;
            this.consoleTextbox.TextChanged += consoleTextbox_TextChanged;
            this.consoleTextbox.Hide();
            // 
            // Form1
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(384, 261);
            this.FormBorderStyle = FormBorderStyle.FixedDialog;
            this.MaximizeBox = false;

            this.Controls.Add(this.keyTypeCombobox);
            this.Controls.Add(this.activateButton);
            this.Controls.Add(this.consoleTextbox);
            this.Controls.Add(this.rememberCheckbox);
            this.Controls.Add(this.loginButton);
            this.Controls.Add(this.registerLink);
            this.Controls.Add(this.status);
            this.Controls.Add(this.passwordTextbox);
            this.Controls.Add(this.userNameTextbox);
            this.Controls.Add(this.passwordLabel);
            this.Controls.Add(this.userNameLabel);
            this.Name = "Form1";
            this.Text = "WinOffice.Org - Auto Activation Tool v" + version;
            this.ResumeLayout(false);
            this.PerformLayout();
        }

        private void InitializeKeyTypeCombobox()
        {
            this.keyTypeCombobox.Items.Add("Windows");
            this.keyTypeCombobox.Items.Add("Office");
            this.keyTypeCombobox.SelectedIndex = 0;
        }
        void OnLoginButton_Click(object sender, EventArgs e)
        {
            string username = this.userNameTextbox.Text;
            string password = this.passwordTextbox.Text;

            bool res = Common.Login(username, password);
            if(res)
            {
                if(this.rememberCheckbox.Checked)
                {
                    Common.saveAccount(username, password);
                }

                foreach(Control control in this.Controls)
                {
                    control.Visible = !control.Visible;
                }

                this.ClientSize = new System.Drawing.Size(500, 400);
            }
        }

        private void consoleTextbox_TextChanged(object sender, EventArgs e)
        {
            HideCaret(this.consoleTextbox.Handle);
        }

        void OnRegisterLabel_Click(object sender, EventArgs e)
        {
            registerLink.LinkVisited = true;
            System.Diagnostics.Process.Start("https://www.winoffice.org/register");
        }

        void OnActivateButton_Click(object sender, EventArgs e)
        {
            this.activateButton.Enabled = false;
            this.activateButton.Text = "Activating...";
            Application.DoEvents();

            this.consoleTextbox.Clear();
            Application.DoEvents();

            if (this.keyTypeCombobox.SelectedIndex == 0) //Windows
            {
                WinActivator.Activate();
            }
            else //Office
            {
                OfficeActivator.Activate();
            }

            Console.WriteLine("Finished.");

            this.activateButton.Enabled = true;
            this.activateButton.Text = "Activate";
        }
    }
}
