namespace MosTrainer
{
    partial class LoginForm
    {
        /// <summary>
        /// Required designer variable.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        /// <summary>
        /// Clean up any resources being been used.
        /// </summary>
        /// <param name="disposing">true if managed resources should be disposed; otherwise, false.</param>
        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        #region Windows Form Designer generated code

        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InitializeComponent()
        {
            this.tlpRoot = new System.Windows.Forms.TableLayoutPanel();
            this.tlpCard = new System.Windows.Forms.TableLayoutPanel();
            this.lblTitle = new System.Windows.Forms.Label();
            this.tlpUser = new System.Windows.Forms.TableLayoutPanel();
            this.lblUsername = new System.Windows.Forms.Label();
            this.txtUsername = new System.Windows.Forms.TextBox();
            this.tlpPass = new System.Windows.Forms.TableLayoutPanel();
            this.lblPassword = new System.Windows.Forms.Label();
            this.txtPassword = new System.Windows.Forms.TextBox();
            this.flpLang = new System.Windows.Forms.FlowLayoutPanel();
            this.chkEN = new System.Windows.Forms.CheckBox();
            this.chkVI = new System.Windows.Forms.CheckBox();
            this.tlpButtons = new System.Windows.Forms.TableLayoutPanel();
            this.btnLogin = new System.Windows.Forms.Button();
            this.btnExit = new System.Windows.Forms.Button();
            this.lblStatus = new System.Windows.Forms.Label();
            this.tlpRoot.SuspendLayout();
            this.tlpCard.SuspendLayout();
            this.tlpUser.SuspendLayout();
            this.tlpPass.SuspendLayout();
            this.flpLang.SuspendLayout();
            this.tlpButtons.SuspendLayout();
            this.SuspendLayout();
            // 
            // tlpRoot
            // 
            this.tlpRoot.ColumnCount = 3;
            this.tlpRoot.ColumnStyles.Add(new System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Percent, 15.41667F));
            this.tlpRoot.ColumnStyles.Add(new System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Percent, 66.45834F));
            this.tlpRoot.ColumnStyles.Add(new System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Percent, 18.125F));
            this.tlpRoot.Controls.Add(this.tlpCard, 1, 1);
            this.tlpRoot.Dock = System.Windows.Forms.DockStyle.Fill;
            this.tlpRoot.Location = new System.Drawing.Point(0, 0);
            this.tlpRoot.Name = "tlpRoot";
            this.tlpRoot.RowCount = 3;
            this.tlpRoot.RowStyles.Add(new System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Percent, 15.55556F));
            this.tlpRoot.RowStyles.Add(new System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Percent, 66.94444F));
            this.tlpRoot.RowStyles.Add(new System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Percent, 17.77778F));
            this.tlpRoot.Size = new System.Drawing.Size(480, 360);
            this.tlpRoot.TabIndex = 0;
            // 
            // tlpCard
            // 
            this.tlpCard.ColumnCount = 1;
            this.tlpCard.ColumnStyles.Add(new System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Percent, 100F));
            this.tlpCard.Controls.Add(this.lblTitle, 0, 0);
            this.tlpCard.Controls.Add(this.tlpUser, 0, 1);
            this.tlpCard.Controls.Add(this.tlpPass, 0, 2);
            this.tlpCard.Controls.Add(this.flpLang, 0, 3);
            this.tlpCard.Controls.Add(this.tlpButtons, 0, 4);
            this.tlpCard.Controls.Add(this.lblStatus, 0, 5);
            this.tlpCard.Dock = System.Windows.Forms.DockStyle.Fill;
            this.tlpCard.Location = new System.Drawing.Point(77, 58);
            this.tlpCard.Name = "tlpCard";
            this.tlpCard.Padding = new System.Windows.Forms.Padding(12);
            this.tlpCard.RowCount = 6;
            this.tlpCard.RowStyles.Add(new System.Windows.Forms.RowStyle());
            this.tlpCard.RowStyles.Add(new System.Windows.Forms.RowStyle());
            this.tlpCard.RowStyles.Add(new System.Windows.Forms.RowStyle());
            this.tlpCard.RowStyles.Add(new System.Windows.Forms.RowStyle());
            this.tlpCard.RowStyles.Add(new System.Windows.Forms.RowStyle());
            this.tlpCard.RowStyles.Add(new System.Windows.Forms.RowStyle());
            this.tlpCard.Size = new System.Drawing.Size(313, 234);
            this.tlpCard.TabIndex = 1;
            // 
            // lblTitle
            // 
            this.lblTitle.Dock = System.Windows.Forms.DockStyle.Fill;
            this.lblTitle.Font = new System.Drawing.Font("Segoe UI", 18F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.lblTitle.Location = new System.Drawing.Point(15, 12);
            this.lblTitle.Name = "lblTitle";
            this.lblTitle.Size = new System.Drawing.Size(283, 45);
            this.lblTitle.TabIndex = 0;
            this.lblTitle.Text = "Đăng nhập";
            this.lblTitle.TextAlign = System.Drawing.ContentAlignment.MiddleCenter;
            // 
            // tlpUser
            // 
            this.tlpUser.ColumnCount = 2;
            this.tlpUser.ColumnStyles.Add(new System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Percent, 35F));
            this.tlpUser.ColumnStyles.Add(new System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Percent, 65F));
            this.tlpUser.Controls.Add(this.lblUsername, 0, 0);
            this.tlpUser.Controls.Add(this.txtUsername, 1, 0);
            this.tlpUser.Dock = System.Windows.Forms.DockStyle.Fill;
            this.tlpUser.Location = new System.Drawing.Point(15, 60);
            this.tlpUser.Name = "tlpUser";
            this.tlpUser.RowCount = 1;
            this.tlpUser.RowStyles.Add(new System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Percent, 100F));
            this.tlpUser.Size = new System.Drawing.Size(283, 30);
            this.tlpUser.TabIndex = 1;
            // 
            // lblUsername
            // 
            this.lblUsername.Dock = System.Windows.Forms.DockStyle.Fill;
            this.lblUsername.Location = new System.Drawing.Point(3, 0);
            this.lblUsername.Name = "lblUsername";
            this.lblUsername.Size = new System.Drawing.Size(93, 30);
            this.lblUsername.TabIndex = 0;
            this.lblUsername.Text = "Tài khoản:";
            this.lblUsername.TextAlign = System.Drawing.ContentAlignment.MiddleRight;
            // 
            // txtUsername
            // 
            this.txtUsername.Dock = System.Windows.Forms.DockStyle.Fill;
            this.txtUsername.Font = new System.Drawing.Font("Segoe UI", 10F);
            this.txtUsername.Location = new System.Drawing.Point(102, 3);
            this.txtUsername.Name = "txtUsername";
            this.txtUsername.Size = new System.Drawing.Size(178, 30);
            this.txtUsername.TabIndex = 1;
            // 
            // tlpPass
            // 
            this.tlpPass.ColumnCount = 2;
            this.tlpPass.ColumnStyles.Add(new System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Percent, 35F));
            this.tlpPass.ColumnStyles.Add(new System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Percent, 65F));
            this.tlpPass.Controls.Add(this.lblPassword, 0, 0);
            this.tlpPass.Controls.Add(this.txtPassword, 1, 0);
            this.tlpPass.Dock = System.Windows.Forms.DockStyle.Fill;
            this.tlpPass.Location = new System.Drawing.Point(15, 96);
            this.tlpPass.Name = "tlpPass";
            this.tlpPass.RowCount = 1;
            this.tlpPass.RowStyles.Add(new System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Percent, 100F));
            this.tlpPass.Size = new System.Drawing.Size(283, 30);
            this.tlpPass.TabIndex = 2;
            // 
            // lblPassword
            // 
            this.lblPassword.Dock = System.Windows.Forms.DockStyle.Fill;
            this.lblPassword.Location = new System.Drawing.Point(3, 0);
            this.lblPassword.Name = "lblPassword";
            this.lblPassword.Size = new System.Drawing.Size(93, 30);
            this.lblPassword.TabIndex = 0;
            this.lblPassword.Text = "Mật khẩu:";
            this.lblPassword.TextAlign = System.Drawing.ContentAlignment.MiddleRight;
            // 
            // txtPassword
            // 
            this.txtPassword.Dock = System.Windows.Forms.DockStyle.Fill;
            this.txtPassword.Font = new System.Drawing.Font("Segoe UI", 10F);
            this.txtPassword.Location = new System.Drawing.Point(102, 3);
            this.txtPassword.Name = "txtPassword";
            this.txtPassword.Size = new System.Drawing.Size(178, 30);
            this.txtPassword.TabIndex = 1;
            this.txtPassword.UseSystemPasswordChar = true;
            // 
            // flpLang
            // 
            this.flpLang.Controls.Add(this.chkEN);
            this.flpLang.Controls.Add(this.chkVI);
            this.flpLang.Dock = System.Windows.Forms.DockStyle.Fill;
            this.flpLang.Location = new System.Drawing.Point(15, 132);
            this.flpLang.Name = "flpLang";
            this.flpLang.Size = new System.Drawing.Size(283, 30);
            this.flpLang.TabIndex = 3;
            // 
            // chkEN
            // 
            this.chkEN.AutoSize = true;
            this.chkEN.Font = new System.Drawing.Font("Segoe UI", 9F);
            this.chkEN.Location = new System.Drawing.Point(3, 3);
            this.chkEN.Name = "chkEN";
            this.chkEN.Size = new System.Drawing.Size(98, 24);
            this.chkEN.TabIndex = 0;
            this.chkEN.Text = "Tiếng Anh";
            this.chkEN.UseVisualStyleBackColor = true;
            // 
            // chkVI
            // 
            this.chkVI.AutoSize = true;
            this.chkVI.Font = new System.Drawing.Font("Segoe UI", 9F);
            this.chkVI.Location = new System.Drawing.Point(107, 3);
            this.chkVI.Name = "chkVI";
            this.chkVI.Size = new System.Drawing.Size(98, 24);
            this.chkVI.TabIndex = 1;
            this.chkVI.Text = "Tiếng Việt";
            this.chkVI.UseVisualStyleBackColor = true;
            // 
            // tlpButtons
            // 
            this.tlpButtons.ColumnCount = 2;
            this.tlpButtons.ColumnStyles.Add(new System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Percent, 50F));
            this.tlpButtons.ColumnStyles.Add(new System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Percent, 50F));
            this.tlpButtons.Controls.Add(this.btnLogin, 0, 0);
            this.tlpButtons.Controls.Add(this.btnExit, 1, 0);
            this.tlpButtons.Dock = System.Windows.Forms.DockStyle.Fill;
            this.tlpButtons.Location = new System.Drawing.Point(15, 168);
            this.tlpButtons.Name = "tlpButtons";
            this.tlpButtons.RowCount = 1;
            this.tlpButtons.RowStyles.Add(new System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Percent, 100F));
            this.tlpButtons.Size = new System.Drawing.Size(283, 40);
            this.tlpButtons.TabIndex = 4;
            // 
            // btnLogin
            // 
            this.btnLogin.BackColor = System.Drawing.Color.LightSteelBlue;
            this.btnLogin.Dock = System.Windows.Forms.DockStyle.Fill;
            this.btnLogin.Font = new System.Drawing.Font("Segoe UI", 10F, System.Drawing.FontStyle.Bold);
            this.btnLogin.Location = new System.Drawing.Point(3, 3);
            this.btnLogin.Name = "btnLogin";
            this.btnLogin.Size = new System.Drawing.Size(135, 34);
            this.btnLogin.TabIndex = 0;
            this.btnLogin.Text = "Đăng nhập";
            this.btnLogin.UseVisualStyleBackColor = false;
            // 
            // btnExit
            // 
            this.btnExit.DialogResult = System.Windows.Forms.DialogResult.Cancel;
            this.btnExit.Dock = System.Windows.Forms.DockStyle.Fill;
            this.btnExit.Location = new System.Drawing.Point(144, 3);
            this.btnExit.Name = "btnExit";
            this.btnExit.Size = new System.Drawing.Size(136, 34);
            this.btnExit.TabIndex = 1;
            this.btnExit.Text = "Exit";
            this.btnExit.UseVisualStyleBackColor = true;
            // 
            // lblStatus
            // 
            this.lblStatus.Dock = System.Windows.Forms.DockStyle.Fill;
            this.lblStatus.ForeColor = System.Drawing.Color.DarkRed;
            this.lblStatus.Location = new System.Drawing.Point(15, 211);
            this.lblStatus.Name = "lblStatus";
            this.lblStatus.Size = new System.Drawing.Size(283, 30);
            this.lblStatus.TabIndex = 5;
            this.lblStatus.TextAlign = System.Drawing.ContentAlignment.MiddleCenter;
            // 
            // LoginForm
            // 
            this.AcceptButton = this.btnLogin;
            this.AutoScaleDimensions = new System.Drawing.SizeF(8F, 16F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.BackColor = System.Drawing.Color.WhiteSmoke;
            this.CancelButton = this.btnExit;
            this.ClientSize = new System.Drawing.Size(480, 360);
            this.Controls.Add(this.tlpRoot);
            this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedDialog;
            this.MaximizeBox = false;
            this.Name = "LoginForm";
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
            this.Text = "Login";
            this.Load += new System.EventHandler(this.Lang_CheckedChanged);
            this.tlpRoot.ResumeLayout(false);
            this.tlpCard.ResumeLayout(false);
            this.tlpUser.ResumeLayout(false);
            this.tlpUser.PerformLayout();
            this.tlpPass.ResumeLayout(false);
            this.tlpPass.PerformLayout();
            this.flpLang.ResumeLayout(false);
            this.flpLang.PerformLayout();
            this.tlpButtons.ResumeLayout(false);
            this.ResumeLayout(false);

        }

        #endregion

        private System.Windows.Forms.TableLayoutPanel tlpRoot;
        private System.Windows.Forms.TableLayoutPanel tlpCard;
        private System.Windows.Forms.Label lblTitle;
        private System.Windows.Forms.TableLayoutPanel tlpUser;
        private System.Windows.Forms.Label lblUsername;
        private System.Windows.Forms.TextBox txtUsername;
        private System.Windows.Forms.TableLayoutPanel tlpPass;
        private System.Windows.Forms.Label lblPassword;
        private System.Windows.Forms.TextBox txtPassword;
        private System.Windows.Forms.FlowLayoutPanel flpLang;
        private System.Windows.Forms.CheckBox chkEN;
        private System.Windows.Forms.CheckBox chkVI;
        private System.Windows.Forms.TableLayoutPanel tlpButtons;
        private System.Windows.Forms.Button btnLogin;
        private System.Windows.Forms.Button btnExit;
        private System.Windows.Forms.Label lblStatus;
        #region Windows Form Designer generated code

        #endregion
    }
}