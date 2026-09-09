using System;
using System.Windows.Forms;

namespace MosTrainer
{
    public partial class LoginForm : Form
    {
        public LoginForm()
        {
            InitializeComponent();

            // default: English tick (tùy bạn)
            chkEN.Checked = true;
            chkVI.Checked = false;

            chkEN.CheckedChanged += Lang_CheckedChanged;
            chkVI.CheckedChanged += Lang_CheckedChanged;
          

            btnLogin.Click += BtnLogin_Click;
            btnExit.Click += (s, e) => Application.Exit();

            // Enter để login
            txtPassword.KeyDown += (s, e) =>
            {
                if (e.KeyCode == Keys.Enter) BtnLogin_Click(s, e);
            };

            lblStatus.Text = "";
        }

        private void Lang_CheckedChanged(object sender, EventArgs e)
        {
            // chỉ cho chọn 1
            if (sender == chkEN && chkEN.Checked) chkVI.Checked = false;
            if (sender == chkVI && chkVI.Checked) chkEN.Checked = false;

            // (tuỳ chọn) đổi text ngay trên form login
            if (chkVI.Checked)
            {
                lblTitle.Text = "Đăng nhập";
                lblUsername.Text = "Tài khoản:";
                lblPassword.Text = "Mật khẩu:";
                btnLogin.Text = "Đăng nhập";
            }
            else
            {
                lblTitle.Text = "Login";
                lblUsername.Text = "Username:";
                lblPassword.Text = "Password:";
                btnLogin.Text = "Login";
            }
        }

        private void BtnLogin_Click(object sender, EventArgs e)
        {
            lblStatus.Text = "";

            var u = txtUsername.Text.Trim();
            var p = txtPassword.Text;

            if (string.IsNullOrWhiteSpace(u) || string.IsNullOrWhiteSpace(p))
            {
                lblStatus.Text = chkVI.Checked
                    ? "Vui lòng nhập tài khoản và mật khẩu."
                    : "Please enter username and password.";
                return;
            }

            // MVP: hard-code để chạy ngay
            bool ok = (u == "admin" && p == "123456");

            if (!ok)
            {
                lblStatus.Text = chkVI.Checked
                    ? "Sai tài khoản hoặc mật khẩu."
                    : "Invalid username or password.";
                return;
            }
            AppSession.Language = chkVI.Checked ? "vi" : "en";
            // Login OK -> mở Form1
            this.Hide();
            var main = new Form1();
            main.FormClosed += (s2, e2) => this.Close();
            main.Show();
        }
    }
}
