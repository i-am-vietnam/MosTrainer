using System;
using System.Windows.Forms;
using MosTrainer.Projects;
using MosTrainer.Testing;
using System.IO;

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
            radTraining.Checked = true;

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
            AppSession.Mode = radTesting.Checked ? AppMode.Testing : AppMode.Training;

            Form1 main;
            if (AppSession.Mode == AppMode.Testing)
            {
                try
                {
                    string projectsRoot = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "Projects");
                    var projects = new ProjectLoader().LoadAll(projectsRoot, AppSession.Language);
                    TestSession testSession = new TestSessionFactory().Create(projects, AppSession.Language);
                    main = new Form1(testSession);
                }
                catch (InvalidOperationException ex)
                {
                    lblStatus.Text = chkVI.Checked
                        ? "Chế độ Testing cần ít nhất 7 project hợp lệ."
                        : ex.Message;
                    return;
                }
                catch (Exception ex)
                {
                    lblStatus.Text = chkVI.Checked
                        ? "Không thể bắt đầu chế độ Testing: " + ex.Message
                        : "Unable to start Testing Mode: " + ex.Message;
                    return;
                }
            }
            else
            {
                main = new Form1();
            }

            // Login OK -> mở Form1
            this.Hide();
            bool isTesting = AppSession.Mode == AppMode.Testing;
            main.FormClosed += (s2, e2) =>
            {
                if (isTesting)
                    this.Show();
                else
                    this.Close();
            };
            main.Show();
        }
    }
}
