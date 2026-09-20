using MosTrainer.Core.Models;
using MosTrainer.Core.Services;
using MosTrainer.Excel;
using MosTrainer.Projects;
using MosTrainer.Testing;
using System;
using System.Collections.Generic;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Windows.Forms;
using MosTrainer.Services;
namespace MosTrainer
{
    public partial class Form1 : Form
    {
        // ===== Services (MVP) =====
        private readonly ProjectLoader _loader = new ProjectLoader();
        private readonly ExcelController _excel = new ExcelController();
        private readonly GradingService _grading;
        private readonly TestSession _testSession;

        // ===== State =====
        private List<ProjectPackage> _projects = new List<ProjectPackage>();
        private ProjectPackage _currentProject = null;   // C# 7.3: không dùng ?
        private string _currentAssetsDir = "";


        // MVP: tạm hardcode ngôn ngữ, sau nối LoginForm (vi/en)
     //   private string _language = "en";
        private string _language = AppSession.Language;


        private string _workingXlsxPath = "";
        private DateTime _projectStartTime;
        private bool _timerRunning = false;
        private bool _testingTimeoutHandled = false;

        public Form1()
            : this(null)
        {
        }

        internal Form1(TestSession testSession)
        {
            _testSession = testSession;
            InitializeComponent();

            _grading = new GradingService(_excel);

            // Gắn event click ở constructor để khỏi quên
            btnGo.Click += btnGo_Click;
            btnPrev.Click += btnPrev_Click;
            btnNext.Click += btnNext_Click;
            btnRestart.Click += btnRestart_Click;
            btnGrade.Click += btnGrade_Click;

            tabTasks.SelectedIndexChanged += tabTasks_SelectedIndexChanged;
            timerMain.Tick += timerMain_Tick;
        }

        // =========================
        // MAIN LOAD (Designer đã gắn vào Load: MainForm_Load)
        // =========================
        private void MainForm_Load(object sender, EventArgs e)
        {
            if (_testSession != null)
            {
                InitializeTestingUi();
                return;
            }

            // (1) Trạng thái ban đầu: chưa chọn project => ẨN tabTasks (đúng yêu cầu của bạn)
            tabTasks.Visible = false;

            // (2) Disable các nút điều hướng/chấm khi chưa có project
            SetProjectButtonsEnabled(false);

            // (3) Reset UI info
            lblProjectInfo.Text = "";
            lblStatus.ForeColor = Color.Black;
            lblStatus.Text = "Ready";
            lblTimer.Text = "00:00:00";

            // (4) Load danh sách projects vào ComboBox
            LoadProjectsToCombo();
        }

        private void InitializeTestingUi()
        {
            tabTasks.Visible = false;
            tabTasks.Enabled = false;

            cbProjects.Visible = false;
            btnGo.Visible = false;
            btnGrade.Visible = false;
            btnRestart.Visible = false;

            LockTestingInteractions();

            string projectId = _testSession.CurrentProject.Meta.ProjectId;
            bool vietnamese = string.Equals(_testSession.Language, "vi", StringComparison.OrdinalIgnoreCase);

            lblProject.Text = vietnamese
                ? "Dự án: " + projectId
                : "Project: " + projectId;
            lblProjectInfo.Text = vietnamese
                ? "Dự án " + _testSession.ProjectNumber + "/" + _testSession.TotalProjects
                : "Project " + _testSession.ProjectNumber + "/" + _testSession.TotalProjects;
            lblProjectInfo.Visible = true;

            lblStatus.ForeColor = Color.Black;
            lblStatus.Text = vietnamese
                ? "Chế độ thi - workbook sẽ được khởi tạo ở giai đoạn tiếp theo."
                : "Testing Mode - workbook setup will be added in the next phase.";

            lblTimer.Visible = true;
            timerMain.Interval = 1000;
            _timerRunning = true;
            UpdateTestingCountdown();

            if (!_testingTimeoutHandled)
                timerMain.Start();
        }

        private void UpdateTestingCountdown()
        {
            DateTime currentUtc = DateTime.UtcNow;
            TimeSpan remaining = _testSession.GetRemainingTime(currentUtc);

            if (_testSession.IsExpiredAt(currentUtc))
            {
                lblTimer.Text = "00:00";
                HandleTestingTimeExpired();
                return;
            }

            int totalSeconds = (int)Math.Ceiling(remaining.TotalSeconds);
            lblTimer.Text = string.Format("{0:00}:{1:00}",
                totalSeconds / 60,
                totalSeconds % 60);
        }

        private void HandleTestingTimeExpired()
        {
            if (_testingTimeoutHandled)
                return;

            _testingTimeoutHandled = true;
            _timerRunning = false;
            timerMain.Stop();
            LockTestingInteractions();

            bool vietnamese = string.Equals(_testSession.Language, "vi", StringComparison.OrdinalIgnoreCase);
            lblStatus.ForeColor = Color.Crimson;
            lblStatus.Text = vietnamese
                ? "Đã hết thời gian làm bài. Hệ thống sẽ tự động nộp bài."
                : "Time expired. Automatic submission will be handled by the testing submission pipeline.";
        }

        private void LockTestingInteractions()
        {
            cbProjects.Enabled = false;
            btnGo.Enabled = false;
            btnPrev.Enabled = false;
            btnNext.Enabled = false;
            btnRestart.Enabled = false;
            btnGrade.Enabled = false;
            tabTasks.Enabled = false;
        }

        // =========================
        // LOAD PROJECT LIST
        // =========================
        private void LoadProjectsToCombo()
        {
            try
            {
                string root = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "Projects");
                _projects = _loader.LoadAll(root, _language);

                cbProjects.DataSource = null;
                cbProjects.DataSource = _projects;
             //   cbProjects.DisplayMember = "Meta.ProjectId";
                cbProjects.DisplayMember = "DisplayName";
                cbProjects.ValueMember = "ProjectFolderPath";

                lblProjectInfo.Text = _projects.Count > 0
                    ? "Found " + _projects.Count + " projects"
                    : "No project found in /Projects";

                lblStatus.Text = "Ready";
            }
            catch (Exception ex)
            {
                lblProjectInfo.Text = "Load projects error";
                lblStatus.Text = ex.Message;
            }
        }

        // =========================
        // GO: Load selected project
        // =========================
        private void btnGo_Click(object sender, EventArgs e)
        {
            _currentProject = cbProjects.SelectedItem as ProjectPackage;
            if (_currentProject == null)
            {
                lblStatus.Text = "Please select a project.";
                return;
            }

            try
            {
                // 1) Build tabs + show instructions
                BuildTaskTabs(_currentProject);

                // 2) Open Excel workbook (copy starter -> working file)
                // OpenProjectExcel(_currentProject);
                // 2) Deploy assets của project ra Documents (cho Task import text, v.v.)
                _currentAssetsDir = AssetsDeployer.DeployProjectAssets(
                    _currentProject.Meta.ProjectId,
                    _currentProject.ProjectFolderPath
                );

                // 3) Open Excel workbook (copy starter -> working file)
                OpenProjectExcel(_currentProject);


                // 3) Sau khi load xong => HIỆN tabTasks (đúng yêu cầu của bạn)
                tabTasks.Visible = true;

                // 4) Enable buttons
                SetProjectButtonsEnabled(true);

                // 5) Start timer
                _projectStartTime = DateTime.Now;
                _timerRunning = true;
                timerMain.Interval = 1000;
                timerMain.Start();

                // 6) Status text
                lblStatus.ForeColor = Color.DarkGreen;
                lblStatus.Text = "Switched to Project: " + _currentProject.Meta.ProjectId;
            }
            catch (Exception ex)
            {
                lblStatus.ForeColor = Color.Crimson;
                lblStatus.Text = "Go error: " + ex.Message;
            }
        }

        // =========================
        // Build tabs: Task 1..N
        // =========================
        private void BuildTaskTabs(ProjectPackage project)
        {
            tabTasks.TabPages.Clear();

            if (project.Tasks == null || project.Tasks.Count == 0)
            {
                lblProjectInfo.Text = "Project " + project.Meta.ProjectId + " has 0 tasks.";
                return;
            }

            for (int i = 0; i < project.Tasks.Count; i++)
            {
                TaskDefinition task = project.Tasks[i];

                // Tab title: dùng TitleKey, nếu thiếu thì Task i
                string tabTitle = _loader.GetText(project, task.TitleKey);
                if (string.IsNullOrWhiteSpace(tabTitle))
                    tabTitle = "Task " + (i + 1);

                TabPage page = new TabPage(tabTitle);

                // Instruction label (fill)
                Label lbl = new Label();
                lbl.Dock = DockStyle.Fill;
                lbl.AutoSize = false;
                lbl.Padding = new Padding(10);
                lbl.Font = new Font("Segoe UI", 10F);
                lbl.Text = _loader.GetText(project, task.InstructionKey);

                page.Controls.Add(lbl);
                tabTasks.TabPages.Add(page);
            }

            tabTasks.SelectedIndex = 0;
            UpdateFooterTaskStatus();
            UpdateProjectInfo();
        }

        private void UpdateProjectInfo()
        {
            if (_currentProject == null) return;

            lblProjectInfo.Text = _currentProject.Meta.ProjectId + " | Tasks: " + _currentProject.Tasks.Count;
        }

        // =========================
        // Open Excel (copy starter)
        // =========================
        private void OpenProjectExcel(ProjectPackage project)
        {
            _excel.Close();

            string starterPath = Path.Combine(project.ProjectFolderPath, project.Meta.Starter);
            if (!File.Exists(starterPath))
                throw new FileNotFoundException("starter.xlsx not found", starterPath);

            string workDir = Path.Combine(
                Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments),
                "MosTrainer",
                "Working",
                project.Meta.ProjectId);

            Directory.CreateDirectory(workDir);

            _workingXlsxPath = Path.Combine(workDir, "work.xlsx");
            File.Copy(starterPath, _workingXlsxPath, true);

            _excel.OpenWorkbook(_workingXlsxPath);
        }

        // =========================
        // Prev / Next
        // =========================
        private void btnPrev_Click(object sender, EventArgs e)
        {
            if (!tabTasks.Visible || tabTasks.TabPages.Count == 0) return;

            int idx = tabTasks.SelectedIndex;
            if (idx > 0) tabTasks.SelectedIndex = idx - 1;
        }

        private void btnNext_Click(object sender, EventArgs e)
        {
            if (!tabTasks.Visible || tabTasks.TabPages.Count == 0) return;

            int idx = tabTasks.SelectedIndex;
            if (idx < tabTasks.TabPages.Count - 1) tabTasks.SelectedIndex = idx + 1;
        }

        private void tabTasks_SelectedIndexChanged(object sender, EventArgs e)
        {
            UpdateFooterTaskStatus();
        }

        private void UpdateFooterTaskStatus()
        {
            if (_currentProject == null || tabTasks.TabPages.Count == 0)
            {
                lblStatus.ForeColor = Color.Black;
                lblStatus.Text = "Ready";
                return;
            }

            lblStatus.ForeColor = Color.Black;
            lblStatus.Text = "Task " + (tabTasks.SelectedIndex + 1) + "/" + _currentProject.Tasks.Count;
        }

        // =========================
        // Restart Project
        // =========================
        private void btnRestart_Click(object sender, EventArgs e)
        {
            if (_currentProject == null) return;

            try
            {
                OpenProjectExcel(_currentProject);
                lblStatus.ForeColor = Color.DarkGreen;
                lblStatus.Text = "Restarted project (fresh workbook).";
            }
            catch (Exception ex)
            {
                lblStatus.ForeColor = Color.Crimson;
                lblStatus.Text = "Restart error: " + ex.Message;
            }
        }

        // =========================
        // Grade Project (MVP: chấm task đang chọn)
        // =========================
        private void btnGrade_Click(object sender, EventArgs e)
        {
            if (_currentProject == null) return;
            if (tabTasks.TabPages.Count == 0) return;

            try
            {
                int idx = tabTasks.SelectedIndex;
                TaskDefinition task = _currentProject.Tasks[idx];

                var result = _grading.CheckTask(task);
                bool pass = result.pass;
                string message = result.message;

                lblStatus.ForeColor = pass ? Color.DarkGreen : Color.Crimson;
                lblStatus.Text = pass
                    ? ("PASS - " + task.TaskId)
                    : ("FAIL - " + task.TaskId + " | " + message);
            }
            catch (Exception ex)
            {
                lblStatus.ForeColor = Color.Crimson;
                lblStatus.Text = "Grade error: " + ex.Message;
            }
        }

        // =========================
        // Timer
        // =========================
        private void timerMain_Tick(object sender, EventArgs e)
        {
            if (!_timerRunning) return;

            if (_testSession != null)
            {
                UpdateTestingCountdown();
                return;
            }

            TimeSpan elapsed = DateTime.Now - _projectStartTime;
            lblTimer.Text = string.Format("{0:00}:{1:00}:{2:00}",
                (int)elapsed.TotalHours, elapsed.Minutes, elapsed.Seconds);
        }

        // =========================
        // Enable/Disable buttons
        // =========================
        private void SetProjectButtonsEnabled(bool enabled)
        {
            btnPrev.Enabled = enabled;
            btnNext.Enabled = enabled;
            btnRestart.Enabled = enabled;
            btnGrade.Enabled = enabled;
        }

        // =========================
        // Close excel when exit
        // =========================
        protected override void OnFormClosed(FormClosedEventArgs e)
        {
            try
            {
                timerMain.Stop();
                _timerRunning = false;
                _excel.Close();
            }
            catch { }

            base.OnFormClosed(e);
        }

        // XÓA hẳn Form1_Load trống nếu Designer không dùng đến
        // (Nếu Designer đang gắn Load vào MainForm_Load thì không cần Form1_Load)
    }
}
