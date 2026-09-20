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
        private readonly TestingWorkspaceService _testingWorkspace;
        private readonly TestSubmissionService _testSubmissionService;
        private TestSubmissionResult _testSubmissionResult;

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
        private bool _testingSwitchInProgress = false;

        public Form1()
            : this(null)
        {
        }

        internal Form1(TestSession testSession)
        {
            _testSession = testSession;
            _testingWorkspace = testSession == null
                ? null
                : new TestingWorkspaceService(testSession);
            InitializeComponent();

            _grading = new GradingService(_excel);
            _testSubmissionService = testSession == null
                ? null
                : new TestSubmissionService(testSession, _testingWorkspace, _excel, _grading);

            // Gắn event click ở constructor để khỏi quên
            btnGo.Click += btnGo_Click;
            btnPrev.Click += btnPrev_Click;
            btnNext.Click += btnNext_Click;
            btnRestart.Click += btnRestart_Click;
            btnGrade.Click += btnGrade_Click;
            btnPrevProject.Click += btnPrevProject_Click;
            btnNextProject.Click += btnNextProject_Click;
            btnSubmitTest.Click += btnSubmitTest_Click;

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

            btnPrevProject.Visible = true;
            btnNextProject.Visible = true;
            btnSubmitTest.Visible = true;

            LockTestingInteractions();
            bool vietnamese = string.Equals(_testSession.Language, "vi", StringComparison.OrdinalIgnoreCase);

            btnPrevProject.Text = vietnamese ? "Dự án trước" : "Previous Project";
            btnNextProject.Text = vietnamese ? "Dự án tiếp" : "Next Project";
            btnSubmitTest.Text = vietnamese ? "Nộp bài" : "Submit Test";
            UpdateTestingProjectDisplay();

            lblStatus.ForeColor = Color.Black;
            lblStatus.Text = vietnamese
                ? "Đang mở workbook của bài thi..."
                : "Opening the Testing workbook...";

            lblTimer.Visible = true;
            timerMain.Interval = 1000;
            _timerRunning = true;
            UpdateTestingCountdown();

            if (_testingTimeoutHandled)
                return;

            timerMain.Start();
            OpenInitialTestingProject();
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

            if (_excel.IsOpened)
            {
                try
                {
                    _excel.SaveWorkbook();
                    _excel.Close();
                }
                catch (Exception ex)
                {
                    lblStatus.Text = vietnamese
                        ? "Đã hết giờ nhưng không thể lưu workbook; workbook được giữ mở để tránh mất bài: " + ex.Message
                        : "Time expired, but the workbook could not be saved and remains open to avoid data loss: " + ex.Message;
                    return;
                }
            }

            lblStatus.Text = vietnamese
                ? "Đã hết thời gian làm bài. Workbook đã được lưu và đóng."
                : "Time expired. The workbook was saved and closed.";
        }

        private void LockTestingInteractions()
        {
            cbProjects.Enabled = false;
            btnGo.Enabled = false;
            btnPrev.Enabled = false;
            btnNext.Enabled = false;
            btnRestart.Enabled = false;
            btnGrade.Enabled = false;
            btnPrevProject.Enabled = false;
            btnNextProject.Enabled = false;
            btnSubmitTest.Enabled = false;
            tabTasks.Enabled = false;
        }

        private void OpenInitialTestingProject()
        {
            try
            {
                TestProjectState state = _testingWorkspace.PrepareProject(_testSession.CurrentProjectIndex);
                OpenTestingWorkbook(state);
                CommitTestingProjectUi(state);
            }
            catch (Exception ex)
            {
                bool vietnamese = string.Equals(_testSession.Language, "vi", StringComparison.OrdinalIgnoreCase);
                LockTestingInteractions();
                lblStatus.ForeColor = Color.Crimson;
                lblStatus.Text = vietnamese
                    ? "Không thể mở workbook bài thi: " + ex.Message
                    : "Unable to open the Testing workbook: " + ex.Message;
            }
            finally
            {
                UpdateTestingCountdown();
            }
        }

        private void OpenTestingWorkbook(TestProjectState state)
        {
            _currentAssetsDir = AssetsDeployer.DeployProjectAssets(
                state.ProjectId,
                state.Project.ProjectFolderPath);
            _excel.OpenWorkbook(state.WorkingWorkbookPath);
        }

        private void CommitTestingProjectUi(TestProjectState state)
        {
            _currentProject = state.Project;
            _workingXlsxPath = state.WorkingWorkbookPath;

            BuildTaskTabs(_currentProject);
            tabTasks.Visible = true;
            tabTasks.Enabled = true;
            btnPrev.Enabled = true;
            btnNext.Enabled = true;

            UpdateTestingProjectDisplay();
            UpdateTestingProjectNavigationButtons();
        }

        private void UpdateTestingProjectDisplay()
        {
            string projectId = _testSession.CurrentProject.Meta.ProjectId;
            bool vietnamese = string.Equals(_testSession.Language, "vi", StringComparison.OrdinalIgnoreCase);

            lblProject.Text = vietnamese
                ? "Dự án: " + projectId
                : "Project: " + projectId;
            lblProjectInfo.Text = vietnamese
                ? "Dự án " + _testSession.ProjectNumber + "/" + _testSession.TotalProjects
                : "Project " + _testSession.ProjectNumber + "/" + _testSession.TotalProjects;
            lblProjectInfo.Visible = true;
        }

        private void UpdateTestingProjectNavigationButtons()
        {
            bool canNavigate = !_testingTimeoutHandled &&
                !_testingSwitchInProgress &&
                !_testSession.IsSubmitting &&
                !_testSession.IsCompleted &&
                _excel.IsOpened;

            btnPrevProject.Enabled = canNavigate && _testSession.CurrentProjectIndex > 0;
            btnNextProject.Enabled = canNavigate &&
                _testSession.CurrentProjectIndex < _testSession.TotalProjects - 1;
            btnSubmitTest.Enabled = canNavigate &&
                _testSession.CurrentProjectIndex == _testSession.TotalProjects - 1;
        }

        private void btnSubmitTest_Click(object sender, EventArgs e)
        {
            if (_testSession == null ||
                _testSession.CurrentProjectIndex != _testSession.TotalProjects - 1 ||
                _testSession.IsSubmitting ||
                _testSession.IsCompleted)
            {
                return;
            }

            if (_testSession.IsExpiredAt(DateTime.UtcNow))
            {
                lblTimer.Text = "00:00";
                HandleTestingTimeExpired();
                return;
            }

            bool vietnamese = string.Equals(_testSession.Language, "vi", StringComparison.OrdinalIgnoreCase);
            string message = vietnamese
                ? "Bạn có chắc chắn muốn nộp bài?\r\nSau khi nộp, bạn sẽ không thể tiếp tục chỉnh sửa bài làm."
                : "Are you sure you want to submit the test?\r\nAfter submission, you cannot continue editing your work.";
            string title = vietnamese ? "Nộp bài" : "Submit Test";

            DialogResult confirmation = MessageBox.Show(
                this,
                message,
                title,
                MessageBoxButtons.YesNo,
                MessageBoxIcon.Question,
                MessageBoxDefaultButton.Button2);

            if (confirmation != DialogResult.Yes)
                return;

            if (_testSession.IsExpiredAt(DateTime.UtcNow))
            {
                lblTimer.Text = "00:00";
                HandleTestingTimeExpired();
                return;
            }

            BeginTestingSubmission(TestSubmissionReason.Manual);
        }

        private void BeginTestingSubmission(TestSubmissionReason reason)
        {
            if (_testSession == null || !_testSession.TryBeginSubmission())
                return;

            bool vietnamese = string.Equals(_testSession.Language, "vi", StringComparison.OrdinalIgnoreCase);
            _timerRunning = false;
            timerMain.Stop();
            LockTestingInteractions();
            lblStatus.ForeColor = Color.Black;
            lblStatus.Text = vietnamese ? "Đang nộp bài..." : "Submitting test...";

            TestSubmissionResult result;
            try
            {
                if (_excel.IsOpened)
                {
                    _excel.SaveWorkbook();
                    _excel.Close();
                }

                result = _testSubmissionService.Submit(reason);
                _testSubmissionResult = result;
                _testSession.MarkCompleted();
            }
            catch (Exception ex)
            {
                _testSession.AbortSubmission();
                RecoverFromTestingSubmissionFailure(ex);
                return;
            }

            string resultMessage = vietnamese
                ? "Đã hoàn thành bài kiểm tra.\r\n\r\nĐiểm: " + result.DisplayScore + " / 1000"
                : "Test completed.\r\n\r\nScore: " + result.DisplayScore + " / 1000";
            string resultTitle = vietnamese
                ? "Kết quả bài kiểm tra"
                : "Test Result";

            MessageBox.Show(
                this,
                resultMessage,
                resultTitle,
                MessageBoxButtons.OK,
                MessageBoxIcon.Information);

            Close();
        }

        private void RecoverFromTestingSubmissionFailure(Exception error)
        {
            MosTrainer.Core.Diagnostics.AppLogger.Error(
                "Form1.BeginTestingSubmission",
                "Testing submission failed before completion.",
                error);

            bool vietnamese = string.Equals(_testSession.Language, "vi", StringComparison.OrdinalIgnoreCase);
            string failureMessage = vietnamese
                ? "Không thể nộp bài do lỗi hệ thống. Bài kiểm tra chưa được đánh dấu hoàn thành."
                : "The test could not be submitted because of a system error. Your test has not been marked as completed.";

            DateTime currentUtc = DateTime.UtcNow;
            if (_testSession.IsExpiredAt(currentUtc))
            {
                _testingTimeoutHandled = true;
                _timerRunning = false;
                timerMain.Stop();
                lblTimer.Text = "00:00";
                LockTestingInteractions();
                lblStatus.ForeColor = Color.Crimson;
                lblStatus.Text = failureMessage;
                return;
            }

            try
            {
                TestProjectState currentState = _testingWorkspace.GetProjectState(
                    _testSession.CurrentProjectIndex);
                if (!_excel.IsOpened)
                    OpenTestingWorkbook(currentState);

                CommitTestingProjectUi(currentState);
                _timerRunning = true;
                UpdateTestingCountdown();
                if (!_testingTimeoutHandled)
                    timerMain.Start();
            }
            catch
            {
                LockTestingInteractions();
                _timerRunning = true;
                UpdateTestingCountdown();
                if (!_testingTimeoutHandled)
                    timerMain.Start();
            }

            lblStatus.ForeColor = Color.Crimson;
            lblStatus.Text = failureMessage;
        }

        private void btnPrevProject_Click(object sender, EventArgs e)
        {
            SwitchTestingProject(_testSession.CurrentProjectIndex - 1);
        }

        private void btnNextProject_Click(object sender, EventArgs e)
        {
            SwitchTestingProject(_testSession.CurrentProjectIndex + 1);
        }

        private void SwitchTestingProject(int targetIndex)
        {
            if (_testSession == null ||
                _testingSwitchInProgress ||
                _testingTimeoutHandled ||
                _testSession.IsSubmitting ||
                _testSession.IsCompleted)
                return;

            DateTime currentUtc = DateTime.UtcNow;
            if (_testSession.IsExpiredAt(currentUtc))
            {
                lblTimer.Text = "00:00";
                HandleTestingTimeExpired();
                return;
            }

            if (targetIndex < 0 || targetIndex >= _testSession.TotalProjects)
                return;

            int currentIndex = _testSession.CurrentProjectIndex;
            TestProjectState currentState = _testingWorkspace.GetProjectState(currentIndex);
            bool vietnamese = string.Equals(_testSession.Language, "vi", StringComparison.OrdinalIgnoreCase);

            _testingSwitchInProgress = true;
            UpdateTestingProjectNavigationButtons();

            try
            {
                TestProjectState targetState = _testingWorkspace.PrepareProject(targetIndex);

                _excel.SaveWorkbook();
                _excel.Close();

                try
                {
                    OpenTestingWorkbook(targetState);
                }
                catch (Exception targetException)
                {
                    string recoveryMessage;
                    try
                    {
                        _excel.OpenWorkbook(currentState.WorkingWorkbookPath);
                        _workingXlsxPath = currentState.WorkingWorkbookPath;
                        recoveryMessage = vietnamese
                            ? " Workbook hiện tại đã được mở lại."
                            : " The current workbook was reopened.";
                    }
                    catch (Exception recoveryException)
                    {
                        recoveryMessage = vietnamese
                            ? " Không thể mở lại workbook hiện tại: " + recoveryException.Message
                            : " The current workbook could not be reopened: " + recoveryException.Message;
                    }

                    throw new InvalidOperationException(
                        (vietnamese
                            ? "Không thể mở workbook đích."
                            : "The target workbook could not be opened.") +
                        recoveryMessage,
                        targetException);
                }

                _testSession.MoveToProject(targetIndex);
                CommitTestingProjectUi(targetState);
                lblStatus.ForeColor = Color.DarkGreen;
                lblStatus.Text = vietnamese
                    ? "Đã chuyển sang " + targetState.ProjectId + "."
                    : "Switched to " + targetState.ProjectId + ".";
            }
            catch (Exception ex)
            {
                lblStatus.ForeColor = Color.Crimson;
                lblStatus.Text = vietnamese
                    ? "Không thể chuyển dự án: " + ex.Message
                    : "Unable to switch project: " + ex.Message;
            }
            finally
            {
                _testingSwitchInProgress = false;
                UpdateTestingCountdown();

                if (!_testingTimeoutHandled)
                    UpdateTestingProjectNavigationButtons();
            }
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
            bool vietnameseTesting = _testSession != null &&
                string.Equals(_testSession.Language, "vi", StringComparison.OrdinalIgnoreCase);
            lblStatus.Text = vietnameseTesting
                ? "Nhiệm vụ " + (tabTasks.SelectedIndex + 1) + "/" + _currentProject.Tasks.Count
                : "Task " + (tabTasks.SelectedIndex + 1) + "/" + _currentProject.Tasks.Count;
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
        protected override void OnFormClosing(FormClosingEventArgs e)
        {
            if (_testSession != null && !_testSession.IsCompleted && _excel.IsOpened)
            {
                try
                {
                    _excel.SaveWorkbook();
                }
                catch (Exception ex)
                {
                    e.Cancel = true;
                    bool vietnamese = string.Equals(_testSession.Language, "vi", StringComparison.OrdinalIgnoreCase);
                    lblStatus.ForeColor = Color.Crimson;
                    lblStatus.Text = vietnamese
                        ? "Không thể đóng bài thi vì workbook chưa lưu được: " + ex.Message
                        : "The test cannot be closed because the workbook could not be saved: " + ex.Message;
                    return;
                }
            }

            base.OnFormClosing(e);
        }

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
