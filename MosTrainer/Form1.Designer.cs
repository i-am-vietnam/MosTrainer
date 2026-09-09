namespace MosTrainer
{
    partial class Form1
    {
        /// <summary>
        /// Required designer variable.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        /// <summary>
        /// Clean up any resources being being disposed; otherwise, false.
        /// </summary>
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
            this.components = new System.ComponentModel.Container();
            this.splitMain = new System.Windows.Forms.SplitContainer();
            this.tlpBottom = new System.Windows.Forms.TableLayoutPanel();
            this.pnlTopBar = new System.Windows.Forms.Panel();
            this.lblProject = new System.Windows.Forms.Label();
            this.cbProjects = new System.Windows.Forms.ComboBox();
            this.btnGo = new System.Windows.Forms.Button();
            this.lblProjectInfo = new System.Windows.Forms.Label();
            this.tabTasks = new System.Windows.Forms.TabControl();
            this.pnlFooter = new System.Windows.Forms.Panel();
            this.tblFooterButtons = new System.Windows.Forms.TableLayoutPanel();
            this.btnPrev = new System.Windows.Forms.Button();
            this.btnNext = new System.Windows.Forms.Button();
            this.btnRestart = new System.Windows.Forms.Button();
            this.btnGrade = new System.Windows.Forms.Button();
            this.lblStatus = new System.Windows.Forms.Label();
            this.lblTimer = new System.Windows.Forms.Label();
            this.timerMain = new System.Windows.Forms.Timer(this.components);
            ((System.ComponentModel.ISupportInitialize)(this.splitMain)).BeginInit();
            this.splitMain.Panel2.SuspendLayout();
            this.splitMain.SuspendLayout();
            this.tlpBottom.SuspendLayout();
            this.pnlTopBar.SuspendLayout();
            this.pnlFooter.SuspendLayout();
            this.tblFooterButtons.SuspendLayout();
            this.SuspendLayout();
            // 
            // splitMain
            // 
            this.splitMain.Dock = System.Windows.Forms.DockStyle.Fill;
            this.splitMain.FixedPanel = System.Windows.Forms.FixedPanel.Panel1;
            this.splitMain.Location = new System.Drawing.Point(0, 0);
            this.splitMain.Name = "splitMain";
            this.splitMain.Orientation = System.Windows.Forms.Orientation.Horizontal;
            this.splitMain.Panel1Collapsed = true;
            this.splitMain.Panel1MinSize = 0;
            // 
            // splitMain.Panel2
            // 
            this.splitMain.Panel2.Controls.Add(this.tlpBottom);
            this.splitMain.Size = new System.Drawing.Size(1912, 275);
            this.splitMain.SplitterDistance = 25;
            this.splitMain.TabIndex = 0;
            // 
            // tlpBottom
            // 
            this.tlpBottom.ColumnCount = 1;
            this.tlpBottom.ColumnStyles.Add(new System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Percent, 100F));
            this.tlpBottom.Controls.Add(this.pnlTopBar, 0, 0);
            this.tlpBottom.Controls.Add(this.tabTasks, 0, 1);
            this.tlpBottom.Controls.Add(this.pnlFooter, 0, 2);
            this.tlpBottom.Dock = System.Windows.Forms.DockStyle.Fill;
            this.tlpBottom.Location = new System.Drawing.Point(0, 0);
            this.tlpBottom.Name = "tlpBottom";
            this.tlpBottom.RowCount = 3;
            this.tlpBottom.RowStyles.Add(new System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Absolute, 58F));
            this.tlpBottom.RowStyles.Add(new System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Percent, 100F));
            this.tlpBottom.RowStyles.Add(new System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Absolute, 56F));
            this.tlpBottom.Size = new System.Drawing.Size(1912, 275);
            this.tlpBottom.TabIndex = 0;
            // 
            // pnlTopBar
            // 
            this.pnlTopBar.Controls.Add(this.lblProject);
            this.pnlTopBar.Controls.Add(this.cbProjects);
            this.pnlTopBar.Controls.Add(this.btnGo);
            this.pnlTopBar.Controls.Add(this.lblProjectInfo);
            this.pnlTopBar.Dock = System.Windows.Forms.DockStyle.Fill;
            this.pnlTopBar.Location = new System.Drawing.Point(3, 3);
            this.pnlTopBar.Name = "pnlTopBar";
            this.pnlTopBar.Padding = new System.Windows.Forms.Padding(12);
            this.pnlTopBar.Size = new System.Drawing.Size(1906, 52);
            this.pnlTopBar.TabIndex = 0;
            // 
            // lblProject
            // 
            this.lblProject.AutoSize = true;
            this.lblProject.Font = new System.Drawing.Font("Segoe UI", 9F);
            this.lblProject.Location = new System.Drawing.Point(24, 14);
            this.lblProject.Name = "lblProject";
            this.lblProject.Size = new System.Drawing.Size(58, 20);
            this.lblProject.TabIndex = 0;
            this.lblProject.Text = "Project:";
            // 
            // cbProjects
            // 
            this.cbProjects.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
            this.cbProjects.FormattingEnabled = true;
            this.cbProjects.Location = new System.Drawing.Point(88, 11);
            this.cbProjects.Name = "cbProjects";
            this.cbProjects.Size = new System.Drawing.Size(260, 28);
            this.cbProjects.TabIndex = 1;
            // 
            // btnGo
            // 
            this.btnGo.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(44)))), ((int)(((byte)(123)))), ((int)(((byte)(229)))));
            this.btnGo.Font = new System.Drawing.Font("Segoe UI Semibold", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.btnGo.ForeColor = System.Drawing.Color.White;
            this.btnGo.Location = new System.Drawing.Point(354, 11);
            this.btnGo.Name = "btnGo";
            this.btnGo.Size = new System.Drawing.Size(60, 28);
            this.btnGo.TabIndex = 2;
            this.btnGo.Text = "Go";
            this.btnGo.UseVisualStyleBackColor = false;
            // 
            // lblProjectInfo
            // 
            this.lblProjectInfo.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
            this.lblProjectInfo.AutoSize = true;
            this.lblProjectInfo.Location = new System.Drawing.Point(1785, 15);
            this.lblProjectInfo.Name = "lblProjectInfo";
            this.lblProjectInfo.Size = new System.Drawing.Size(97, 20);
            this.lblProjectInfo.TabIndex = 3;
            this.lblProjectInfo.Text = "Project 0 of 0";
            this.lblProjectInfo.Visible = false;
            // 
            // tabTasks
            // 
            this.tabTasks.Dock = System.Windows.Forms.DockStyle.Fill;
            this.tabTasks.Font = new System.Drawing.Font("Segoe UI", 10F);
            this.tabTasks.Location = new System.Drawing.Point(3, 61);
            this.tabTasks.Name = "tabTasks";
            this.tabTasks.SelectedIndex = 0;
            this.tabTasks.Size = new System.Drawing.Size(1906, 155);
            this.tabTasks.TabIndex = 5;
            this.tabTasks.Visible = false;
            // 
            // pnlFooter
            // 
            this.pnlFooter.Controls.Add(this.tblFooterButtons);
            this.pnlFooter.Controls.Add(this.lblTimer);
            this.pnlFooter.Dock = System.Windows.Forms.DockStyle.Fill;
            this.pnlFooter.Location = new System.Drawing.Point(3, 222);
            this.pnlFooter.Name = "pnlFooter";
            this.pnlFooter.Padding = new System.Windows.Forms.Padding(8);
            this.pnlFooter.Size = new System.Drawing.Size(1906, 50);
            this.pnlFooter.TabIndex = 6;
            // 
            // tblFooterButtons
            // 
            this.tblFooterButtons.ColumnCount = 5;
            this.tblFooterButtons.ColumnStyles.Add(new System.Windows.Forms.ColumnStyle());
            this.tblFooterButtons.ColumnStyles.Add(new System.Windows.Forms.ColumnStyle());
            this.tblFooterButtons.ColumnStyles.Add(new System.Windows.Forms.ColumnStyle());
            this.tblFooterButtons.ColumnStyles.Add(new System.Windows.Forms.ColumnStyle());
            this.tblFooterButtons.ColumnStyles.Add(new System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Percent, 100F));
            this.tblFooterButtons.Controls.Add(this.btnPrev, 0, 0);
            this.tblFooterButtons.Controls.Add(this.btnNext, 1, 0);
            this.tblFooterButtons.Controls.Add(this.btnRestart, 2, 0);
            this.tblFooterButtons.Controls.Add(this.btnGrade, 3, 0);
            this.tblFooterButtons.Controls.Add(this.lblStatus, 4, 0);
            this.tblFooterButtons.Dock = System.Windows.Forms.DockStyle.Fill;
            this.tblFooterButtons.Location = new System.Drawing.Point(8, 8);
            this.tblFooterButtons.Name = "tblFooterButtons";
            this.tblFooterButtons.RowCount = 1;
            this.tblFooterButtons.RowStyles.Add(new System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Percent, 100F));
            this.tblFooterButtons.Size = new System.Drawing.Size(1890, 34);
            this.tblFooterButtons.TabIndex = 0;
            // 
            // btnPrev
            // 
            this.btnPrev.Location = new System.Drawing.Point(3, 3);
            this.btnPrev.Name = "btnPrev";
            this.btnPrev.Size = new System.Drawing.Size(60, 28);
            this.btnPrev.TabIndex = 0;
            this.btnPrev.Text = "<<";
            this.btnPrev.UseVisualStyleBackColor = true;
            // 
            // btnNext
            // 
            this.btnNext.Location = new System.Drawing.Point(69, 3);
            this.btnNext.Name = "btnNext";
            this.btnNext.Size = new System.Drawing.Size(60, 28);
            this.btnNext.TabIndex = 1;
            this.btnNext.Text = ">>";
            this.btnNext.UseVisualStyleBackColor = true;
            // 
            // btnRestart
            // 
            this.btnRestart.Location = new System.Drawing.Point(135, 3);
            this.btnRestart.Name = "btnRestart";
            this.btnRestart.Size = new System.Drawing.Size(120, 28);
            this.btnRestart.TabIndex = 2;
            this.btnRestart.Text = "Restart Project";
            this.btnRestart.UseVisualStyleBackColor = true;
            // 
            // btnGrade
            // 
            this.btnGrade.Location = new System.Drawing.Point(261, 3);
            this.btnGrade.Name = "btnGrade";
            this.btnGrade.Size = new System.Drawing.Size(120, 28);
            this.btnGrade.TabIndex = 3;
            this.btnGrade.Text = "Grade Project";
            this.btnGrade.UseVisualStyleBackColor = true;
            // 
            // lblStatus
            // 
            this.lblStatus.Dock = System.Windows.Forms.DockStyle.Fill;
            this.lblStatus.ForeColor = System.Drawing.Color.DarkGreen;
            this.lblStatus.Location = new System.Drawing.Point(387, 0);
            this.lblStatus.Name = "lblStatus";
            this.lblStatus.Size = new System.Drawing.Size(1500, 34);
            this.lblStatus.TabIndex = 4;
            this.lblStatus.Text = "Ready";
            this.lblStatus.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
            // 
            // lblTimer
            // 
            this.lblTimer.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
            this.lblTimer.AutoSize = true;
            this.lblTimer.Font = new System.Drawing.Font("Segoe UI", 10F, System.Drawing.FontStyle.Bold);
            this.lblTimer.ForeColor = System.Drawing.Color.Crimson;
            this.lblTimer.Location = new System.Drawing.Point(1810, 12);
            this.lblTimer.Name = "lblTimer";
            this.lblTimer.Size = new System.Drawing.Size(80, 23);
            this.lblTimer.TabIndex = 6;
            this.lblTimer.Text = "00:00:00";
            this.lblTimer.Visible = false;
            // 
            // timerMain
            // 
            this.timerMain.Interval = 1000;
            // 
            // Form1
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(8F, 20F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.BackColor = System.Drawing.Color.WhiteSmoke;
            this.ClientSize = new System.Drawing.Size(1912, 275);
            this.Controls.Add(this.splitMain);
            this.Font = new System.Drawing.Font("Segoe UI", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.Name = "Form1";
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
            this.Text = "MOS Trainer";
            this.Load += new System.EventHandler(this.MainForm_Load);
            this.splitMain.Panel2.ResumeLayout(false);
            ((System.ComponentModel.ISupportInitialize)(this.splitMain)).EndInit();
            this.splitMain.ResumeLayout(false);
            this.tlpBottom.ResumeLayout(false);
            this.pnlTopBar.ResumeLayout(false);
            this.pnlTopBar.PerformLayout();
            this.pnlFooter.ResumeLayout(false);
            this.pnlFooter.PerformLayout();
            this.tblFooterButtons.ResumeLayout(false);
            this.ResumeLayout(false);

        }

        #endregion

        private System.Windows.Forms.SplitContainer splitMain;
        private System.Windows.Forms.TableLayoutPanel tlpBottom;
        private System.Windows.Forms.Panel pnlTopBar;
        private System.Windows.Forms.Label lblProject;
        private System.Windows.Forms.ComboBox cbProjects;
        private System.Windows.Forms.Button btnGo;
        private System.Windows.Forms.Label lblProjectInfo;
        private System.Windows.Forms.TabControl tabTasks;
        private System.Windows.Forms.Panel pnlFooter;
        private System.Windows.Forms.TableLayoutPanel tblFooterButtons;
        private System.Windows.Forms.Button btnPrev;
        private System.Windows.Forms.Button btnNext;
        private System.Windows.Forms.Button btnRestart;
        private System.Windows.Forms.Button btnGrade;
        private System.Windows.Forms.Label lblStatus;
        private System.Windows.Forms.Label lblTimer;
        private System.Windows.Forms.Timer timerMain;
    }
}
