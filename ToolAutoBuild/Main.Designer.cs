namespace AutoRunBuild
{
    partial class Main
    {
        /// <summary>
        /// Required designer variable.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        /// <summary>
        /// Clean up any resources being used.
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
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(Main));
            this.label1 = new System.Windows.Forms.Label();
            this.btnOpenFolderSVN = new System.Windows.Forms.Button();
            this.txtPathSVN = new System.Windows.Forms.TextBox();
            this.groupBox1 = new System.Windows.Forms.GroupBox();
            this.btnOpenFolderSrc = new System.Windows.Forms.Button();
            this.label3 = new System.Windows.Forms.Label();
            this.txtPathSrc = new System.Windows.Forms.TextBox();
            this.btnRun = new System.Windows.Forms.Button();
            this.txtLogs = new System.Windows.Forms.RichTextBox();
            this.label2 = new System.Windows.Forms.Label();
            this.backgroundWorker1 = new System.ComponentModel.BackgroundWorker();
            this.progressBar = new System.Windows.Forms.ProgressBar();
            this.groupBox1.SuspendLayout();
            this.SuspendLayout();
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Font = new System.Drawing.Font("Century Gothic", 9.75F);
            this.label1.Location = new System.Drawing.Point(3, 20);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(96, 17);
            this.label1.TabIndex = 5;
            this.label1.Text = "Path Files SVN";
            // 
            // btnOpenFolderSVN
            // 
            this.btnOpenFolderSVN.Image = ((System.Drawing.Image)(resources.GetObject("btnOpenFolderSVN.Image")));
            this.btnOpenFolderSVN.Location = new System.Drawing.Point(344, 15);
            this.btnOpenFolderSVN.Name = "btnOpenFolderSVN";
            this.btnOpenFolderSVN.Size = new System.Drawing.Size(26, 24);
            this.btnOpenFolderSVN.TabIndex = 4;
            this.btnOpenFolderSVN.UseVisualStyleBackColor = true;
            this.btnOpenFolderSVN.Click += new System.EventHandler(this.btnOpenFolderSVN_Click);
            // 
            // txtPathSVN
            // 
            this.txtPathSVN.Enabled = false;
            this.txtPathSVN.Font = new System.Drawing.Font("Century Gothic", 10F);
            this.txtPathSVN.Location = new System.Drawing.Point(99, 15);
            this.txtPathSVN.Name = "txtPathSVN";
            this.txtPathSVN.ReadOnly = true;
            this.txtPathSVN.Size = new System.Drawing.Size(240, 24);
            this.txtPathSVN.TabIndex = 3;
            this.txtPathSVN.TabStop = false;
            // 
            // groupBox1
            // 
            this.groupBox1.Controls.Add(this.progressBar);
            this.groupBox1.Controls.Add(this.btnOpenFolderSrc);
            this.groupBox1.Controls.Add(this.label3);
            this.groupBox1.Controls.Add(this.txtPathSrc);
            this.groupBox1.Controls.Add(this.btnRun);
            this.groupBox1.Controls.Add(this.txtLogs);
            this.groupBox1.Controls.Add(this.label2);
            this.groupBox1.Controls.Add(this.btnOpenFolderSVN);
            this.groupBox1.Controls.Add(this.label1);
            this.groupBox1.Controls.Add(this.txtPathSVN);
            this.groupBox1.Font = new System.Drawing.Font("Century Gothic", 9.75F);
            this.groupBox1.Location = new System.Drawing.Point(9, 1);
            this.groupBox1.Name = "groupBox1";
            this.groupBox1.Size = new System.Drawing.Size(376, 286);
            this.groupBox1.TabIndex = 6;
            this.groupBox1.TabStop = false;
            this.groupBox1.Text = "Setting";
            // 
            // btnOpenFolderSrc
            // 
            this.btnOpenFolderSrc.Image = ((System.Drawing.Image)(resources.GetObject("btnOpenFolderSrc.Image")));
            this.btnOpenFolderSrc.Location = new System.Drawing.Point(344, 45);
            this.btnOpenFolderSrc.Name = "btnOpenFolderSrc";
            this.btnOpenFolderSrc.Size = new System.Drawing.Size(26, 24);
            this.btnOpenFolderSrc.TabIndex = 26;
            this.btnOpenFolderSrc.UseVisualStyleBackColor = true;
            this.btnOpenFolderSrc.Click += new System.EventHandler(this.btnOpenFolderSrc_Click);
            // 
            // label3
            // 
            this.label3.AutoSize = true;
            this.label3.Font = new System.Drawing.Font("Century Gothic", 9.75F);
            this.label3.Location = new System.Drawing.Point(3, 50);
            this.label3.Name = "label3";
            this.label3.Size = new System.Drawing.Size(89, 17);
            this.label3.TabIndex = 27;
            this.label3.Text = "Path Files Src";
            // 
            // txtPathSrc
            // 
            this.txtPathSrc.Enabled = false;
            this.txtPathSrc.Font = new System.Drawing.Font("Century Gothic", 10F);
            this.txtPathSrc.Location = new System.Drawing.Point(99, 45);
            this.txtPathSrc.Name = "txtPathSrc";
            this.txtPathSrc.ReadOnly = true;
            this.txtPathSrc.Size = new System.Drawing.Size(240, 24);
            this.txtPathSrc.TabIndex = 25;
            // 
            // btnRun
            // 
            this.btnRun.Image = ((System.Drawing.Image)(resources.GetObject("btnRun.Image")));
            this.btnRun.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft;
            this.btnRun.Location = new System.Drawing.Point(292, 255);
            this.btnRun.Name = "btnRun";
            this.btnRun.Size = new System.Drawing.Size(78, 24);
            this.btnRun.TabIndex = 24;
            this.btnRun.Text = "    Run";
            this.btnRun.UseVisualStyleBackColor = true;
            this.btnRun.Click += new System.EventHandler(this.btnRun_Click);
            // 
            // txtLogs
            // 
            this.txtLogs.Font = new System.Drawing.Font("Century Gothic", 9.5F);
            this.txtLogs.Location = new System.Drawing.Point(6, 100);
            this.txtLogs.Name = "txtLogs";
            this.txtLogs.ReadOnly = true;
            this.txtLogs.Size = new System.Drawing.Size(364, 151);
            this.txtLogs.TabIndex = 23;
            this.txtLogs.Text = "";
            // 
            // label2
            // 
            this.label2.AutoSize = true;
            this.label2.Font = new System.Drawing.Font("Century Gothic", 9.75F);
            this.label2.Location = new System.Drawing.Point(3, 80);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(37, 17);
            this.label2.TabIndex = 6;
            this.label2.Text = "Logs";
            // 
            // progressBar
            // 
            this.progressBar.Location = new System.Drawing.Point(6, 256);
            this.progressBar.Name = "progressBar";
            this.progressBar.Size = new System.Drawing.Size(281, 22);
            this.progressBar.TabIndex = 28;
            // 
            // Main
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 12F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(394, 295);
            this.Controls.Add(this.groupBox1);
            this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedSingle;
            this.Icon = ((System.Drawing.Icon)(resources.GetObject("$this.Icon")));
            this.Name = "Main";
            this.Text = "Auto Run Build Srouce";
            this.Load += new System.EventHandler(this.Main_Load);
            this.groupBox1.ResumeLayout(false);
            this.groupBox1.PerformLayout();
            this.ResumeLayout(false);

        }

        #endregion

        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.Button btnOpenFolderSVN;
        private System.Windows.Forms.TextBox txtPathSVN;
        private System.Windows.Forms.GroupBox groupBox1;
        private System.ComponentModel.BackgroundWorker backgroundWorker1;
        private System.Windows.Forms.RichTextBox txtLogs;
        private System.Windows.Forms.Label label2;
        private System.Windows.Forms.Button btnRun;
        private System.Windows.Forms.Button btnOpenFolderSrc;
        private System.Windows.Forms.Label label3;
        private System.Windows.Forms.TextBox txtPathSrc;
        private System.Windows.Forms.ProgressBar progressBar;
    }
}

