namespace GoogleForms
{
    partial class MainForm
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
            this.btnBrowse = new System.Windows.Forms.Button();
            this.btnXLSBrowse = new System.Windows.Forms.Button();
            this.label1 = new System.Windows.Forms.Label();
            this.label2 = new System.Windows.Forms.Label();
            this.tbCsvInputFile = new System.Windows.Forms.TextBox();
            this.tbXLSTemplate = new System.Windows.Forms.TextBox();
            this.pbProgress = new System.Windows.Forms.ProgressBar();
            this.tbXLSOutputFolder = new System.Windows.Forms.TextBox();
            this.btnOutputBrowse = new System.Windows.Forms.Button();
            this.label3 = new System.Windows.Forms.Label();
            this.btnGenerate = new System.Windows.Forms.Button();
            this.SuspendLayout();
            // 
            // btnBrowse
            // 
            this.btnBrowse.Location = new System.Drawing.Point(466, 22);
            this.btnBrowse.Name = "btnBrowse";
            this.btnBrowse.Size = new System.Drawing.Size(169, 23);
            this.btnBrowse.TabIndex = 0;
            this.btnBrowse.Text = "&Browse (set CSV input file)";
            this.btnBrowse.UseVisualStyleBackColor = true;
            this.btnBrowse.Click += new System.EventHandler(this.btnBrowse_Click);
            // 
            // btnXLSBrowse
            // 
            this.btnXLSBrowse.Location = new System.Drawing.Point(466, 60);
            this.btnXLSBrowse.Name = "btnXLSBrowse";
            this.btnXLSBrowse.Size = new System.Drawing.Size(169, 23);
            this.btnXLSBrowse.TabIndex = 1;
            this.btnXLSBrowse.Text = "&Browse (set XLS template file)";
            this.btnXLSBrowse.UseVisualStyleBackColor = true;
            this.btnXLSBrowse.Click += new System.EventHandler(this.btnXLSBrowse_Click);
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Location = new System.Drawing.Point(12, 9);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(120, 13);
            this.label1.TabIndex = 2;
            this.label1.Text = "CSV file with responses:";
            // 
            // label2
            // 
            this.label2.AutoSize = true;
            this.label2.Location = new System.Drawing.Point(12, 48);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(70, 13);
            this.label2.TabIndex = 3;
            this.label2.Text = "XLS template";
            // 
            // tbCsvInputFile
            // 
            this.tbCsvInputFile.Enabled = false;
            this.tbCsvInputFile.Location = new System.Drawing.Point(12, 25);
            this.tbCsvInputFile.Name = "tbCsvInputFile";
            this.tbCsvInputFile.ReadOnly = true;
            this.tbCsvInputFile.Size = new System.Drawing.Size(448, 20);
            this.tbCsvInputFile.TabIndex = 4;
            // 
            // tbXLSTemplate
            // 
            this.tbXLSTemplate.Enabled = false;
            this.tbXLSTemplate.Location = new System.Drawing.Point(12, 63);
            this.tbXLSTemplate.Name = "tbXLSTemplate";
            this.tbXLSTemplate.ReadOnly = true;
            this.tbXLSTemplate.Size = new System.Drawing.Size(448, 20);
            this.tbXLSTemplate.TabIndex = 5;
            // 
            // pbProgress
            // 
            this.pbProgress.Enabled = false;
            this.pbProgress.Location = new System.Drawing.Point(12, 128);
            this.pbProgress.Name = "pbProgress";
            this.pbProgress.Size = new System.Drawing.Size(448, 23);
            this.pbProgress.TabIndex = 6;
            // 
            // tbXLSOutputFolder
            // 
            this.tbXLSOutputFolder.Enabled = false;
            this.tbXLSOutputFolder.Location = new System.Drawing.Point(12, 102);
            this.tbXLSOutputFolder.Name = "tbXLSOutputFolder";
            this.tbXLSOutputFolder.ReadOnly = true;
            this.tbXLSOutputFolder.Size = new System.Drawing.Size(448, 20);
            this.tbXLSOutputFolder.TabIndex = 7;
            // 
            // btnOutputBrowse
            // 
            this.btnOutputBrowse.Location = new System.Drawing.Point(466, 99);
            this.btnOutputBrowse.Name = "btnOutputBrowse";
            this.btnOutputBrowse.Size = new System.Drawing.Size(169, 23);
            this.btnOutputBrowse.TabIndex = 8;
            this.btnOutputBrowse.Text = "&Browse (set XLS output folder)";
            this.btnOutputBrowse.UseVisualStyleBackColor = true;
            this.btnOutputBrowse.Click += new System.EventHandler(this.btnOutputBrowse_Click);
            // 
            // label3
            // 
            this.label3.AutoSize = true;
            this.label3.Location = new System.Drawing.Point(12, 86);
            this.label3.Name = "label3";
            this.label3.Size = new System.Drawing.Size(92, 13);
            this.label3.TabIndex = 9;
            this.label3.Text = "XLS output folder:";
            // 
            // btnGenerate
            // 
            this.btnGenerate.Location = new System.Drawing.Point(466, 128);
            this.btnGenerate.Name = "btnGenerate";
            this.btnGenerate.Size = new System.Drawing.Size(169, 23);
            this.btnGenerate.TabIndex = 10;
            this.btnGenerate.Text = "Generate output XLS files";
            this.btnGenerate.UseVisualStyleBackColor = true;
            this.btnGenerate.Click += new System.EventHandler(this.btnGenerate_Click);
            // 
            // MainForm
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(643, 167);
            this.Controls.Add(this.btnGenerate);
            this.Controls.Add(this.label3);
            this.Controls.Add(this.btnOutputBrowse);
            this.Controls.Add(this.tbXLSOutputFolder);
            this.Controls.Add(this.pbProgress);
            this.Controls.Add(this.tbXLSTemplate);
            this.Controls.Add(this.tbCsvInputFile);
            this.Controls.Add(this.label2);
            this.Controls.Add(this.label1);
            this.Controls.Add(this.btnXLSBrowse);
            this.Controls.Add(this.btnBrowse);
            this.MaximizeBox = false;
            this.MinimizeBox = false;
            this.Name = "MainForm";
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
            this.Text = "Google Forms v1.0.0.0";
            this.FormClosed += new System.Windows.Forms.FormClosedEventHandler(this.MainForm_FormClosed);
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.Button btnBrowse;
        private System.Windows.Forms.Button btnXLSBrowse;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.Label label2;
        private System.Windows.Forms.TextBox tbCsvInputFile;
        private System.Windows.Forms.TextBox tbXLSTemplate;
        private System.Windows.Forms.ProgressBar pbProgress;
        private System.Windows.Forms.TextBox tbXLSOutputFolder;
        private System.Windows.Forms.Button btnOutputBrowse;
        private System.Windows.Forms.Label label3;
        private System.Windows.Forms.Button btnGenerate;
    }
}

