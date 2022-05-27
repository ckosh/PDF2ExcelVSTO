
namespace PDF2ExcelVsto
{
    partial class FormOptions
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
            this.buttonBatchMode = new System.Windows.Forms.Button();
            this.button1 = new System.Windows.Forms.Button();
            this.textBoxDelay = new System.Windows.Forms.TextBox();
            this.label1 = new System.Windows.Forms.Label();
            this.textBoxTempFolder = new System.Windows.Forms.TextBox();
            this.label2 = new System.Windows.Forms.Label();
            this.checkBoxDebugMode = new System.Windows.Forms.CheckBox();
            this.labelsent = new System.Windows.Forms.Label();
            this.textBoxObsticalMinutes = new System.Windows.Forms.TextBox();
            this.label3 = new System.Windows.Forms.Label();
            this.label4 = new System.Windows.Forms.Label();
            this.label5 = new System.Windows.Forms.Label();
            this.labelowners = new System.Windows.Forms.Label();
            this.label6 = new System.Windows.Forms.Label();
            this.SuspendLayout();
            // 
            // buttonBatchMode
            // 
            this.buttonBatchMode.Location = new System.Drawing.Point(16, 12);
            this.buttonBatchMode.Name = "buttonBatchMode";
            this.buttonBatchMode.Size = new System.Drawing.Size(133, 45);
            this.buttonBatchMode.TabIndex = 0;
            this.buttonBatchMode.Text = "Batch";
            this.buttonBatchMode.UseVisualStyleBackColor = true;
            this.buttonBatchMode.Click += new System.EventHandler(this.buttonBatchMode_Click);
            // 
            // button1
            // 
            this.button1.Location = new System.Drawing.Point(265, 211);
            this.button1.Name = "button1";
            this.button1.Size = new System.Drawing.Size(94, 23);
            this.button1.TabIndex = 1;
            this.button1.Text = "בדיקה";
            this.button1.UseVisualStyleBackColor = true;
            this.button1.Click += new System.EventHandler(this.button1_Click);
            // 
            // textBoxDelay
            // 
            this.textBoxDelay.Location = new System.Drawing.Point(17, 117);
            this.textBoxDelay.Name = "textBoxDelay";
            this.textBoxDelay.Size = new System.Drawing.Size(133, 20);
            this.textBoxDelay.TabIndex = 2;
            this.textBoxDelay.Text = "60";
            this.textBoxDelay.TextChanged += new System.EventHandler(this.textBoxDelay_TextChanged);
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Location = new System.Drawing.Point(155, 120);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(55, 13);
            this.label1.TabIndex = 3;
            this.label1.Text = "delay sec.";
            // 
            // textBoxTempFolder
            // 
            this.textBoxTempFolder.BackColor = System.Drawing.SystemColors.InactiveCaption;
            this.textBoxTempFolder.Location = new System.Drawing.Point(16, 143);
            this.textBoxTempFolder.Name = "textBoxTempFolder";
            this.textBoxTempFolder.Size = new System.Drawing.Size(133, 20);
            this.textBoxTempFolder.TabIndex = 4;
            this.textBoxTempFolder.Text = "C:\\temp";
            // 
            // label2
            // 
            this.label2.AutoSize = true;
            this.label2.Location = new System.Drawing.Point(155, 146);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(86, 13);
            this.label2.TabIndex = 5;
            this.label2.Text = "Temporary folder";
            // 
            // checkBoxDebugMode
            // 
            this.checkBoxDebugMode.AutoSize = true;
            this.checkBoxDebugMode.Location = new System.Drawing.Point(205, 26);
            this.checkBoxDebugMode.Name = "checkBoxDebugMode";
            this.checkBoxDebugMode.Size = new System.Drawing.Size(85, 17);
            this.checkBoxDebugMode.TabIndex = 6;
            this.checkBoxDebugMode.Text = "debug mode";
            this.checkBoxDebugMode.UseVisualStyleBackColor = true;
            this.checkBoxDebugMode.UseWaitCursor = true;
            // 
            // labelsent
            // 
            this.labelsent.AutoSize = true;
            this.labelsent.Location = new System.Drawing.Point(15, 69);
            this.labelsent.Name = "labelsent";
            this.labelsent.Size = new System.Drawing.Size(13, 13);
            this.labelsent.TabIndex = 7;
            this.labelsent.Text = "0";
            // 
            // textBoxObsticalMinutes
            // 
            this.textBoxObsticalMinutes.Location = new System.Drawing.Point(18, 171);
            this.textBoxObsticalMinutes.Name = "textBoxObsticalMinutes";
            this.textBoxObsticalMinutes.Size = new System.Drawing.Size(130, 20);
            this.textBoxObsticalMinutes.TabIndex = 8;
            this.textBoxObsticalMinutes.Text = "30";
            // 
            // label3
            // 
            this.label3.AutoSize = true;
            this.label3.Location = new System.Drawing.Point(155, 174);
            this.label3.Name = "label3";
            this.label3.Size = new System.Drawing.Size(65, 13);
            this.label3.TabIndex = 9;
            this.label3.Text = "Obstacl time";
            // 
            // label4
            // 
            this.label4.AutoSize = true;
            this.label4.Location = new System.Drawing.Point(155, 120);
            this.label4.Name = "label4";
            this.label4.Size = new System.Drawing.Size(55, 13);
            this.label4.TabIndex = 10;
            this.label4.Text = "delay sec.";
            // 
            // label5
            // 
            this.label5.AutoSize = true;
            this.label5.Location = new System.Drawing.Point(50, 69);
            this.label5.Name = "label5";
            this.label5.Size = new System.Drawing.Size(75, 13);
            this.label5.TabIndex = 11;
            this.label5.Text = "number of files";
            this.label5.Click += new System.EventHandler(this.label5_Click);
            // 
            // labelowners
            // 
            this.labelowners.AutoSize = true;
            this.labelowners.Location = new System.Drawing.Point(16, 91);
            this.labelowners.Name = "labelowners";
            this.labelowners.Size = new System.Drawing.Size(13, 13);
            this.labelowners.TabIndex = 12;
            this.labelowners.Text = "0";
            // 
            // label6
            // 
            this.label6.AutoSize = true;
            this.label6.Location = new System.Drawing.Point(51, 91);
            this.label6.Name = "label6";
            this.label6.Size = new System.Drawing.Size(91, 13);
            this.label6.TabIndex = 13;
            this.label6.Text = "number of owners";
            // 
            // FormOptions
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(371, 248);
            this.Controls.Add(this.label6);
            this.Controls.Add(this.labelowners);
            this.Controls.Add(this.label5);
            this.Controls.Add(this.label4);
            this.Controls.Add(this.label3);
            this.Controls.Add(this.textBoxObsticalMinutes);
            this.Controls.Add(this.labelsent);
            this.Controls.Add(this.checkBoxDebugMode);
            this.Controls.Add(this.label2);
            this.Controls.Add(this.textBoxTempFolder);
            this.Controls.Add(this.label1);
            this.Controls.Add(this.textBoxDelay);
            this.Controls.Add(this.button1);
            this.Controls.Add(this.buttonBatchMode);
            this.Name = "FormOptions";
            this.Text = "batch options";
            this.Load += new System.EventHandler(this.FormOptions_Load);
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.Button buttonBatchMode;
        private System.Windows.Forms.Button button1;
        private System.Windows.Forms.TextBox textBoxDelay;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.TextBox textBoxTempFolder;
        private System.Windows.Forms.Label label2;
        private System.Windows.Forms.CheckBox checkBoxDebugMode;
        private System.Windows.Forms.Label labelsent;
        private System.Windows.Forms.TextBox textBoxObsticalMinutes;
        private System.Windows.Forms.Label label3;
        private System.Windows.Forms.Label label4;
        private System.Windows.Forms.Label label5;
        private System.Windows.Forms.Label labelowners;
        private System.Windows.Forms.Label label6;
    }
}