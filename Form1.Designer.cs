namespace evidence
{
    partial class Form1
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
            this.textBox1 = new System.Windows.Forms.TextBox();
            this.notifylabel = new System.Windows.Forms.Label();
            this.label1 = new System.Windows.Forms.Label();
            this.panel1 = new System.Windows.Forms.Panel();
            this.radioFail = new System.Windows.Forms.RadioButton();
            this.radioPass = new System.Windows.Forms.RadioButton();
            this.radioInfo = new System.Windows.Forms.RadioButton();
            this.comboScripts = new System.Windows.Forms.ComboBox();
            this.panel1.SuspendLayout();
            this.SuspendLayout();
            // 
            // textBox1
            // 
            this.textBox1.Location = new System.Drawing.Point(11, 54);
            this.textBox1.Multiline = true;
            this.textBox1.Name = "textBox1";
            this.textBox1.Size = new System.Drawing.Size(296, 36);
            this.textBox1.TabIndex = 2;
            // 
            // notifylabel
            // 
            this.notifylabel.AutoSize = true;
            this.notifylabel.Location = new System.Drawing.Point(16, 12);
            this.notifylabel.Name = "notifylabel";
            this.notifylabel.Size = new System.Drawing.Size(0, 13);
            this.notifylabel.TabIndex = 3;
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Location = new System.Drawing.Point(12, 39);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(82, 13);
            this.label1.TabIndex = 4;
            this.label1.Text = "Annotation Text";
            // 
            // panel1
            // 
            this.panel1.Controls.Add(this.radioFail);
            this.panel1.Controls.Add(this.radioPass);
            this.panel1.Controls.Add(this.radioInfo);
            this.panel1.Location = new System.Drawing.Point(12, 7);
            this.panel1.Name = "panel1";
            this.panel1.Size = new System.Drawing.Size(155, 33);
            this.panel1.TabIndex = 8;
            // 
            // radioFail
            // 
            this.radioFail.AutoSize = true;
            this.radioFail.Location = new System.Drawing.Point(107, 6);
            this.radioFail.Name = "radioFail";
            this.radioFail.Size = new System.Drawing.Size(41, 17);
            this.radioFail.TabIndex = 10;
            this.radioFail.TabStop = true;
            this.radioFail.Text = "Fail";
            this.radioFail.UseVisualStyleBackColor = true;
            // 
            // radioPass
            // 
            this.radioPass.AutoSize = true;
            this.radioPass.Location = new System.Drawing.Point(54, 6);
            this.radioPass.Name = "radioPass";
            this.radioPass.Size = new System.Drawing.Size(48, 17);
            this.radioPass.TabIndex = 9;
            this.radioPass.TabStop = true;
            this.radioPass.Text = "Pass";
            this.radioPass.UseVisualStyleBackColor = true;
            // 
            // radioInfo
            // 
            this.radioInfo.AutoSize = true;
            this.radioInfo.Checked = true;
            this.radioInfo.Location = new System.Drawing.Point(6, 6);
            this.radioInfo.Name = "radioInfo";
            this.radioInfo.Size = new System.Drawing.Size(43, 17);
            this.radioInfo.TabIndex = 8;
            this.radioInfo.TabStop = true;
            this.radioInfo.Text = "Info";
            this.radioInfo.UseVisualStyleBackColor = true;
            // 
            // comboScripts
            // 
            this.comboScripts.FormattingEnabled = true;
            this.comboScripts.Location = new System.Drawing.Point(174, 18);
            this.comboScripts.Name = "comboScripts";
            this.comboScripts.Size = new System.Drawing.Size(121, 21);
            this.comboScripts.TabIndex = 9;
            this.comboScripts.SelectedIndexChanged += new System.EventHandler(this.comboScripts_SelectedIndexChanged);
            // 
            // Form1
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(319, 91);
            this.Controls.Add(this.comboScripts);
            this.Controls.Add(this.panel1);
            this.Controls.Add(this.label1);
            this.Controls.Add(this.notifylabel);
            this.Controls.Add(this.textBox1);
            this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedToolWindow;
            this.Name = "Form1";
            this.StartPosition = System.Windows.Forms.FormStartPosition.Manual;
            this.Text = "Evidence";
            this.TopMost = true;
            this.FormClosed += new System.Windows.Forms.FormClosedEventHandler(this.Form1_FormClosed);
            this.panel1.ResumeLayout(false);
            this.panel1.PerformLayout();
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion
        private System.Windows.Forms.TextBox textBox1;
        private System.Windows.Forms.Label notifylabel;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.Panel panel1;
        private System.Windows.Forms.RadioButton radioFail;
        private System.Windows.Forms.RadioButton radioPass;
        private System.Windows.Forms.RadioButton radioInfo;
        private System.Windows.Forms.ComboBox comboScripts;
    }
}

