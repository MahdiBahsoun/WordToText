namespace WordToText
{
    partial class WordToText
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
            this.btnSelectWordFiles = new System.Windows.Forms.Button();
            this.label1 = new System.Windows.Forms.Label();
            this.btnExtract = new System.Windows.Forms.Button();
            this.SuspendLayout();
            // 
            // btnSelectWordFiles
            // 
            this.btnSelectWordFiles.Location = new System.Drawing.Point(165, 59);
            this.btnSelectWordFiles.Name = "btnSelectWordFiles";
            this.btnSelectWordFiles.Size = new System.Drawing.Size(82, 34);
            this.btnSelectWordFiles.TabIndex = 7;
            this.btnSelectWordFiles.Text = "Browse";
            this.btnSelectWordFiles.UseVisualStyleBackColor = true;
            this.btnSelectWordFiles.Click += new System.EventHandler(this.btnSelectWordFiles_Click);
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Location = new System.Drawing.Point(38, 70);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(93, 13);
            this.label1.TabIndex = 6;
            this.label1.Text = "Select Word Files:";
            // 
            // btnExtract
            // 
            this.btnExtract.Location = new System.Drawing.Point(115, 157);
            this.btnExtract.Name = "btnExtract";
            this.btnExtract.Size = new System.Drawing.Size(82, 45);
            this.btnExtract.TabIndex = 5;
            this.btnExtract.Text = "Extract Text";
            this.btnExtract.UseVisualStyleBackColor = true;
            this.btnExtract.Click += new System.EventHandler(this.btnExtract_Click);
            // 
            // WordToText
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(284, 261);
            this.Controls.Add(this.btnSelectWordFiles);
            this.Controls.Add(this.label1);
            this.Controls.Add(this.btnExtract);
            this.Name = "WordToText";
            this.Text = "WordToText";
            this.Load += new System.EventHandler(this.WordToText_Load);
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.Button btnSelectWordFiles;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.Button btnExtract;
    }
}

