﻿namespace MediNetCorpPDF_Extractor
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
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(Form1));
            this.openFileDialog1 = new System.Windows.Forms.OpenFileDialog();
            this.btn_upload_excel = new System.Windows.Forms.Button();
            this.btn_pdf_upload = new System.Windows.Forms.Button();
            this.openFileDialog2 = new System.Windows.Forms.OpenFileDialog();
            this.label1 = new System.Windows.Forms.Label();
            this.label2 = new System.Windows.Forms.Label();
            this.SuspendLayout();
            // 
            // openFileDialog1
            // 
            this.openFileDialog1.FileName = "openFileDialog1";
            // 
            // btn_upload_excel
            // 
            this.btn_upload_excel.Location = new System.Drawing.Point(60, 109);
            this.btn_upload_excel.Name = "btn_upload_excel";
            this.btn_upload_excel.Size = new System.Drawing.Size(164, 23);
            this.btn_upload_excel.TabIndex = 1;
            this.btn_upload_excel.Text = "Upload Excel File";
            this.btn_upload_excel.UseVisualStyleBackColor = true;
            this.btn_upload_excel.Click += new System.EventHandler(this.btn_upload_excel_Click);
            // 
            // btn_pdf_upload
            // 
            this.btn_pdf_upload.Location = new System.Drawing.Point(60, 43);
            this.btn_pdf_upload.Name = "btn_pdf_upload";
            this.btn_pdf_upload.Size = new System.Drawing.Size(164, 23);
            this.btn_pdf_upload.TabIndex = 0;
            this.btn_pdf_upload.Text = "Upload Pdf File";
            this.btn_pdf_upload.UseVisualStyleBackColor = true;
            this.btn_pdf_upload.Click += new System.EventHandler(this.btn_pdf_upload_Click);
            // 
            // openFileDialog2
            // 
            this.openFileDialog2.FileName = "openFileDialog1";
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Location = new System.Drawing.Point(12, 48);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(41, 13);
            this.label1.TabIndex = 2;
            this.label1.Text = "Step 1.";
            // 
            // label2
            // 
            this.label2.AutoSize = true;
            this.label2.Location = new System.Drawing.Point(12, 114);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(41, 13);
            this.label2.TabIndex = 3;
            this.label2.Text = "Step 2.";
            // 
            // Form1
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(284, 262);
            this.Controls.Add(this.label2);
            this.Controls.Add(this.label1);
            this.Controls.Add(this.btn_pdf_upload);
            this.Controls.Add(this.btn_upload_excel);
            this.Icon = ((System.Drawing.Icon)(resources.GetObject("$this.Icon")));
            this.Name = "Form1";
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
            this.Text = "MediNetCorp PDF Extractor";
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.OpenFileDialog openFileDialog1;
        private System.Windows.Forms.Button btn_upload_excel;
        private System.Windows.Forms.Button btn_pdf_upload;
        private System.Windows.Forms.OpenFileDialog openFileDialog2;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.Label label2;
    }
}

