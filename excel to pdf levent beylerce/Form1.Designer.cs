using System.Drawing;
using System.Windows.Forms;

namespace excel_to_pdf_levent_beylerce
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
            this.buttonAddConvert = new System.Windows.Forms.Button();
            this.listBox1 = new System.Windows.Forms.ListBox();
            this.button1 = new System.Windows.Forms.Button();
            this.button2 = new System.Windows.Forms.Button();
            this.progressBar1 = new System.Windows.Forms.ProgressBar();
            this.labelLog = new System.Windows.Forms.Label();
            this.button3 = new System.Windows.Forms.Button();
            this.button4 = new System.Windows.Forms.Button();
            this.SuspendLayout();
            // 
            // buttonAddConvert
            // 
            this.buttonAddConvert.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(162)));
            this.buttonAddConvert.ForeColor = System.Drawing.Color.SeaGreen;
            this.buttonAddConvert.Location = new System.Drawing.Point(309, 131);
            this.buttonAddConvert.Name = "buttonAddConvert";
            this.buttonAddConvert.Size = new System.Drawing.Size(203, 63);
            this.buttonAddConvert.TabIndex = 3;
            this.buttonAddConvert.Text = "DOSYA EKLE (Excel  &&  Word)";
            this.buttonAddConvert.UseVisualStyleBackColor = true;
            this.buttonAddConvert.Click += new System.EventHandler(this.buttonAddConvert_Click);
            // 
            // listBox1
            // 
            this.listBox1.ForeColor = System.Drawing.Color.ForestGreen;
            this.listBox1.FormattingEnabled = true;
            this.listBox1.Location = new System.Drawing.Point(12, 21);
            this.listBox1.Name = "listBox1";
            this.listBox1.Size = new System.Drawing.Size(291, 173);
            this.listBox1.TabIndex = 4;
            // 
            // button1
            // 
            this.button1.BackgroundImage = global::excel_to_pdf_levent_beylerce.Properties.Resources.arrow_151_64;
            this.button1.BackgroundImageLayout = System.Windows.Forms.ImageLayout.Zoom;
            this.button1.FlatAppearance.BorderColor = System.Drawing.SystemColors.Control;
            this.button1.FlatAppearance.BorderSize = 0;
            this.button1.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.button1.Location = new System.Drawing.Point(309, 21);
            this.button1.Name = "button1";
            this.button1.Size = new System.Drawing.Size(30, 50);
            this.button1.TabIndex = 5;
            this.button1.UseVisualStyleBackColor = true;
            this.button1.Click += new System.EventHandler(this.button1_Click_1);
            // 
            // button2
            // 
            this.button2.BackgroundImage = global::excel_to_pdf_levent_beylerce.Properties.Resources.arrow_213_64;
            this.button2.BackgroundImageLayout = System.Windows.Forms.ImageLayout.Zoom;
            this.button2.FlatAppearance.BorderColor = System.Drawing.SystemColors.Control;
            this.button2.FlatAppearance.BorderSize = 0;
            this.button2.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.button2.Location = new System.Drawing.Point(309, 77);
            this.button2.Name = "button2";
            this.button2.Size = new System.Drawing.Size(30, 50);
            this.button2.TabIndex = 5;
            this.button2.UseVisualStyleBackColor = true;
            this.button2.Click += new System.EventHandler(this.button2_Click_1);
            // 
            // progressBar1
            // 
            this.progressBar1.Location = new System.Drawing.Point(12, 200);
            this.progressBar1.Name = "progressBar1";
            this.progressBar1.Size = new System.Drawing.Size(500, 23);
            this.progressBar1.TabIndex = 6;
            // 
            // labelLog
            // 
            this.labelLog.Location = new System.Drawing.Point(12, 226);
            this.labelLog.Name = "labelLog";
            this.labelLog.Size = new System.Drawing.Size(500, 23);
            this.labelLog.TabIndex = 7;
            // 
            // button3
            // 
            this.button3.BackColor = System.Drawing.Color.LightCoral;
            this.button3.FlatAppearance.BorderColor = System.Drawing.SystemColors.Control;
            this.button3.FlatAppearance.BorderSize = 0;
            this.button3.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.button3.Location = new System.Drawing.Point(345, 96);
            this.button3.Name = "button3";
            this.button3.Size = new System.Drawing.Size(167, 31);
            this.button3.TabIndex = 8;
            this.button3.Text = "Listeden Kaldır\r\n";
            this.button3.UseVisualStyleBackColor = false;
            this.button3.Click += new System.EventHandler(this.button3_Click);
            // 
            // button4
            // 
            this.button4.BackColor = System.Drawing.Color.SpringGreen;
            this.button4.FlatAppearance.BorderColor = System.Drawing.SystemColors.Control;
            this.button4.FlatAppearance.BorderSize = 0;
            this.button4.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.button4.Location = new System.Drawing.Point(345, 59);
            this.button4.Name = "button4";
            this.button4.Size = new System.Drawing.Size(167, 31);
            this.button4.TabIndex = 8;
            this.button4.Text = "Listeye Ekle";
            this.button4.UseVisualStyleBackColor = false;
            this.button4.Click += new System.EventHandler(this.button4_Click);
            // 
            // Form1
            // 
            this.AllowDrop = true;
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(524, 251);
            this.Controls.Add(this.button4);
            this.Controls.Add(this.button3);
            this.Controls.Add(this.labelLog);
            this.Controls.Add(this.progressBar1);
            this.Controls.Add(this.button2);
            this.Controls.Add(this.button1);
            this.Controls.Add(this.listBox1);
            this.Controls.Add(this.buttonAddConvert);
            this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedToolWindow;
            this.MaximumSize = new System.Drawing.Size(540, 290);
            this.MinimumSize = new System.Drawing.Size(540, 290);
            this.Name = "Form1";
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
            this.Text = "Copyright © Levent Beylerce Since 2025";
            this.Load += new System.EventHandler(this.Form1_Load);
            this.DragDrop += new System.Windows.Forms.DragEventHandler(this.Form1_DragDrop);
            this.DragEnter += new System.Windows.Forms.DragEventHandler(this.Form1_DragEnter);
            this.ResumeLayout(false);

        }

        #endregion
        private System.Windows.Forms.Button buttonAddConvert;
        private ListBox listBox1;
        private Button button1;
        private Button button2;
        private ProgressBar progressBar1;
        private Label labelLog;
        private Button button3;
        private Button button4;
    }
}

