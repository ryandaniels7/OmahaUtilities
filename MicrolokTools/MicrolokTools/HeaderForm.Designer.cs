namespace MicrolokTools
{
    partial class HeaderForm
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
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(HeaderForm));
            this.label1 = new System.Windows.Forms.Label();
            this.LocationBox = new System.Windows.Forms.TextBox();
            this.ProgBox = new System.Windows.Forms.TextBox();
            this.label2 = new System.Windows.Forms.Label();
            this.DateBox = new System.Windows.Forms.TextBox();
            this.label3 = new System.Windows.Forms.Label();
            this.textBox4 = new System.Windows.Forms.TextBox();
            this.label4 = new System.Windows.Forms.Label();
            this.oCancelButton = new System.Windows.Forms.Button();
            this.OKButton = new System.Windows.Forms.Button();
            this.SuspendLayout();
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Location = new System.Drawing.Point(12, 9);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(51, 13);
            this.label1.TabIndex = 0;
            this.label1.Text = "Location:";
            // 
            // LocationBox
            // 
            this.LocationBox.Location = new System.Drawing.Point(69, 9);
            this.LocationBox.Name = "LocationBox";
            this.LocationBox.Size = new System.Drawing.Size(133, 20);
            this.LocationBox.TabIndex = 1;
            // 
            // ProgBox
            // 
            this.ProgBox.Location = new System.Drawing.Point(98, 35);
            this.ProgBox.Name = "ProgBox";
            this.ProgBox.Size = new System.Drawing.Size(104, 20);
            this.ProgBox.TabIndex = 3;
            // 
            // label2
            // 
            this.label2.AutoSize = true;
            this.label2.Location = new System.Drawing.Point(12, 35);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(80, 13);
            this.label2.TabIndex = 2;
            this.label2.Text = "Program Name:";
            // 
            // DateBox
            // 
            this.DateBox.Location = new System.Drawing.Point(51, 61);
            this.DateBox.Name = "DateBox";
            this.DateBox.Size = new System.Drawing.Size(151, 20);
            this.DateBox.TabIndex = 5;
            // 
            // label3
            // 
            this.label3.AutoSize = true;
            this.label3.Location = new System.Drawing.Point(12, 61);
            this.label3.Name = "label3";
            this.label3.Size = new System.Drawing.Size(33, 13);
            this.label3.TabIndex = 4;
            this.label3.Text = "Date:";
            // 
            // textBox4
            // 
            this.textBox4.Location = new System.Drawing.Point(69, 87);
            this.textBox4.Name = "textBox4";
            this.textBox4.Size = new System.Drawing.Size(133, 20);
            this.textBox4.TabIndex = 7;
            this.textBox4.Text = "QuEST";
            // 
            // label4
            // 
            this.label4.AutoSize = true;
            this.label4.Location = new System.Drawing.Point(12, 87);
            this.label4.Name = "label4";
            this.label4.Size = new System.Drawing.Size(52, 13);
            this.label4.TabIndex = 6;
            this.label4.Text = "Designer:";
            // 
            // oCancelButton
            // 
            this.oCancelButton.Location = new System.Drawing.Point(108, 113);
            this.oCancelButton.Name = "oCancelButton";
            this.oCancelButton.Size = new System.Drawing.Size(60, 23);
            this.oCancelButton.TabIndex = 11;
            this.oCancelButton.Text = "Cancel";
            this.oCancelButton.UseVisualStyleBackColor = true;
            this.oCancelButton.Click += new System.EventHandler(this.CancelButton_Click);
            // 
            // OKButton
            // 
            this.OKButton.Location = new System.Drawing.Point(42, 113);
            this.OKButton.Name = "OKButton";
            this.OKButton.Size = new System.Drawing.Size(60, 23);
            this.OKButton.TabIndex = 9;
            this.OKButton.Text = "OK";
            this.OKButton.UseVisualStyleBackColor = true;
            this.OKButton.Click += new System.EventHandler(this.OKButton_Click);
            // 
            // HeaderForm
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(211, 145);
            this.Controls.Add(this.OKButton);
            this.Controls.Add(this.oCancelButton);
            this.Controls.Add(this.textBox4);
            this.Controls.Add(this.label4);
            this.Controls.Add(this.DateBox);
            this.Controls.Add(this.label3);
            this.Controls.Add(this.ProgBox);
            this.Controls.Add(this.label2);
            this.Controls.Add(this.LocationBox);
            this.Controls.Add(this.label1);
            this.Icon = ((System.Drawing.Icon)(resources.GetObject("$this.Icon")));
            this.Name = "HeaderForm";
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
            this.Text = "Header";
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.Label label2;
        private System.Windows.Forms.Label label3;
        private System.Windows.Forms.Label label4;
        private System.Windows.Forms.Button oCancelButton;
        private System.Windows.Forms.Button OKButton;
        public System.Windows.Forms.TextBox LocationBox;
        public System.Windows.Forms.TextBox ProgBox;
        public System.Windows.Forms.TextBox DateBox;
        public System.Windows.Forms.TextBox textBox4;
    }
}