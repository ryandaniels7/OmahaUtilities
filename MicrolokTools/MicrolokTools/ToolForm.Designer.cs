namespace MicrolokTools
{
    partial class ToolForm
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
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(ToolForm));
            this.oCancelButton = new System.Windows.Forms.Button();
            this.ExtensionButton = new System.Windows.Forms.Button();
            this.NonVitalButton = new System.Windows.Forms.Button();
            this.MLLConvertButton = new System.Windows.Forms.Button();
            this.LogBitsButton = new System.Windows.Forms.Button();
            this.QLCPButton = new System.Windows.Forms.Button();
            this.NotesButton = new System.Windows.Forms.Button();
            this.SwitchButton = new System.Windows.Forms.Button();
            this.BooleanButton = new System.Windows.Forms.Button();
            this.SuspendLayout();
            // 
            // oCancelButton
            // 
            this.oCancelButton.Location = new System.Drawing.Point(210, 97);
            this.oCancelButton.Name = "oCancelButton";
            this.oCancelButton.Size = new System.Drawing.Size(83, 37);
            this.oCancelButton.TabIndex = 5;
            this.oCancelButton.Text = "Exit";
            this.oCancelButton.UseVisualStyleBackColor = true;
            this.oCancelButton.Click += new System.EventHandler(this.oCancelButton_Click);
            // 
            // ExtensionButton
            // 
            this.ExtensionButton.AutoSizeMode = System.Windows.Forms.AutoSizeMode.GrowAndShrink;
            this.ExtensionButton.Location = new System.Drawing.Point(121, 54);
            this.ExtensionButton.Name = "ExtensionButton";
            this.ExtensionButton.Size = new System.Drawing.Size(83, 37);
            this.ExtensionButton.TabIndex = 7;
            this.ExtensionButton.Text = "Extension";
            this.ExtensionButton.UseVisualStyleBackColor = true;
            this.ExtensionButton.Click += new System.EventHandler(this.ExtensionButton_Click);
            // 
            // NonVitalButton
            // 
            this.NonVitalButton.Location = new System.Drawing.Point(32, 54);
            this.NonVitalButton.Name = "NonVitalButton";
            this.NonVitalButton.Size = new System.Drawing.Size(83, 37);
            this.NonVitalButton.TabIndex = 6;
            this.NonVitalButton.Text = "Non Vital";
            this.NonVitalButton.UseVisualStyleBackColor = true;
            this.NonVitalButton.Click += new System.EventHandler(this.NonVitalButton_Click);
            // 
            // MLLConvertButton
            // 
            this.MLLConvertButton.Location = new System.Drawing.Point(121, 12);
            this.MLLConvertButton.Name = "MLLConvertButton";
            this.MLLConvertButton.Size = new System.Drawing.Size(83, 37);
            this.MLLConvertButton.TabIndex = 9;
            this.MLLConvertButton.Text = "Convert MLL";
            this.MLLConvertButton.UseVisualStyleBackColor = true;
            this.MLLConvertButton.Click += new System.EventHandler(this.MLLConvertButton_Click);
            // 
            // LogBitsButton
            // 
            this.LogBitsButton.Location = new System.Drawing.Point(32, 12);
            this.LogBitsButton.Name = "LogBitsButton";
            this.LogBitsButton.Size = new System.Drawing.Size(83, 37);
            this.LogBitsButton.TabIndex = 8;
            this.LogBitsButton.Text = "Log Bits";
            this.LogBitsButton.UseVisualStyleBackColor = true;
            this.LogBitsButton.Click += new System.EventHandler(this.LogBitsButton_Click);
            // 
            // QLCPButton
            // 
            this.QLCPButton.Location = new System.Drawing.Point(210, 12);
            this.QLCPButton.Name = "QLCPButton";
            this.QLCPButton.Size = new System.Drawing.Size(83, 37);
            this.QLCPButton.TabIndex = 10;
            this.QLCPButton.Text = "QLCP Builder";
            this.QLCPButton.UseVisualStyleBackColor = true;
            this.QLCPButton.Click += new System.EventHandler(this.QLCPButton_Click);
            // 
            // NotesButton
            // 
            this.NotesButton.Location = new System.Drawing.Point(210, 54);
            this.NotesButton.Name = "NotesButton";
            this.NotesButton.Size = new System.Drawing.Size(83, 37);
            this.NotesButton.TabIndex = 11;
            this.NotesButton.Text = "Remove Notes";
            this.NotesButton.UseVisualStyleBackColor = true;
            this.NotesButton.Click += new System.EventHandler(this.NotesButton_Click);
            // 
            // SwitchButton
            // 
            this.SwitchButton.Location = new System.Drawing.Point(32, 97);
            this.SwitchButton.Name = "SwitchButton";
            this.SwitchButton.Size = new System.Drawing.Size(83, 37);
            this.SwitchButton.TabIndex = 12;
            this.SwitchButton.Text = "Switch Check";
            this.SwitchButton.UseVisualStyleBackColor = true;
            this.SwitchButton.Click += new System.EventHandler(this.SwitchButton_Click);
            // 
            // BooleanButton
            // 
            this.BooleanButton.Location = new System.Drawing.Point(121, 97);
            this.BooleanButton.Name = "BooleanButton";
            this.BooleanButton.Size = new System.Drawing.Size(83, 37);
            this.BooleanButton.TabIndex = 13;
            this.BooleanButton.Text = "Boolean";
            this.BooleanButton.UseVisualStyleBackColor = true;
            this.BooleanButton.Click += new System.EventHandler(this.BooleanButton_Click);
            // 
            // ToolForm
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.AutoSizeMode = System.Windows.Forms.AutoSizeMode.GrowAndShrink;
            this.ClientSize = new System.Drawing.Size(323, 146);
            this.Controls.Add(this.BooleanButton);
            this.Controls.Add(this.SwitchButton);
            this.Controls.Add(this.NotesButton);
            this.Controls.Add(this.QLCPButton);
            this.Controls.Add(this.MLLConvertButton);
            this.Controls.Add(this.LogBitsButton);
            this.Controls.Add(this.ExtensionButton);
            this.Controls.Add(this.NonVitalButton);
            this.Controls.Add(this.oCancelButton);
            //this.Icon = ((System.Drawing.Icon)(resources.GetObject("$this.Icon")));
            this.MaximizeBox = false;
            this.MinimizeBox = false;
            this.Name = "ToolForm";
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
            this.Text = " Microlok";
            this.Load += new System.EventHandler(this.ToolForm_Load);
            this.ResumeLayout(false);

        }

        #endregion

        private System.Windows.Forms.Button oCancelButton;
        private System.Windows.Forms.Button ExtensionButton;
        private System.Windows.Forms.Button NonVitalButton;
        private System.Windows.Forms.Button MLLConvertButton;
        private System.Windows.Forms.Button LogBitsButton;
        private System.Windows.Forms.Button QLCPButton;
        private System.Windows.Forms.Button NotesButton;
        private System.Windows.Forms.Button SwitchButton;
        private System.Windows.Forms.Button BooleanButton;
    }
}

