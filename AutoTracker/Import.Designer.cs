namespace AutoTracker
{
    partial class Import
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
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(Import));
            this.btn_Browse1 = new System.Windows.Forms.Button();
            this.inputBox1 = new System.Windows.Forms.TextBox();
            this.inputBox2 = new System.Windows.Forms.TextBox();
            this.btn_Browse2 = new System.Windows.Forms.Button();
            this.button1 = new System.Windows.Forms.Button();
            this.outputBox1 = new System.Windows.Forms.TextBox();
            this.newFileTextBox = new System.Windows.Forms.TextBox();
            this.cancelBtn = new System.Windows.Forms.Button();
            this.inputBox3 = new System.Windows.Forms.TextBox();
            this.openTier = new System.Windows.Forms.Button();
            this.checkBox1 = new System.Windows.Forms.CheckBox();
            this.SuspendLayout();
            // 
            // btn_Browse1
            // 
            this.btn_Browse1.BackColor = System.Drawing.SystemColors.ButtonHighlight;
            this.btn_Browse1.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.btn_Browse1.Font = new System.Drawing.Font("Calibri", 9.75F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.btn_Browse1.Location = new System.Drawing.Point(12, 18);
            this.btn_Browse1.Margin = new System.Windows.Forms.Padding(3, 4, 3, 4);
            this.btn_Browse1.Name = "btn_Browse1";
            this.btn_Browse1.Size = new System.Drawing.Size(105, 30);
            this.btn_Browse1.TabIndex = 0;
            this.btn_Browse1.Text = "Open ASU";
            this.btn_Browse1.UseVisualStyleBackColor = false;
            this.btn_Browse1.Click += new System.EventHandler(this.btn_Browse1_Click);
            // 
            // inputBox1
            // 
            this.inputBox1.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle;
            this.inputBox1.Enabled = false;
            this.inputBox1.Location = new System.Drawing.Point(123, 21);
            this.inputBox1.Margin = new System.Windows.Forms.Padding(3, 4, 3, 4);
            this.inputBox1.Name = "inputBox1";
            this.inputBox1.Size = new System.Drawing.Size(484, 25);
            this.inputBox1.TabIndex = 1;
            // 
            // inputBox2
            // 
            this.inputBox2.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle;
            this.inputBox2.Enabled = false;
            this.inputBox2.Location = new System.Drawing.Point(123, 59);
            this.inputBox2.Margin = new System.Windows.Forms.Padding(3, 4, 3, 4);
            this.inputBox2.Name = "inputBox2";
            this.inputBox2.Size = new System.Drawing.Size(484, 25);
            this.inputBox2.TabIndex = 3;
            // 
            // btn_Browse2
            // 
            this.btn_Browse2.BackColor = System.Drawing.SystemColors.ButtonHighlight;
            this.btn_Browse2.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.btn_Browse2.Font = new System.Drawing.Font("Calibri", 9.75F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.btn_Browse2.Location = new System.Drawing.Point(12, 56);
            this.btn_Browse2.Margin = new System.Windows.Forms.Padding(3, 4, 3, 4);
            this.btn_Browse2.Name = "btn_Browse2";
            this.btn_Browse2.Size = new System.Drawing.Size(105, 30);
            this.btn_Browse2.TabIndex = 2;
            this.btn_Browse2.Text = "Open UMD";
            this.btn_Browse2.UseVisualStyleBackColor = false;
            this.btn_Browse2.Click += new System.EventHandler(this.btn_Browse2_Click);
            // 
            // button1
            // 
            this.button1.BackColor = System.Drawing.SystemColors.ButtonHighlight;
            this.button1.Enabled = false;
            this.button1.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.button1.Font = new System.Drawing.Font("Calibri", 11.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.button1.Location = new System.Drawing.Point(505, 141);
            this.button1.Margin = new System.Windows.Forms.Padding(3, 4, 3, 4);
            this.button1.Name = "button1";
            this.button1.Size = new System.Drawing.Size(100, 30);
            this.button1.TabIndex = 4;
            this.button1.Text = "PROCESS";
            this.button1.UseVisualStyleBackColor = false;
            this.button1.Click += new System.EventHandler(this.button1_Click);
            // 
            // outputBox1
            // 
            this.outputBox1.Location = new System.Drawing.Point(123, 146);
            this.outputBox1.Margin = new System.Windows.Forms.Padding(3, 4, 3, 4);
            this.outputBox1.Name = "outputBox1";
            this.outputBox1.Size = new System.Drawing.Size(339, 25);
            this.outputBox1.TabIndex = 5;
            this.outputBox1.Text = "CSV OUTPUT";
            this.outputBox1.Visible = false;
            // 
            // newFileTextBox
            // 
            this.newFileTextBox.Location = new System.Drawing.Point(123, 117);
            this.newFileTextBox.Margin = new System.Windows.Forms.Padding(3, 4, 3, 4);
            this.newFileTextBox.Name = "newFileTextBox";
            this.newFileTextBox.Size = new System.Drawing.Size(339, 25);
            this.newFileTextBox.TabIndex = 9;
            this.newFileTextBox.Text = "XML OUTPUT";
            this.newFileTextBox.Visible = false;
            // 
            // cancelBtn
            // 
            this.cancelBtn.BackColor = System.Drawing.SystemColors.ButtonHighlight;
            this.cancelBtn.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.cancelBtn.Font = new System.Drawing.Font("Calibri", 11.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.cancelBtn.Location = new System.Drawing.Point(399, 141);
            this.cancelBtn.Margin = new System.Windows.Forms.Padding(3, 4, 3, 4);
            this.cancelBtn.Name = "cancelBtn";
            this.cancelBtn.Size = new System.Drawing.Size(100, 30);
            this.cancelBtn.TabIndex = 10;
            this.cancelBtn.Text = "CANCEL";
            this.cancelBtn.UseVisualStyleBackColor = false;
            this.cancelBtn.Click += new System.EventHandler(this.cancelBtn_Click);
            //
            // inputBox3
            //
            this.inputBox3.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle;
            this.inputBox3.Enabled = false;
            this.inputBox3.Location = new System.Drawing.Point(123, 97);
            this.inputBox3.Margin = new System.Windows.Forms.Padding(3, 4, 3, 4);
            this.inputBox3.Name = "inputBox3";
            this.inputBox3.Size = new System.Drawing.Size(484, 25);
            this.inputBox3.TabIndex = 12;
            // 
            // openTier
            // 
            this.openTier.BackColor = System.Drawing.SystemColors.ButtonHighlight;
            this.openTier.Enabled = false;
            this.openTier.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.openTier.Font = new System.Drawing.Font("Calibri", 10.2F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.openTier.Location = new System.Drawing.Point(12, 96);
            this.openTier.Margin = new System.Windows.Forms.Padding(3, 4, 3, 4);
            this.openTier.Name = "openTier";
            this.openTier.Size = new System.Drawing.Size(105, 30);
            this.openTier.TabIndex = 11;
            this.openTier.Text = "Tier Data";
            this.openTier.UseVisualStyleBackColor = false;
            this.openTier.Click += new System.EventHandler(this.openTier_Click);
            // 
            // checkBox1
            // 
            this.checkBox1.AutoSize = true;
            this.checkBox1.Location = new System.Drawing.Point(13, 132);
            this.checkBox1.Name = "checkBox1";
            this.checkBox1.Size = new System.Drawing.Size(83,24);
            this.checkBox1.TabIndex = 13;
            this.checkBox1.Text = "Tier Data";
            this.checkBox1.UseVisualStyleBackColor = true;
            this.checkBox1.CheckedChanged += new System.EventHandler(this.checkBox1_CheckedChanged);
            // 
            // Import
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(7F, 20F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(619, 177);
            this.Controls.Add(this.checkBox1);
            this.Controls.Add(this.inputBox3);
            this.Controls.Add(this.openTier);
            this.Controls.Add(this.cancelBtn);
            this.Controls.Add(this.newFileTextBox);
            this.Controls.Add(this.outputBox1);
            this.Controls.Add(this.button1);
            this.Controls.Add(this.inputBox2);
            this.Controls.Add(this.btn_Browse2);
            this.Controls.Add(this.inputBox1);
            this.Controls.Add(this.btn_Browse1);
            this.Font = new System.Drawing.Font("Arial Narrow", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedToolWindow;
            this.Icon = ((System.Drawing.Icon)(resources.GetObject("$this.Icon")));
            this.Margin = new System.Windows.Forms.Padding(3, 4, 3, 4);
            this.Name = "Import";
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterParent;
            this.Text = "Import Excel Files";
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.Button btn_Browse1;
        private System.Windows.Forms.TextBox inputBox1;
        private System.Windows.Forms.TextBox inputBox2;
        private System.Windows.Forms.Button btn_Browse2;
        private System.Windows.Forms.Button button1;
        private System.Windows.Forms.TextBox outputBox1;
        private System.Windows.Forms.TextBox newFileTextBox;
        private System.Windows.Forms.Button cancelBtn;
        private System.Windows.Forms.TextBox inputBox3;
        private System.Windows.Forms.Button openTier;
        private System.Windows.Forms.CheckBox checkBox1;
    }
}

