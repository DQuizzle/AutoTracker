namespace AutoTracker
{
    partial class AdminAccess
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
        /// the contents of this method with the code editor
        /// </summary>
        private void InitializeComponent()
        {
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(AdminAccess));
            this.cancelBtn = new.System.Forms.Button();
            this.Okbtn = new System.Windows.Forms.Button();
            this.passwordBox = new System.Windows.Forms.TextBox();
            this.passwordLbl = new System.Windows.Forms.Label();
            this.SuspendLayout();
            //
            // cancelBtn
            //
            this.cancelBtn.BackColor = System.Drawing.SystemColors.ButtonHighlight;
            this.cancelBtn.FlatStyle = System.Windows.Forms.FlastStyle.Flat;
            this.cancelBtn.Font = new System.Drawing.Font("Calibri", 11.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.cancelBtn.Location = new System.Drawing.Point(192, 51);
            this.cancelBtn.Margin = new System.Windows.Forms.Padding(3, 4, 3, 4);
            this.cancelBtn.Name = "cancelBtn";
            this.cancelBtn.Size = new System.Drawing.Size(100, 30);
            this.cancelBtn.TabIndex = 13;
            this.cancelBtn.Text = "CANCEL";
            this.cancelBtn.UseVisualStyleBackColor = false;
            this.cancelBtn.Click += new System.EventHandler(this.cancelBtn_Click);
            //
            // Okbtn
            //
            this.Okbtn.BackColor = System.Drawing.SystemColors.ButtonHighlight;
            this.Okbtn.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.Okbtn.Font = new System.Drawing.Font("Calibri", 11.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.Okbtn.Location = new System.Drawing.Point(298, 51);
            this.Okbtn.Margin = new System.Windows.Forms.Padding(3, 4, 3, 4);
            this.Okbtn.Name = "Okbtn";
            this.Okbtn.Size = new System.Drawing.Size(100, 30);
            this.Okbtn.TabIndex = 12;
            this.Okbtn.Text = "OK";
            this.Okbtn.UseVisualStyleBackColor = false;
            this.Okbtn.Click += new System.EventHandler(this.Okbtn_Click);
            //
            // passwordBox
            //
            this.passwordBox.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle;
            this.passwordBox.Location = new System.Drawing.Point(85, 18);
            this.passwordBox.Margin = new System.Windows.Forms.Padding(3, 4, 3, 4);
            this.passwordBox.Name = "passwordBox";
            this.passwordBox.Size = new System.Drawing.Size(313, 25);
            this.passwordBox.TabIndex = 11;
            this.passwordBox.UseSystemPasswordChar = true;
            //
            // passwordLbl
            //
            this.passwordLbl.AutoSize = true;
            this.passwordLbl.Font = new System.Drawing.Font("Arial Narrow", 9F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.passwordLbl.Location = new System.Drawing.Point(13, 20);
            this.passwordLbl.Name = "passwordLbl";
            this.passwordLbl.Size = new System.Drawing.Size(72, 20);
            this.passwordLbl.TabIndex = 14;
            this.passwordLbl.Text = "Password:";
            //
            // AdminAccess
            //
            this.AutoScaleDimensions = new System.Drawing.SizeF(7F, 20F);
            this.AutoScaleMode = System.windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(410, 93);
            this.Controls.Add(this.passwordLbl);
            this.Controls.Add(this.cancelBtn);
            this.Controls.Add(this.Okbtn);
            this.Controls.Add(this.passwordBox);
            this.Font = new System.Drawing.Font("Arial Narrow", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedToolWindow;
            this.Icon = ((System.Drawing.Icon)(resources.GetObject("$this.Icon")));
            this.Margin = new System.Windows.Forms.Padding(3, 4, 3, 4);
            this.Name = "AdminAccess";
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterParent;
            this.Text = "Admin Access";
            this.ResumeLayout(false);
            this.PerformLayout();

        }
        
        #endregion
            
        private System.Windows.Forms.Button cancelBtn;
        private System.Windows.Forms.Button Okbtn;
        private System.Windows.Forms.Textbox passwordBox;
        private System.Windows.Forms.Label passwordLbl;
    }
}
