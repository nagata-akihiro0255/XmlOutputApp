namespace WorldGyomu.CommonControl
{
    partial class WaitSplash
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
			this.messageLabel = new System.Windows.Forms.Label();
			this.attention = new System.Windows.Forms.Label();
			this.SuspendLayout();
			// 
			// messageLabel
			// 
			this.messageLabel.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
            | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
			this.messageLabel.BackColor = System.Drawing.Color.WhiteSmoke;
			this.messageLabel.Font = new System.Drawing.Font("MS UI Gothic", 20.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(128)));
			this.messageLabel.Location = new System.Drawing.Point(7, 6);
			this.messageLabel.Margin = new System.Windows.Forms.Padding(4, 0, 4, 0);
			this.messageLabel.Name = "messageLabel";
			this.messageLabel.Size = new System.Drawing.Size(652, 329);
			this.messageLabel.TabIndex = 2;
			this.messageLabel.Text = "処理中です、しばらくお待ち下さい。";
			this.messageLabel.TextAlign = System.Drawing.ContentAlignment.MiddleCenter;
			// 
			// attention
			// 
			this.attention.AutoSize = true;
			this.attention.Location = new System.Drawing.Point(172, 241);
			this.attention.Name = "attention";
			this.attention.Size = new System.Drawing.Size(302, 15);
			this.attention.TabIndex = 3;
			this.attention.Text = "＊処理が完了するまでPCをこのままにしてください。";
			this.attention.Click += new System.EventHandler(this.attention_Click);
			// 
			// WaitSplash
			// 
			this.AutoScaleDimensions = new System.Drawing.SizeF(8F, 15F);
			this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
			this.BackColor = System.Drawing.Color.Silver;
			this.ClientSize = new System.Drawing.Size(665, 341);
			this.Controls.Add(this.attention);
			this.Controls.Add(this.messageLabel);
			this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.None;
			this.Margin = new System.Windows.Forms.Padding(4);
			this.Name = "WaitSplash";
			this.ShowInTaskbar = false;
			this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
			this.Text = "WaitSplash";
			this.ResumeLayout(false);
			this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.Label messageLabel;
		private System.Windows.Forms.Label attention;
	}
}