
namespace XmlOutputApp
{
	partial class BoringResult
	{
		/// <summary>
		/// 必要なデザイナー変数です。
		/// </summary>
		private System.ComponentModel.IContainer components = null;

		/// <summary>
		/// 使用中のリソースをすべてクリーンアップします。
		/// </summary>
		/// <param name="disposing">マネージド リソースを破棄する場合は true を指定し、その他の場合は false を指定します。</param>
		protected override void Dispose(bool disposing)
		{
			if (disposing && (components != null))
			{
				components.Dispose();
			}
			base.Dispose(disposing);
		}

		#region Windows フォーム デザイナーで生成されたコード

		/// <summary>
		/// デザイナー サポートに必要なメソッドです。このメソッドの内容を
		/// コード エディターで変更しないでください。
		/// </summary>
		private void InitializeComponent()
		{
			System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(BoringResult));
			this.btnGanban = new System.Windows.Forms.Button();
			this.btnDositu = new System.Windows.Forms.Button();
			this.btnZisuberiAll = new System.Windows.Forms.Button();
			this.btnZisuberiHyou = new System.Windows.Forms.Button();
			this.btnClose = new System.Windows.Forms.Button();
			this.SuspendLayout();
			// 
			// btnGanban
			// 
			this.btnGanban.Location = new System.Drawing.Point(166, 30);
			this.btnGanban.Name = "btnGanban";
			this.btnGanban.Size = new System.Drawing.Size(484, 53);
			this.btnGanban.TabIndex = 0;
			this.btnGanban.Text = "岩盤ボーリング柱状図出力";
			this.btnGanban.UseVisualStyleBackColor = true;
			this.btnGanban.Click += new System.EventHandler(this.btnGanban_Click);
			// 
			// btnDositu
			// 
			this.btnDositu.Location = new System.Drawing.Point(166, 120);
			this.btnDositu.Name = "btnDositu";
			this.btnDositu.Size = new System.Drawing.Size(484, 53);
			this.btnDositu.TabIndex = 1;
			this.btnDositu.Text = "土質ボーリング柱状図(標準貫入試験用)出力";
			this.btnDositu.UseVisualStyleBackColor = true;
			this.btnDositu.Click += new System.EventHandler(this.btnDositu_Click);
			// 
			// btnZisuberiAll
			// 
			this.btnZisuberiAll.Location = new System.Drawing.Point(166, 210);
			this.btnZisuberiAll.Name = "btnZisuberiAll";
			this.btnZisuberiAll.Size = new System.Drawing.Size(484, 53);
			this.btnZisuberiAll.TabIndex = 2;
			this.btnZisuberiAll.Text = "地すべりボーリング柱状図(オールコアボーリング用)出力";
			this.btnZisuberiAll.UseVisualStyleBackColor = true;
			this.btnZisuberiAll.Click += new System.EventHandler(this.btnZisuberiAll_Click);
			// 
			// btnZisuberiHyou
			// 
			this.btnZisuberiHyou.Location = new System.Drawing.Point(166, 300);
			this.btnZisuberiHyou.Name = "btnZisuberiHyou";
			this.btnZisuberiHyou.Size = new System.Drawing.Size(484, 53);
			this.btnZisuberiHyou.TabIndex = 3;
			this.btnZisuberiHyou.Text = "地すべりボーリング柱状図(標準貫入試験用)出力";
			this.btnZisuberiHyou.UseVisualStyleBackColor = true;
			this.btnZisuberiHyou.Click += new System.EventHandler(this.btnZisuberiHyou_Click);
			// 
			// btnClose
			// 
			this.btnClose.Location = new System.Drawing.Point(606, 389);
			this.btnClose.Name = "btnClose";
			this.btnClose.Size = new System.Drawing.Size(144, 49);
			this.btnClose.TabIndex = 4;
			this.btnClose.Text = "閉じる";
			this.btnClose.UseVisualStyleBackColor = true;
			this.btnClose.Click += new System.EventHandler(this.btnClose_Click);
			// 
			// BoringResult
			// 
			this.AutoScaleDimensions = new System.Drawing.SizeF(8F, 15F);
			this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
			this.BackColor = System.Drawing.SystemColors.Control;
			this.ClientSize = new System.Drawing.Size(800, 450);
			this.Controls.Add(this.btnClose);
			this.Controls.Add(this.btnZisuberiHyou);
			this.Controls.Add(this.btnZisuberiAll);
			this.Controls.Add(this.btnDositu);
			this.Controls.Add(this.btnGanban);
			this.Icon = ((System.Drawing.Icon)(resources.GetObject("$this.Icon")));
			this.Name = "BoringResult";
			this.Text = "調査ボーリング結果出力";
			this.ResumeLayout(false);

		}

		#endregion

		private System.Windows.Forms.Button btnGanban;
		private System.Windows.Forms.Button btnDositu;
		private System.Windows.Forms.Button btnZisuberiAll;
		private System.Windows.Forms.Button btnZisuberiHyou;
		private System.Windows.Forms.Button btnClose;
	}
}

