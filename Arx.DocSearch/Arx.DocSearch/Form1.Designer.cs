namespace Arx.DocSearch
{
	partial class MainForm
	{
		/// <summary>
		/// 必要なデザイナー変数です。
		/// </summary>
		private System.ComponentModel.IContainer components = null;

		/// <summary>
		/// 使用中のリソースをすべてクリーンアップします。
		/// </summary>
		/// <param name="disposing">マネージ リソースが破棄される場合 true、破棄されない場合は false です。</param>
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
			this.srcLabel = new System.Windows.Forms.Label();
			this.srcText = new System.Windows.Forms.TextBox();
			this.srcButton = new System.Windows.Forms.Button();
			this.targetButton = new System.Windows.Forms.Button();
			this.targetText = new System.Windows.Forms.TextBox();
			this.targetLabel = new System.Windows.Forms.Label();
			this.listView1 = new System.Windows.Forms.ListView();
			this.openFileDialog1 = new System.Windows.Forms.OpenFileDialog();
			this.folderBrowserDialog1 = new System.Windows.Forms.FolderBrowserDialog();
			this.SuspendLayout();
			// 
			// srcLabel
			// 
			this.srcLabel.AutoSize = true;
			this.srcLabel.Location = new System.Drawing.Point(22, 26);
			this.srcLabel.Name = "srcLabel";
			this.srcLabel.Size = new System.Drawing.Size(65, 12);
			this.srcLabel.TabIndex = 0;
			this.srcLabel.Text = "検索テキスト";
			// 
			// srcText
			// 
			this.srcText.Location = new System.Drawing.Point(93, 21);
			this.srcText.Name = "srcText";
			this.srcText.Size = new System.Drawing.Size(181, 19);
			this.srcText.TabIndex = 1;
			// 
			// srcButton
			// 
			this.srcButton.Location = new System.Drawing.Point(289, 20);
			this.srcButton.Name = "srcButton";
			this.srcButton.Size = new System.Drawing.Size(56, 19);
			this.srcButton.TabIndex = 2;
			this.srcButton.Text = "選択";
			this.srcButton.UseVisualStyleBackColor = true;
			// 
			// targetButton
			// 
			this.targetButton.Location = new System.Drawing.Point(289, 58);
			this.targetButton.Name = "targetButton";
			this.targetButton.Size = new System.Drawing.Size(56, 19);
			this.targetButton.TabIndex = 5;
			this.targetButton.Text = "選択";
			this.targetButton.UseVisualStyleBackColor = true;
			// 
			// targetText
			// 
			this.targetText.Location = new System.Drawing.Point(93, 59);
			this.targetText.Name = "targetText";
			this.targetText.Size = new System.Drawing.Size(181, 19);
			this.targetText.TabIndex = 4;
			// 
			// targetLabel
			// 
			this.targetLabel.AutoSize = true;
			this.targetLabel.Location = new System.Drawing.Point(22, 64);
			this.targetLabel.Name = "targetLabel";
			this.targetLabel.Size = new System.Drawing.Size(41, 12);
			this.targetLabel.TabIndex = 3;
			this.targetLabel.Text = "検索先";
			// 
			// listView1
			// 
			this.listView1.Location = new System.Drawing.Point(30, 103);
			this.listView1.Name = "listView1";
			this.listView1.Size = new System.Drawing.Size(314, 197);
			this.listView1.TabIndex = 6;
			this.listView1.UseCompatibleStateImageBehavior = false;
			// 
			// openFileDialog1
			// 
			this.openFileDialog1.FileName = "openFileDialog1";
			// 
			// MainForm
			// 
			this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 12F);
			this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
			this.ClientSize = new System.Drawing.Size(370, 331);
			this.Controls.Add(this.listView1);
			this.Controls.Add(this.targetButton);
			this.Controls.Add(this.targetText);
			this.Controls.Add(this.targetLabel);
			this.Controls.Add(this.srcButton);
			this.Controls.Add(this.srcText);
			this.Controls.Add(this.srcLabel);
			this.Name = "MainForm";
			this.Text = "Form1";
			this.ResumeLayout(false);
			this.PerformLayout();

		}

		#endregion

		private System.Windows.Forms.Label srcLabel;
		private System.Windows.Forms.TextBox srcText;
		private System.Windows.Forms.Button srcButton;
		private System.Windows.Forms.Button targetButton;
		private System.Windows.Forms.TextBox targetText;
		private System.Windows.Forms.Label targetLabel;
		private System.Windows.Forms.ListView listView1;
		private System.Windows.Forms.OpenFileDialog openFileDialog1;
		private System.Windows.Forms.FolderBrowserDialog folderBrowserDialog1;
	}
}

