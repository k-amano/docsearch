namespace Arx.DocSearch.Client
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
			this.logButton = new System.Windows.Forms.Button();
			this.searchJpButton = new System.Windows.Forms.Button();
			this.countLabel = new System.Windows.Forms.Label();
			this.roughlinesLabel = new System.Windows.Forms.Label();
			this.roughLinesText = new System.Windows.Forms.TextBox();
			this.wordCountText = new System.Windows.Forms.TextBox();
			this.matchCountLabel = new System.Windows.Forms.Label();
			this.totalCountLabel = new System.Windows.Forms.Label();
			this.conditionLabel = new System.Windows.Forms.Label();
			this.rateText = new System.Windows.Forms.TextBox();
			this.rateLabel = new System.Windows.Forms.Label();
			this.messageLabel = new System.Windows.Forms.Label();
			this.textFileButton = new System.Windows.Forms.Button();
			this.indexButton = new System.Windows.Forms.Button();
			this.listView1 = new System.Windows.Forms.ListView();
			this.targetButton = new System.Windows.Forms.Button();
			this.openFileDialog2 = new System.Windows.Forms.OpenFileDialog();
			this.compareButton = new System.Windows.Forms.Button();
			this.searchButton = new System.Windows.Forms.Button();
			this.targetText = new System.Windows.Forms.TextBox();
			this.folderBrowserDialog1 = new System.Windows.Forms.FolderBrowserDialog();
			this.targetLabel = new System.Windows.Forms.Label();
			this.srcButton = new System.Windows.Forms.Button();
			this.openFileDialog1 = new System.Windows.Forms.OpenFileDialog();
			this.srcLabel = new System.Windows.Forms.Label();
			this.progressText = new System.Windows.Forms.RichTextBox();
			this.folderBrowserDialog2 = new System.Windows.Forms.FolderBrowserDialog();
			this.srcCombo = new System.Windows.Forms.ComboBox();
			this.SuspendLayout();
			// 
			// logButton
			// 
			this.logButton.Location = new System.Drawing.Point(620, 454);
			this.logButton.Name = "logButton";
			this.logButton.Size = new System.Drawing.Size(75, 23);
			this.logButton.TabIndex = 47;
			this.logButton.Text = "ログを開く";
			this.logButton.UseVisualStyleBackColor = true;
			this.logButton.Click += new System.EventHandler(this.logButton_Click);
			// 
			// searchJpButton
			// 
			this.searchJpButton.Location = new System.Drawing.Point(648, 73);
			this.searchJpButton.Name = "searchJpButton";
			this.searchJpButton.Size = new System.Drawing.Size(64, 23);
			this.searchJpButton.TabIndex = 46;
			this.searchJpButton.Text = "和文検索";
			this.searchJpButton.UseVisualStyleBackColor = true;
			this.searchJpButton.Click += new System.EventHandler(this.searchJpButton_Click);
			// 
			// countLabel
			// 
			this.countLabel.AutoSize = true;
			this.countLabel.Location = new System.Drawing.Point(27, 486);
			this.countLabel.Name = "countLabel";
			this.countLabel.Size = new System.Drawing.Size(35, 12);
			this.countLabel.TabIndex = 45;
			this.countLabel.Text = "label1";
			// 
			// roughlinesLabel
			// 
			this.roughlinesLabel.AutoSize = true;
			this.roughlinesLabel.Location = new System.Drawing.Point(606, 80);
			this.roughlinesLabel.Name = "roughlinesLabel";
			this.roughlinesLabel.Size = new System.Drawing.Size(17, 12);
			this.roughlinesLabel.TabIndex = 44;
			this.roughlinesLabel.Text = "文";
			// 
			// roughLinesText
			// 
			this.roughLinesText.Location = new System.Drawing.Point(556, 76);
			this.roughLinesText.Name = "roughLinesText";
			this.roughLinesText.Size = new System.Drawing.Size(44, 19);
			this.roughLinesText.TabIndex = 43;
			this.roughLinesText.TextAlign = System.Windows.Forms.HorizontalAlignment.Right;
			// 
			// wordCountText
			// 
			this.wordCountText.Location = new System.Drawing.Point(405, 76);
			this.wordCountText.Name = "wordCountText";
			this.wordCountText.Size = new System.Drawing.Size(44, 19);
			this.wordCountText.TabIndex = 42;
			this.wordCountText.TextAlign = System.Windows.Forms.HorizontalAlignment.Right;
			// 
			// matchCountLabel
			// 
			this.matchCountLabel.AutoSize = true;
			this.matchCountLabel.Location = new System.Drawing.Point(312, 80);
			this.matchCountLabel.Name = "matchCountLabel";
			this.matchCountLabel.Size = new System.Drawing.Size(87, 12);
			this.matchCountLabel.TabIndex = 41;
			this.matchCountLabel.Text = "%以上、 word 数";
			// 
			// totalCountLabel
			// 
			this.totalCountLabel.AutoSize = true;
			this.totalCountLabel.Location = new System.Drawing.Point(20, 74);
			this.totalCountLabel.Name = "totalCountLabel";
			this.totalCountLabel.Size = new System.Drawing.Size(85, 12);
			this.totalCountLabel.TabIndex = 40;
			this.totalCountLabel.Text = "totalCountLabel";
			// 
			// conditionLabel
			// 
			this.conditionLabel.AutoSize = true;
			this.conditionLabel.Location = new System.Drawing.Point(449, 80);
			this.conditionLabel.Name = "conditionLabel";
			this.conditionLabel.Size = new System.Drawing.Size(101, 12);
			this.conditionLabel.TabIndex = 39;
			this.conditionLabel.Text = "以上、ラフ検索単位";
			// 
			// rateText
			// 
			this.rateText.Location = new System.Drawing.Point(262, 77);
			this.rateText.Name = "rateText";
			this.rateText.Size = new System.Drawing.Size(44, 19);
			this.rateText.TabIndex = 38;
			this.rateText.TextAlign = System.Windows.Forms.HorizontalAlignment.Right;
			// 
			// rateLabel
			// 
			this.rateLabel.AutoSize = true;
			this.rateLabel.Location = new System.Drawing.Point(215, 81);
			this.rateLabel.Name = "rateLabel";
			this.rateLabel.Size = new System.Drawing.Size(41, 12);
			this.rateLabel.TabIndex = 37;
			this.rateLabel.Text = "一致率";
			// 
			// messageLabel
			// 
			this.messageLabel.AutoSize = true;
			this.messageLabel.Location = new System.Drawing.Point(27, 465);
			this.messageLabel.Name = "messageLabel";
			this.messageLabel.Size = new System.Drawing.Size(35, 12);
			this.messageLabel.TabIndex = 36;
			this.messageLabel.Text = "label1";
			// 
			// textFileButton
			// 
			this.textFileButton.Location = new System.Drawing.Point(682, 45);
			this.textFileButton.Name = "textFileButton";
			this.textFileButton.Size = new System.Drawing.Size(99, 21);
			this.textFileButton.TabIndex = 35;
			this.textFileButton.Text = "テキスト抽出";
			this.textFileButton.UseVisualStyleBackColor = true;
			this.textFileButton.Click += new System.EventHandler(this.textFileButton_Click);
			// 
			// indexButton
			// 
			this.indexButton.Location = new System.Drawing.Point(682, 18);
			this.indexButton.Name = "indexButton";
			this.indexButton.Size = new System.Drawing.Size(99, 21);
			this.indexButton.TabIndex = 34;
			this.indexButton.Text = "インデックス作成";
			this.indexButton.UseVisualStyleBackColor = true;
			this.indexButton.Click += new System.EventHandler(this.indexButton_Click);
			// 
			// listView1
			// 
			this.listView1.Location = new System.Drawing.Point(29, 103);
			this.listView1.Name = "listView1";
			this.listView1.Size = new System.Drawing.Size(752, 213);
			this.listView1.TabIndex = 31;
			this.listView1.UseCompatibleStateImageBehavior = false;
			// 
			// targetButton
			// 
			this.targetButton.Location = new System.Drawing.Point(620, 45);
			this.targetButton.Name = "targetButton";
			this.targetButton.Size = new System.Drawing.Size(56, 21);
			this.targetButton.TabIndex = 30;
			this.targetButton.Text = "選択";
			this.targetButton.UseVisualStyleBackColor = true;
			this.targetButton.Click += new System.EventHandler(this.targetButton_Click);
			// 
			// openFileDialog2
			// 
			this.openFileDialog2.FileName = "openFileDialog2";
			// 
			// compareButton
			// 
			this.compareButton.Location = new System.Drawing.Point(706, 454);
			this.compareButton.Name = "compareButton";
			this.compareButton.Size = new System.Drawing.Size(75, 23);
			this.compareButton.TabIndex = 33;
			this.compareButton.Text = "比較";
			this.compareButton.UseVisualStyleBackColor = true;
			this.compareButton.Click += new System.EventHandler(this.compareButton_Click);
			// 
			// searchButton
			// 
			this.searchButton.Location = new System.Drawing.Point(718, 73);
			this.searchButton.Name = "searchButton";
			this.searchButton.Size = new System.Drawing.Size(64, 23);
			this.searchButton.TabIndex = 32;
			this.searchButton.Text = "英文検索";
			this.searchButton.UseVisualStyleBackColor = true;
			this.searchButton.Click += new System.EventHandler(this.searchButton_Click);
			// 
			// targetText
			// 
			this.targetText.Location = new System.Drawing.Point(91, 45);
			this.targetText.Name = "targetText";
			this.targetText.Size = new System.Drawing.Size(523, 19);
			this.targetText.TabIndex = 29;
			// 
			// targetLabel
			// 
			this.targetLabel.AutoSize = true;
			this.targetLabel.Location = new System.Drawing.Point(20, 50);
			this.targetLabel.Name = "targetLabel";
			this.targetLabel.Size = new System.Drawing.Size(41, 12);
			this.targetLabel.TabIndex = 28;
			this.targetLabel.Text = "検索先";
			// 
			// srcButton
			// 
			this.srcButton.Location = new System.Drawing.Point(620, 18);
			this.srcButton.Name = "srcButton";
			this.srcButton.Size = new System.Drawing.Size(56, 21);
			this.srcButton.TabIndex = 27;
			this.srcButton.Text = "選択";
			this.srcButton.UseVisualStyleBackColor = true;
			this.srcButton.Click += new System.EventHandler(this.srcButton_Click);
			// 
			// openFileDialog1
			// 
			this.openFileDialog1.FileName = "openFileDialog1";
			// 
			// srcLabel
			// 
			this.srcLabel.AutoSize = true;
			this.srcLabel.Location = new System.Drawing.Point(20, 25);
			this.srcLabel.Name = "srcLabel";
			this.srcLabel.Size = new System.Drawing.Size(65, 12);
			this.srcLabel.TabIndex = 25;
			this.srcLabel.Text = "検索テキスト";
			// 
			// progressText
			// 
			this.progressText.Location = new System.Drawing.Point(29, 322);
			this.progressText.Name = "progressText";
			this.progressText.ReadOnly = true;
			this.progressText.ScrollBars = System.Windows.Forms.RichTextBoxScrollBars.ForcedVertical;
			this.progressText.Size = new System.Drawing.Size(752, 126);
			this.progressText.TabIndex = 48;
			this.progressText.Text = "";
			// 
			// srcCombo
			// 
			this.srcCombo.FormattingEnabled = true;
			this.srcCombo.Location = new System.Drawing.Point(93, 20);
			this.srcCombo.Name = "srcCombo";
			this.srcCombo.Size = new System.Drawing.Size(523, 20);
			this.srcCombo.TabIndex = 49;
			// 
			// MainForm
			// 
			this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 12F);
			this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
			this.ClientSize = new System.Drawing.Size(803, 509);
			this.Controls.Add(this.srcCombo);
			this.Controls.Add(this.progressText);
			this.Controls.Add(this.logButton);
			this.Controls.Add(this.searchJpButton);
			this.Controls.Add(this.countLabel);
			this.Controls.Add(this.roughlinesLabel);
			this.Controls.Add(this.roughLinesText);
			this.Controls.Add(this.wordCountText);
			this.Controls.Add(this.matchCountLabel);
			this.Controls.Add(this.totalCountLabel);
			this.Controls.Add(this.conditionLabel);
			this.Controls.Add(this.rateText);
			this.Controls.Add(this.rateLabel);
			this.Controls.Add(this.messageLabel);
			this.Controls.Add(this.textFileButton);
			this.Controls.Add(this.indexButton);
			this.Controls.Add(this.listView1);
			this.Controls.Add(this.targetButton);
			this.Controls.Add(this.compareButton);
			this.Controls.Add(this.searchButton);
			this.Controls.Add(this.targetText);
			this.Controls.Add(this.targetLabel);
			this.Controls.Add(this.srcButton);
			this.Controls.Add(this.srcLabel);
			this.Name = "MainForm";
			this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
			this.Text = "重複文書検索システム";
			this.FormClosing += new System.Windows.Forms.FormClosingEventHandler(this.MainForm_FormClosing);
			this.Load += new System.EventHandler(this.MainForm_Load);
			this.ResumeLayout(false);
			this.PerformLayout();

		}

		#endregion

		private System.Windows.Forms.Button logButton;
		private System.Windows.Forms.Button searchJpButton;
		private System.Windows.Forms.Label countLabel;
		private System.Windows.Forms.Label roughlinesLabel;
		private System.Windows.Forms.TextBox roughLinesText;
		private System.Windows.Forms.TextBox wordCountText;
		private System.Windows.Forms.Label matchCountLabel;
		private System.Windows.Forms.Label totalCountLabel;
		private System.Windows.Forms.Label conditionLabel;
		private System.Windows.Forms.TextBox rateText;
		private System.Windows.Forms.Label rateLabel;
		private System.Windows.Forms.Label messageLabel;
		private System.Windows.Forms.Button textFileButton;
		private System.Windows.Forms.Button indexButton;
		private System.Windows.Forms.ListView listView1;
		private System.Windows.Forms.Button targetButton;
		private System.Windows.Forms.OpenFileDialog openFileDialog2;
		private System.Windows.Forms.Button compareButton;
		private System.Windows.Forms.Button searchButton;
		private System.Windows.Forms.TextBox targetText;
		private System.Windows.Forms.FolderBrowserDialog folderBrowserDialog1;
		private System.Windows.Forms.Label targetLabel;
		private System.Windows.Forms.Button srcButton;
		private System.Windows.Forms.OpenFileDialog openFileDialog1;
		private System.Windows.Forms.Label srcLabel;
		private System.Windows.Forms.RichTextBox progressText;
		private System.Windows.Forms.FolderBrowserDialog folderBrowserDialog2;
		private System.Windows.Forms.ComboBox srcCombo;
	}
}

