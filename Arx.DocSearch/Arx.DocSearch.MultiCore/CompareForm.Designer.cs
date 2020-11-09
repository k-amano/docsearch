namespace Arx.DocSearch.MultiCore
{
	partial class CompareForm
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
			this.components = new System.ComponentModel.Container();
			this.contextMenuStrip1 = new System.Windows.Forms.ContextMenuStrip(this.components);
			this.コピーCToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
			this.検索SToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
			this.buttonPanel = new System.Windows.Forms.Panel();
			this.macthCountButton = new System.Windows.Forms.Button();
			this.closeButton = new System.Windows.Forms.Button();
			this.splitContainer1 = new System.Windows.Forms.SplitContainer();
			this.srcText = new System.Windows.Forms.RichTextBox();
			this.targetText = new System.Windows.Forms.RichTextBox();
			this.contextMenuStrip2 = new System.Windows.Forms.ContextMenuStrip(this.components);
			this.コピーCToolStripMenuItem1 = new System.Windows.Forms.ToolStripMenuItem();
			this.検索SToolStripMenuItem1 = new System.Windows.Forms.ToolStripMenuItem();
			this.infoLabel = new System.Windows.Forms.Label();
			this.bottpmPanel = new System.Windows.Forms.Panel();
			this.contextMenuStrip1.SuspendLayout();
			this.buttonPanel.SuspendLayout();
			((System.ComponentModel.ISupportInitialize)(this.splitContainer1)).BeginInit();
			this.splitContainer1.Panel1.SuspendLayout();
			this.splitContainer1.Panel2.SuspendLayout();
			this.splitContainer1.SuspendLayout();
			this.contextMenuStrip2.SuspendLayout();
			this.bottpmPanel.SuspendLayout();
			this.SuspendLayout();
			// 
			// contextMenuStrip1
			// 
			this.contextMenuStrip1.Items.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.コピーCToolStripMenuItem,
            this.検索SToolStripMenuItem});
			this.contextMenuStrip1.Name = "contextMenuStrip1";
			this.contextMenuStrip1.Size = new System.Drawing.Size(120, 48);
			// 
			// コピーCToolStripMenuItem
			// 
			this.コピーCToolStripMenuItem.Name = "コピーCToolStripMenuItem";
			this.コピーCToolStripMenuItem.Size = new System.Drawing.Size(152, 22);
			this.コピーCToolStripMenuItem.Text = "コピー(&C)";
			this.コピーCToolStripMenuItem.Click += new System.EventHandler(this.コピーCToolStripMenuItem_Click);
			// 
			// 検索SToolStripMenuItem
			// 
			this.検索SToolStripMenuItem.Name = "検索SToolStripMenuItem";
			this.検索SToolStripMenuItem.Size = new System.Drawing.Size(152, 22);
			this.検索SToolStripMenuItem.Text = "検索(&S)";
			this.検索SToolStripMenuItem.Click += new System.EventHandler(this.検索SToolStripMenuItem_Click);
			// 
			// buttonPanel
			// 
			this.buttonPanel.Controls.Add(this.macthCountButton);
			this.buttonPanel.Controls.Add(this.closeButton);
			this.buttonPanel.Dock = System.Windows.Forms.DockStyle.Right;
			this.buttonPanel.Location = new System.Drawing.Point(727, 0);
			this.buttonPanel.Name = "buttonPanel";
			this.buttonPanel.Size = new System.Drawing.Size(188, 83);
			this.buttonPanel.TabIndex = 0;
			// 
			// macthCountButton
			// 
			this.macthCountButton.Location = new System.Drawing.Point(3, 52);
			this.macthCountButton.Name = "macthCountButton";
			this.macthCountButton.Size = new System.Drawing.Size(82, 19);
			this.macthCountButton.TabIndex = 6;
			this.macthCountButton.Text = "一致ワード数";
			this.macthCountButton.UseVisualStyleBackColor = true;
			this.macthCountButton.Click += new System.EventHandler(this.macthCountButton_Click);
			// 
			// closeButton
			// 
			this.closeButton.Location = new System.Drawing.Point(91, 52);
			this.closeButton.Name = "closeButton";
			this.closeButton.Size = new System.Drawing.Size(82, 19);
			this.closeButton.TabIndex = 5;
			this.closeButton.Text = "閉じる";
			this.closeButton.UseVisualStyleBackColor = true;
			this.closeButton.Click += new System.EventHandler(this.closeButton_Click);
			// 
			// splitContainer1
			// 
			this.splitContainer1.Dock = System.Windows.Forms.DockStyle.Fill;
			this.splitContainer1.Location = new System.Drawing.Point(0, 0);
			this.splitContainer1.Name = "splitContainer1";
			// 
			// splitContainer1.Panel1
			// 
			this.splitContainer1.Panel1.Controls.Add(this.srcText);
			// 
			// splitContainer1.Panel2
			// 
			this.splitContainer1.Panel2.Controls.Add(this.targetText);
			this.splitContainer1.Size = new System.Drawing.Size(915, 369);
			this.splitContainer1.SplitterDistance = 305;
			this.splitContainer1.TabIndex = 5;
			// 
			// srcText
			// 
			this.srcText.ContextMenuStrip = this.contextMenuStrip1;
			this.srcText.Dock = System.Windows.Forms.DockStyle.Fill;
			this.srcText.Location = new System.Drawing.Point(0, 0);
			this.srcText.Name = "srcText";
			this.srcText.Size = new System.Drawing.Size(305, 369);
			this.srcText.TabIndex = 2;
			this.srcText.Text = "";
			// 
			// targetText
			// 
			this.targetText.ContextMenuStrip = this.contextMenuStrip2;
			this.targetText.Dock = System.Windows.Forms.DockStyle.Fill;
			this.targetText.Location = new System.Drawing.Point(0, 0);
			this.targetText.Name = "targetText";
			this.targetText.Size = new System.Drawing.Size(606, 369);
			this.targetText.TabIndex = 3;
			this.targetText.Text = "";
			// 
			// contextMenuStrip2
			// 
			this.contextMenuStrip2.Items.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.コピーCToolStripMenuItem1,
            this.検索SToolStripMenuItem1});
			this.contextMenuStrip2.Name = "contextMenuStrip2";
			this.contextMenuStrip2.Size = new System.Drawing.Size(153, 70);
			// 
			// コピーCToolStripMenuItem1
			// 
			this.コピーCToolStripMenuItem1.Name = "コピーCToolStripMenuItem1";
			this.コピーCToolStripMenuItem1.Size = new System.Drawing.Size(152, 22);
			this.コピーCToolStripMenuItem1.Text = "コピー(&C)";
			this.コピーCToolStripMenuItem1.Click += new System.EventHandler(this.コピーCToolStripMenuItem1_Click);
			// 
			// 検索SToolStripMenuItem1
			// 
			this.検索SToolStripMenuItem1.Name = "検索SToolStripMenuItem1";
			this.検索SToolStripMenuItem1.Size = new System.Drawing.Size(152, 22);
			this.検索SToolStripMenuItem1.Text = "検索(&S)";
			this.検索SToolStripMenuItem1.Click += new System.EventHandler(this.検索SToolStripMenuItem1_Click);
			// 
			// infoLabel
			// 
			this.infoLabel.AutoSize = true;
			this.infoLabel.Location = new System.Drawing.Point(12, 14);
			this.infoLabel.Name = "infoLabel";
			this.infoLabel.Size = new System.Drawing.Size(35, 12);
			this.infoLabel.TabIndex = 4;
			this.infoLabel.Text = "label1";
			// 
			// bottpmPanel
			// 
			this.bottpmPanel.Controls.Add(this.infoLabel);
			this.bottpmPanel.Controls.Add(this.buttonPanel);
			this.bottpmPanel.Dock = System.Windows.Forms.DockStyle.Bottom;
			this.bottpmPanel.Location = new System.Drawing.Point(0, 369);
			this.bottpmPanel.Name = "bottpmPanel";
			this.bottpmPanel.Size = new System.Drawing.Size(915, 83);
			this.bottpmPanel.TabIndex = 4;
			// 
			// CompareForm
			// 
			this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 12F);
			this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
			this.ClientSize = new System.Drawing.Size(915, 452);
			this.Controls.Add(this.splitContainer1);
			this.Controls.Add(this.bottpmPanel);
			this.Name = "CompareForm";
			this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
			this.Text = "文書比較";
			this.FormClosing += new System.Windows.Forms.FormClosingEventHandler(this.CompareForm_FormClosing);
			this.Load += new System.EventHandler(this.CompareForm_Load);
			this.SizeChanged += new System.EventHandler(this.CompareForm_SizeChanged);
			this.contextMenuStrip1.ResumeLayout(false);
			this.buttonPanel.ResumeLayout(false);
			this.splitContainer1.Panel1.ResumeLayout(false);
			this.splitContainer1.Panel2.ResumeLayout(false);
			((System.ComponentModel.ISupportInitialize)(this.splitContainer1)).EndInit();
			this.splitContainer1.ResumeLayout(false);
			this.contextMenuStrip2.ResumeLayout(false);
			this.bottpmPanel.ResumeLayout(false);
			this.bottpmPanel.PerformLayout();
			this.ResumeLayout(false);

		}

		#endregion

		private System.Windows.Forms.ContextMenuStrip contextMenuStrip1;
		private System.Windows.Forms.ToolStripMenuItem コピーCToolStripMenuItem;
		private System.Windows.Forms.ToolStripMenuItem 検索SToolStripMenuItem;
		private System.Windows.Forms.Panel buttonPanel;
		private System.Windows.Forms.Button macthCountButton;
		private System.Windows.Forms.Button closeButton;
		private System.Windows.Forms.SplitContainer splitContainer1;
		private System.Windows.Forms.RichTextBox srcText;
		private System.Windows.Forms.RichTextBox targetText;
		private System.Windows.Forms.ContextMenuStrip contextMenuStrip2;
		private System.Windows.Forms.ToolStripMenuItem コピーCToolStripMenuItem1;
		private System.Windows.Forms.ToolStripMenuItem 検索SToolStripMenuItem1;
		private System.Windows.Forms.Label infoLabel;
		private System.Windows.Forms.Panel bottpmPanel;
	}
}