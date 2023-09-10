namespace Arx.DocSearch.Client
{
	partial class SelectSourceForm
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
			this.closeButton = new System.Windows.Forms.Button();
			this.clearButton = new System.Windows.Forms.Button();
			this.selectButton = new System.Windows.Forms.Button();
			this.openFileDialog1 = new System.Windows.Forms.OpenFileDialog();
			this.FileListText = new System.Windows.Forms.TextBox();
			this.SuspendLayout();
			// 
			// closeButton
			// 
			this.closeButton.Location = new System.Drawing.Point(452, 297);
			this.closeButton.Name = "closeButton";
			this.closeButton.Size = new System.Drawing.Size(75, 23);
			this.closeButton.TabIndex = 7;
			this.closeButton.Text = "閉じる";
			this.closeButton.UseVisualStyleBackColor = true;
			this.closeButton.Click += new System.EventHandler(this.closeButton_Click);
			// 
			// clearButton
			// 
			this.clearButton.Location = new System.Drawing.Point(371, 297);
			this.clearButton.Name = "clearButton";
			this.clearButton.Size = new System.Drawing.Size(75, 23);
			this.clearButton.TabIndex = 6;
			this.clearButton.Text = "クリア";
			this.clearButton.UseVisualStyleBackColor = true;
			this.clearButton.Click += new System.EventHandler(this.clearButton_Click);
			// 
			// selectButton
			// 
			this.selectButton.Location = new System.Drawing.Point(290, 297);
			this.selectButton.Name = "selectButton";
			this.selectButton.Size = new System.Drawing.Size(75, 23);
			this.selectButton.TabIndex = 5;
			this.selectButton.Text = "選択";
			this.selectButton.UseVisualStyleBackColor = true;
			this.selectButton.Click += new System.EventHandler(this.selectButton_Click);
			// 
			// openFileDialog1
			// 
			this.openFileDialog1.FileName = "openFileDialog1";
			// 
			// FileListText
			// 
			this.FileListText.AllowDrop = true;
			this.FileListText.Location = new System.Drawing.Point(12, 12);
			this.FileListText.Multiline = true;
			this.FileListText.Name = "FileListText";
			this.FileListText.ReadOnly = true;
			this.FileListText.ScrollBars = System.Windows.Forms.ScrollBars.Vertical;
			this.FileListText.Size = new System.Drawing.Size(515, 276);
			this.FileListText.TabIndex = 4;
			this.FileListText.DragDrop += new System.Windows.Forms.DragEventHandler(this.FileListText_DragDrop);
			this.FileListText.DragEnter += new System.Windows.Forms.DragEventHandler(this.FileListText_DragEnter);
			// 
			// SelectSourceForm
			// 
			this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 12F);
			this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
			this.ClientSize = new System.Drawing.Size(539, 332);
			this.Controls.Add(this.closeButton);
			this.Controls.Add(this.clearButton);
			this.Controls.Add(this.selectButton);
			this.Controls.Add(this.FileListText);
			this.Name = "SelectSourceForm";
			this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
			this.Text = "検索テキストの選択";
			this.FormClosing += new System.Windows.Forms.FormClosingEventHandler(this.SelectSourceForm_FormClosing);
			this.Load += new System.EventHandler(this.SelectSourceForm_Load);
			this.ResumeLayout(false);
			this.PerformLayout();

		}

		#endregion

		private System.Windows.Forms.Button closeButton;
		private System.Windows.Forms.Button clearButton;
		private System.Windows.Forms.Button selectButton;
		private System.Windows.Forms.OpenFileDialog openFileDialog1;
		private System.Windows.Forms.TextBox FileListText;
	}
}