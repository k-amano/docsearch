namespace Arx.DocSearch.Client
{
	partial class MatchCountForm
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
			this.bottomPanel = new System.Windows.Forms.Panel();
			this.listView1 = new System.Windows.Forms.ListView();
			this.panel1 = new System.Windows.Forms.Panel();
			this.closeButton = new System.Windows.Forms.Button();
			this.infoLabel = new System.Windows.Forms.Label();
			this.bottomPanel.SuspendLayout();
			this.panel1.SuspendLayout();
			this.SuspendLayout();
			// 
			// bottomPanel
			// 
			this.bottomPanel.Controls.Add(this.infoLabel);
			this.bottomPanel.Controls.Add(this.panel1);
			this.bottomPanel.Dock = System.Windows.Forms.DockStyle.Bottom;
			this.bottomPanel.Location = new System.Drawing.Point(0, 262);
			this.bottomPanel.Name = "bottomPanel";
			this.bottomPanel.Size = new System.Drawing.Size(783, 40);
			this.bottomPanel.TabIndex = 0;
			// 
			// listView1
			// 
			this.listView1.Dock = System.Windows.Forms.DockStyle.Fill;
			this.listView1.Location = new System.Drawing.Point(0, 0);
			this.listView1.Name = "listView1";
			this.listView1.Size = new System.Drawing.Size(783, 262);
			this.listView1.TabIndex = 8;
			this.listView1.UseCompatibleStateImageBehavior = false;
			// 
			// panel1
			// 
			this.panel1.Controls.Add(this.closeButton);
			this.panel1.Dock = System.Windows.Forms.DockStyle.Right;
			this.panel1.Location = new System.Drawing.Point(678, 0);
			this.panel1.Name = "panel1";
			this.panel1.Size = new System.Drawing.Size(105, 40);
			this.panel1.TabIndex = 0;
			// 
			// closeButton
			// 
			this.closeButton.Location = new System.Drawing.Point(11, 9);
			this.closeButton.Name = "closeButton";
			this.closeButton.Size = new System.Drawing.Size(82, 19);
			this.closeButton.TabIndex = 9;
			this.closeButton.Text = "閉じる";
			this.closeButton.UseVisualStyleBackColor = true;
			this.closeButton.Click += new System.EventHandler(this.closeButton_Click);
			// 
			// infoLabel
			// 
			this.infoLabel.AutoSize = true;
			this.infoLabel.Location = new System.Drawing.Point(12, 12);
			this.infoLabel.Name = "infoLabel";
			this.infoLabel.Size = new System.Drawing.Size(35, 12);
			this.infoLabel.TabIndex = 10;
			this.infoLabel.Text = "label1";
			// 
			// MatchCountForm
			// 
			this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 12F);
			this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
			this.ClientSize = new System.Drawing.Size(783, 302);
			this.Controls.Add(this.listView1);
			this.Controls.Add(this.bottomPanel);
			this.Name = "MatchCountForm";
			this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
			this.Text = "一致ワード数一覧";
			this.FormClosing += new System.Windows.Forms.FormClosingEventHandler(this.MatchCountForm_FormClosing);
			this.Load += new System.EventHandler(this.MatchCountForm_Load);
			this.bottomPanel.ResumeLayout(false);
			this.bottomPanel.PerformLayout();
			this.panel1.ResumeLayout(false);
			this.ResumeLayout(false);

		}

		#endregion

		private System.Windows.Forms.Panel bottomPanel;
		private System.Windows.Forms.ListView listView1;
		private System.Windows.Forms.Label infoLabel;
		private System.Windows.Forms.Panel panel1;
		private System.Windows.Forms.Button closeButton;

	}
}