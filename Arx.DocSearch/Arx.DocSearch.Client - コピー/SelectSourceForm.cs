using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.IO;
using System.Text;
using System.Windows.Forms;
using System.Diagnostics;

namespace Arx.DocSearch.Client
{
	public partial class SelectSourceForm : Form
	{
		public SelectSourceForm(MainForm mainForm)
		{
			this.mainForm = mainForm;
			InitializeComponent();
		}
		private MainForm mainForm;
		private List<string> srcFiles;

		private void SelectSourceForm_Load(object sender, EventArgs e)
		{
			this.srcFiles = this.mainForm.SrcFiles;
			foreach (string fileName in this.srcFiles)
			{
				if (0 < this.FileListText.Text.Length)
				{
					this.FileListText.AppendText(Environment.NewLine);
				}
				this.FileListText.AppendText(fileName);
			}
		}

		private void SelectSourceForm_FormClosing(object sender, FormClosingEventArgs e)
		{
			this.srcFiles.Clear();
			foreach (string line in this.FileListText.Lines)
			{
				if (File.Exists(line)) this.srcFiles.Add(line);
				Debug.WriteLine(line);
			}
			this.mainForm.SrcFiles = this.srcFiles;
		}

		private void selectButton_Click(object sender, EventArgs e)
		{
			if (this.openFileDialog1.ShowDialog() == DialogResult.OK)
			{
				if (0 < this.FileListText.Text.Length)
				{
					this.FileListText.AppendText(Environment.NewLine);
				}
				this.FileListText.AppendText(this.openFileDialog1.FileName);
			}
		}

		private void clearButton_Click(object sender, EventArgs e)
		{
			this.FileListText.Clear();
			this.srcFiles.Clear();
		}

		private void closeButton_Click(object sender, EventArgs e)
		{
			this.Close();
		}

		private void FileListText_DragEnter(object sender, DragEventArgs e)
		{
			//ファイルがドラッグされている場合、カーソルを変更する
			if (e.Data.GetDataPresent(DataFormats.FileDrop))
			{
				e.Effect = DragDropEffects.Copy;
			}
		}

		private void FileListText_DragDrop(object sender, DragEventArgs e)
		{
			Debug.WriteLine("FileListText_DragDrop");
			//ドロップされたファイルの一覧を取得
			string[] fileNames = (string[])e.Data.GetData(DataFormats.FileDrop, false);
			if (fileNames.Length <= 0)
			{
				return;
			}
			// ドロップ先がTextBoxであるかチェック
			TextBox txtTarget = sender as TextBox;
			if (txtTarget == null)
			{
				return;
			}
			//ファイル名をTextBoxに追加する
			foreach (string fileName in fileNames)
			{
				if (this.srcFiles.Contains(fileName)) continue;
				if (0 < txtTarget.Text.Length)
				{
					txtTarget.AppendText(Environment.NewLine);
				}
				txtTarget.AppendText(fileName);
				this.srcFiles.Add(fileName);
			}
		}
	}
}
