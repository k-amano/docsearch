using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Diagnostics;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace Arx.DocSearch.ClientStarter
{
    public partial class MainForm : Form
    {
        public MainForm()
        {
            InitializeComponent();
        }

        private void startButton_Click(object sender, EventArgs e)
        {
            this.StartClientProgram();
            string fname = Path.Combine(Application.StartupPath, "client.tmp");
            while (File.Exists(fname)) { this.StartClientProgram(); }
        }

        private void StartClientProgram()
        {
            string pname = Path.Combine(Application.StartupPath, "Arx.DocSearch.Client.exe");
            if (File.Exists(pname))
            {
                using (Process p = Process.Start(pname)) {
                    while (true)
                    {
                        Thread.Sleep(3000);
                        if (p.HasExited) break;
                    }
                }

                Thread.Sleep(3000);
            }
            else
            {
                MessageBox.Show(string.Format("プログラム'{0}'が存在しません。", pname),
                    "エラー",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Error);
            }
        }
    }
}
