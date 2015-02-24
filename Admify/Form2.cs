using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Runtime.InteropServices;
using System.Diagnostics;
using System.Security;
using System.ServiceModel;

namespace Admify
{
    // interface for wcf pipe client server communication
    [ServiceContract]
    public interface IAdmifyExecutor {
        [OperationContract]
        void executeFile(string file);
    }

    public partial class Form2 : Form, IAdmifyExecutor
    {
        private static string userName;
        private static string userPwd;
        private static Form parentForm;
        private ServiceHost host;

        [DllImport("user32.dll", SetLastError = true, CharSet = CharSet.Auto)]
        public static extern int SendMessage(IntPtr hWnd, int Msg, int wParam, int lParam);
        [DllImport("user32.dll", SetLastError = true, CharSet = CharSet.Auto)]
        public static extern bool ReleaseCapture();
        [DllImport("user32.dll", SetLastError = true, CharSet = CharSet.Auto)]
        public static extern bool SetForegroundWindow(IntPtr hWnd);
        [DllImport("user32.dll", SetLastError = true, CharSet = CharSet.Auto)]
        public static extern int SetWindowLong(IntPtr hwnd, int index, int value);
        [DllImport("user32.dll", SetLastError = true, CharSet = CharSet.Auto)]
        public static extern int GetWindowLong(IntPtr hwnd, int index);

        // this is needed because the class doubles as the wcf server (pretty horrible)
        public Form2() {

        }

        public Form2(string user, string pwd, Form pForm)
        {
            userName = user;
            userPwd = pwd;
            parentForm = pForm;
            InitializeComponent();
        }

        // on load the minimize and maximize button are removed (even if invisible they otherwise dictate the minimum size)
        // and the window is set to be always on top
        // size is set to 40x40 pixels
        private void Form2_Load(object sender, EventArgs e)
        {
	        //SetForegroundWindow(this.Handle)
	        // GWL_STYLE = -16
	        int value = GetWindowLong(this.Handle, -16);
	        SetWindowLong(this.Handle, -16, value & -131073 & -65537);
            this.Size = new System.Drawing.Size(40, 40);
            this.DesktopLocation = new Point(0, 400);

            // create a wcf server that listens ot pipes
            host = new ServiceHost(typeof(Form2), new Uri[] { new Uri("net.pipe://localhost") });
            host.AddServiceEndpoint(typeof(IAdmifyExecutor), new NetNamedPipeBinding(), "PipeExecute");
            host.Open();
        }

        private void Form2_MouseDown(object sender, MouseEventArgs e)
        {
            if (e.Button == MouseButtons.Left)
            {
                ReleaseCapture();
                SendMessage(this.Handle, 0xA1, 0x2, 0);
            }
        }

        private void Form2_FormClosed(object sender, FormClosedEventArgs e)
        {
            parentForm.Close();

            // close the wcf server
            host.Close();
        }

        private void Form2_DragEnter(object sender, DragEventArgs e)
        {
            e.Effect = DragDropEffects.Copy;
        }

        private void Form2_DragDrop(object sender, DragEventArgs e)
        {
            foreach (string file in (string[])e.Data.GetData(DataFormats.FileDrop))
            {
                executeFile(file);
            }
        }

        public void executeFile(string file) {
            // The process is started with the received credentials
            // this works fine as along as an executable is started
            // if you want to start anything else (lnk file, etc) this will only work without credentials because using ShellExecute only works without them
            // to circumvent this I wrote a special wrapper in C++ that takes command line arguments and executes them. It also maps Z:
            // otherwise there has to be a handler for lnks, a special one for msi installer lnks and a third one for file

            Process proc = new Process {
                StartInfo = {
                    UseShellExecute = false,
                    FileName = @"C:\Program Files (x86)\Admify\ShellExecuteProxy.exe",
                    UserName = userName,
                    Domain = "tilak",
                    Password = GetSecureString(userPwd),

                    WorkingDirectory = @"C:\",                   

                    // Arguments is only a string so it needs escaped " in case of spaces
                    Arguments = "\"" + file  + "\""
                }
            };
            proc.Start();
        }

        private static SecureString GetSecureString(string str)
        {
            SecureString secureString = new SecureString();
            foreach (char ch in str)
            {
                secureString.AppendChar(ch);
            }
            secureString.MakeReadOnly();
            return secureString;
        }
    }
}
