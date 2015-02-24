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
            // if you want to start anything else (lnk file, etc) this will only work without credentials
            // to circumvent this the mshta exe is called with vbscript code as parameter which uses a simple wscript.run
            // this horrible workaround is the only true way to call arbitrary files without parsing them for ages
            // otherwise there has to be a handler for lnks, a special one for msi installer lnks and a third one for file
            // this only works with internet explorer installed
            //code for testing in cmd: mshta vbscript:Execute("CreateObject (""WScript.Shell"").Run ""cmd.exe"", ,False:Close")
            Process proc = new Process {
                StartInfo = {
                    UseShellExecute = false,
                    FileName = "mshta.exe",
                    UserName = userName,
                    Domain = "tilak",
                    Password = GetSecureString(userPwd),
                    WorkingDirectory = "C:\\",
                    // the arguments code needs 6 pairs of escaped " the reason is 1 pair because it is a string, 1 additional pair because the path could include spaces,
                    // that 2nd pair will be interpreted in vbscript, where this is is escaped with an additional "
                    // those 3 " will all be parsed by mshta which ALSO escapes all " with an additional " so they are doubled to 6
                    // finally they have to escaped with \ because of c# syntax
                    // Before the File is executed Z: is always mapped to allow calls with "default shortcuts" we often use
                    // Multiline vbscript commands in a single line are seperated by :
                    Arguments = "vbscript:Execute(\"CreateObject(\"\"WScript.Network\"\").MapNetworkDrive \"\"Z:\"\", \"\"\\\\tilak\\share\\appl\"\", False:CreateObject (\"\"WScript.Shell\"\").Run \"\"\"\"\"\"" + file + "\"\"\"\"\"\", ,False:Close\")"
                    // Arguments = "vbscript:Execute(\"CreateObject (\"\"WScript.Shell\"\").Run \"\"\"\"\"\"" + file + "\"\"\"\"\"\", ,False:Close\")"
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
