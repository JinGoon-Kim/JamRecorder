using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Text;
using System.Windows.Forms;

using System.Threading;
using System.IO;
using System.Diagnostics;
using Microsoft.Win32; // registry 
using System.Runtime.InteropServices;
using System.Security.Principal;

using System.Management;



namespace JamRecorder
{



    public partial class Form1 : Form
    {

        Process process;
        Process cmd;
        string ffmpeg_filename = System.Environment.CurrentDirectory + "\\ffmpeg.exe";
        string cmd_filename = "cmd.exe";

        string work_directory = System.Environment.CurrentDirectory;



        // ---------------------------------------------
        // UTF8 변환 
        // ---------------------------------------------
        public static string ToUTF8(string text)
        {
            return Encoding.UTF8.GetString(Encoding.Default.GetBytes(text));
        }


        // ---------------------------------------------
        // audio device 찾기
        // ---------------------------------------------
        [DllImport("Winmm.dll", SetLastError = true, CharSet = CharSet.Auto)]
        public static extern uint waveOutGetNumDevs();

        [DllImport("Winmm.dll", SetLastError = true, CharSet = CharSet.Auto)]
        public static extern int waveOutGetDevCaps(IntPtr uDeviceID, out WAVEOUTCAPS pwoc, int cbwoc);

        [StructLayout(LayoutKind.Sequential, CharSet = CharSet.Auto)]
        public struct WAVEOUTCAPS
        {
            public ushort wMid;
            public ushort wPid;
            public uint vDriverVersion;
            [MarshalAs(UnmanagedType.ByValTStr, SizeConst = 32)]
            public string szPname;
            public int dwFormats;
            public ushort wChannels;
            public ushort wReserved1;
            public int dwSupport;
        }

        [DllImport("Winmm.dll", SetLastError = true, CharSet = CharSet.Auto)]
        public static extern int waveOutMessage(IntPtr deviceID, uint uMsg, out uint dwParam1, out uint dwParam2);

        public const int WAVE_MAPPER = (-1);
        public const int MMSYSERR_BASE = 0;
        /* general error return values */
        public const int MMSYSERR_NOERROR = 0;                    /* no error */
        public const int MMSYSERR_ERROR = (MMSYSERR_BASE + 1);  /* unspecified error */
        public const int MMSYSERR_BADDEVICEID = (MMSYSERR_BASE + 2);  /* device ID out of range */
        public const int MMSYSERR_NOTENABLED = (MMSYSERR_BASE + 3);  /* driver failed enable */
        public const int MMSYSERR_ALLOCATED = (MMSYSERR_BASE + 4);  /* device already allocated */
        public const int MMSYSERR_INVALHANDLE = (MMSYSERR_BASE + 5);  /* device handle is invalid */
        public const int MMSYSERR_NODRIVER = (MMSYSERR_BASE + 6);  /* no device driver present */
        public const int MMSYSERR_NOMEM = (MMSYSERR_BASE + 7);  /* memory allocation error */
        public const int MMSYSERR_NOTSUPPORTED = (MMSYSERR_BASE + 8);  /* function isn't supported */
        public const int MMSYSERR_BADERRNUM = (MMSYSERR_BASE + 9);  /* error value out of range */
        public const int MMSYSERR_INVALFLAG = (MMSYSERR_BASE + 10); /* invalid flag passed */
        public const int MMSYSERR_INVALPARAM = (MMSYSERR_BASE + 11); /* invalid parameter passed */
        public const int MMSYSERR_HANDLEBUSY = (MMSYSERR_BASE + 12); /* handle being used simultaneously on another thread (eg callback) */
        public const int MMSYSERR_INVALIDALIAS = (MMSYSERR_BASE + 13); /* specified alias not found */
        public const int MMSYSERR_BADDB = (MMSYSERR_BASE + 14); /* bad registry database */
        public const int MMSYSERR_KEYNOTFOUND = (MMSYSERR_BASE + 15); /* registry key not found */
        public const int MMSYSERR_READERROR = (MMSYSERR_BASE + 16); /* registry read error */
        public const int MMSYSERR_WRITEERROR = (MMSYSERR_BASE + 17); /* registry write error */
        public const int MMSYSERR_DELETEERROR = (MMSYSERR_BASE + 18); /* registry delete error */
        public const int MMSYSERR_VALNOTFOUND = (MMSYSERR_BASE + 19); /* registry value not found */
        public const int MMSYSERR_NODRIVERCB = (MMSYSERR_BASE + 20); /* driver does not call DriverCallback */
        public const int MMSYSERR_MOREDATA = (MMSYSERR_BASE + 21); /* more data to be returned */
        public const int MMSYSERR_LASTERROR = (MMSYSERR_BASE + 21); /* last error in range */

        public const int DRV_RESERVED = 0x0800;

        public const int DRVM_MAPPER = (0x2000);
        public const int DRVM_USER = 0x4000;
        public const int DRVM_MAPPER_STATUS = (DRVM_MAPPER + 0);
        public const int DRVM_MAPPER_RECONFIGURE = (DRVM_MAPPER + 1);
        public const int DRVM_MAPPER_PREFERRED_GET = (DRVM_MAPPER + 21);
        public const int DRVM_MAPPER_CONSOLEVOICECOM_GET = (DRVM_MAPPER + 23);

        public const int DRV_QUERYDEVNODE = (DRV_RESERVED + 2);
        public const int DRV_QUERYMAPPABLE = (DRV_RESERVED + 5);
        public const int DRV_QUERYMODULE = (DRV_RESERVED + 9);
        public const int DRV_PNPINSTALL = (DRV_RESERVED + 11);
        public const int DRV_QUERYDEVICEINTERFACE = (DRV_RESERVED + 12);
        public const int DRV_QUERYDEVICEINTERFACESIZE = (DRV_RESERVED + 13);
        public const int DRV_QUERYSTRINGID = (DRV_RESERVED + 14);
        public const int DRV_QUERYSTRINGIDSIZE = (DRV_RESERVED + 15);
        public const int DRV_QUERYIDFROMSTRINGID = (DRV_RESERVED + 16);
        public const int DRV_QUERYFUNCTIONINSTANCEID = (DRV_RESERVED + 17);
        public const int DRV_QUERYFUNCTIONINSTANCEIDSIZE = (DRV_RESERVED + 18);
        public const int DRVM_MAPPER_PREFERRED_FLAGS_PREFERREDONLY = 0x00000001;
        // ---------------------------------------------
        // ctrl + c 기능 
        // ---------------------------------------------

        [DllImport("user32.dll")]
        private static extern int SetActiveWindow(int hwnd);

        [DllImport("kernel32.dll")]
        static extern bool GenerateConsoleCtrlEvent(int dwCtrlEvent, int dwProcessGroupId);

        [DllImport("kernel32.dll")]
        static extern bool SetConsoleCtrlHandler(IntPtr handlerRoutine, bool add);

        [DllImport("kernel32.dll")]
        static extern bool AttachConsole(int dwProcessId);

        [DllImport("kernel32.dll")]
        static extern bool FreeConsole();
        // ---------------------------------------------
        // DLL 체크 
        // ---------------------------------------------
        [DllImport("kernel32")]
        public extern static bool FreeLibrary(int hLibModule);

        [DllImport("kernel32")]
        public extern static int LoadLibrary(string lpLibFileName);



        // ---------------------------------------------
        //  ffmpeg 다 죽이기
        // ---------------------------------------------
        static void KillAllFFMPEG()
        {
            Process KillFFMPEG = new Process();
            ProcessStartInfo taskkillStartInfo = new ProcessStartInfo
            {
                FileName = "taskkill",
                Arguments = "/F /IM ffmpeg.exe",
                UseShellExecute = false,
                CreateNoWindow = true
            };

            KillFFMPEG.StartInfo = taskkillStartInfo;
            KillFFMPEG.Start();
        }

        // ---------------------------------------------
        // 기존 실행되어 있는 ffmpeg 를 ctrl + c 로 종료 
        // ---------------------------------------------
        static void stopFFMPEG()
        {
            Process[] processList;
            processList = Process.GetProcessesByName("ffmpeg");
            if (processList.Length > 0)
            {
                for (int i = 0; i < processList.Length; i++)
                {
                    AttachConsole(processList[i].Id);
                    SetConsoleCtrlHandler(IntPtr.Zero, true);
                    GenerateConsoleCtrlEvent(0, 0);
                    FreeConsole();
                }
            }
            Thread.Sleep(100);

        }

        // ---------------------------------------------
        // 레지스트리에 프로토콜 등록 , JamRecorder://record 
        // ---------------------------------------------
        private static void AddReg()
        {
            RegistryKey key;
            key = Registry.ClassesRoot.CreateSubKey("JamRecorder");
            key.SetValue("", "URL: JamRecorder Protocol");
            key.SetValue("URL Protocol", "");

            key = key.CreateSubKey("shell");
            key = key.CreateSubKey("open");
            key = key.CreateSubKey("command");
            key.SetValue("", Application.ExecutablePath + " %1");
        }

        // ---------------------------------------------
        // DLL 등록 확인 
        // ---------------------------------------------
        public bool IsDllRegistered(string DllName)
        {
            int libId = LoadLibrary(DllName);
            if (libId > 0) FreeLibrary(libId);
            return (libId > 0);
        }

        // ---------------------------------------------
        // 레지스트리 키 체크 
        // ---------------------------------------------
        public bool IsKeyRegistered()
        {
            RegistryKey key;
            key = Registry.ClassesRoot.OpenSubKey("JamRecorder");
            if (key == null)
            {
                return false;
            }
            else
            {
                return true;
            }
        }

        // ---------------------------------------------
        // regsvr32.exe를 이용하여 DLL 등록
        // ---------------------------------------------
        public static void registerDLL(string dllPath, bool isRegist)
        {
            try
            {
                //'/s' : indicates regsvr32.exe to run silently.
                string fileinfo = (isRegist ? "/s" : "/u /s") + " " + "\"" + dllPath + "\"";
                Process reg = new Process();
                reg.StartInfo.FileName = "regsvr32.exe";
                reg.StartInfo.Arguments = fileinfo;
                reg.StartInfo.UseShellExecute = false;
                reg.StartInfo.CreateNoWindow = true;
                reg.StartInfo.RedirectStandardOutput = true;
                reg.Start();
                reg.WaitForExit();
                reg.Close();
            }
            catch (Exception ex)
            {
                MessageBox.Show("DLL register ERROR! :" + ex.Message);
            }
        }

        // ---------------------------------------------
        // ffmpeg의 프로세스 ID 가지고, ctrl + c 로 ffmpeg 종료 
        // ---------------------------------------------
        private void StopRecord()
        {
            AttachConsole(cmd.Id);
            SetConsoleCtrlHandler(IntPtr.Zero, true);
            GenerateConsoleCtrlEvent(0, 0);
            FreeConsole();
            System.Diagnostics.Process.Start("explorer.exe", work_directory + "\\record");
        }

        // ---------------------------------------------
        // 관리자권한 체크 
        // ---------------------------------------------
        public static bool IsAdministrator()
        {
            WindowsIdentity identity = WindowsIdentity.GetCurrent();

            if (null != identity)
            {
                WindowsPrincipal principal = new WindowsPrincipal(identity);
                return principal.IsInRole(WindowsBuiltInRole.Administrator);
            }

            return false;
        }
        // ---------------------------------------------
        // 작업영역 사이즈 가져오기 
        // ---------------------------------------------
        public string getWorkingAreaSize()
        {
            int W = Screen.PrimaryScreen.WorkingArea.Width; //작업영역 가로크기
            int H = Screen.PrimaryScreen.WorkingArea.Height; // 작업영역 세로크기
            string size = W.ToString() + "x" + H.ToString();

            return size;
        }

        // ---------------------------------------------
        // 초기화 : 숨김 처리 후 관리자권한 실행 
        // ---------------------------------------------
        public void init() 
        {           
            if (IsAdministrator() == false)
            {
                try
                {
                    process = new System.Diagnostics.Process();
                    process.StartInfo.UseShellExecute = false;
                    process.StartInfo.FileName = Application.ExecutablePath;
                    process.StartInfo.WorkingDirectory = work_directory;
                    process.StartInfo.Verb = "runas";
                    process.Start();
                }
                catch (Exception ex)
                {
                    MessageBox.Show(ex.Message.ToString());
                }

                return;
            }
        }




        // ---------------------------------------------
        // ffmpeg 실행 
        // ---------------------------------------------
        private void StartRecord()
        {

            // record 디렉토리 없으면 생성 
            DirectoryInfo Di = new DirectoryInfo(work_directory + "\\record");
            if (Di.Exists == false) Di.Create();

            cmd = new Process();
            string out_filename = DateTime.Now.ToString("d") + "_" + DateTime.Now.ToString("HHmmss") + ".mp4";
            string size = getWorkingAreaSize();
            string size_option = " -offset_x 0 -offset_y 0 -video_size " + size + " ";

            // 오디오 옵션 변경 
            string audio_option = " -f dshow -i audio=virtual-audio-capturer ";
            audio_option = " -f dshow -i audio=\"" + getAudioDevice() + "\" ";
            audio_option = " -f dshow -i audio=\"" + getAudioDevice() + "\" ";


            string chcp = "/C chcp 850 > null &&   " + ffmpeg_filename;

            string option = chcp + audio_option  + " -f gdigrab " + size_option + "  -i desktop  -framerate 30  -probesize 40M -preset ultrafast -pix_fmt yuv420p record\\" + out_filename;

            cmd.StartInfo.FileName = cmd_filename;
            cmd.StartInfo.WorkingDirectory = work_directory; // The output directory  
            cmd.StartInfo.CreateNoWindow = true;
            cmd.StartInfo.UseShellExecute = false;
            cmd.StartInfo.Arguments = option;

            // 로그 저장 
            using (StreamWriter outputFile = new StreamWriter(@"log.txt", true))
            {
                outputFile.WriteLine(cmd.StartInfo.Arguments);
            }

            try
            {
                cmd.Start();
            }
            catch (Exception e)
            {
                MessageBox.Show(e.Message);
            }

            Thread.Sleep(500);
        }



        public string getAudioDevice()
        {

            ProcessStartInfo startInfo = new ProcessStartInfo();
            startInfo.CreateNoWindow = true;
            startInfo.UseShellExecute = false;
            startInfo.FileName = ffmpeg_filename; 
            startInfo.WindowStyle = ProcessWindowStyle.Hidden;
            startInfo.Arguments = "    -hide_banner -list_devices true -f dshow -i dummy ";

            startInfo.StandardErrorEncoding = Encoding.UTF8;
            startInfo.RedirectStandardOutput = true;
            startInfo.RedirectStandardError = true;

            using (Process exeProcess = Process.Start(startInfo))
            {
                string output = exeProcess.StandardError.ReadToEnd();
                exeProcess.WaitForExit();

                  MessageBox.Show(output);

                // 엔터값 기준 배열 나누기 
                string[] tmp = output.Split('\n');


                // DirectShow audio devices 문자열 찾기 
                int i = 0;
                int idx;
                string tmp2 = "";

                string result = "";

                foreach (string str in tmp)
                {
                    i++;
                    idx = str.IndexOf("DirectShow audio devices");
                    if (idx > -1)
                    {
                        tmp2 = tmp[i];
                    }
                }
                if (tmp2 != "")
                {
                    tmp = tmp2.Split('"');
                    result = tmp[1];
                }

                // MessageBox.Show(result); 
              //  result = "speaker(Realtek High Definition Audio)";


                return result;
               
            }

        }



        public void test() {
    
            string option = "ffmpeg -list_devices true -f dshow -i dummy ";

            ProcessStartInfo cmd = new ProcessStartInfo();
            Process process = new Process();
            cmd.FileName = @"cmd";
            cmd.WindowStyle = ProcessWindowStyle.Hidden;             // cmd창이 숨겨지도록 하기
            cmd.CreateNoWindow = true;                               // cmd창을 띄우지 안도록 하기

            cmd.UseShellExecute = false;
            cmd.RedirectStandardOutput = true;        // cmd창에서 데이터를 가져오기
            cmd.RedirectStandardInput = true;          // cmd창으로 데이터 보내기
            cmd.RedirectStandardError = true;          // cmd창에서 오류 내용 가져오기

            process.EnableRaisingEvents = false;
            process.StartInfo = cmd;
            process.Start();
            process.StandardInput.Write(option + Environment.NewLine);
            process.StandardInput.Close();

            string result = process.StandardOutput.ReadToEnd();
            StringBuilder sb = new StringBuilder();
            sb.Append("[Result Info]" + DateTime.Now + "\r\n");
            sb.Append(result);
            sb.Append("\r\n");
            MessageBox.Show(sb.ToString()); 
            process.WaitForExit();
            process.Close();
        }





        public Form1()
        {
            InitializeComponent();
            this.WindowState = FormWindowState.Minimized;
            this.ShowInTaskbar = false;
            this.Visible = false;
            init(); 
        }

        private void Form1_Load(object sender, EventArgs e)
        {
            bool is_stop = false;

            string[] args = Environment.GetCommandLineArgs();
            Process[] processList;

            if (args.Length == 2)
            {
                string[] tmp = args[1].Split(':');
                if (tmp.Length == 2)
                {
                    // jamracorder:stop  크롬에서 클릭시 
                    if (tmp[1] == "stop")
                    {
                        is_stop = true;
                        stopFFMPEG();
                        Application.Exit();
                    }

                }
            }

            // 중복 실행 확인 
            System.Diagnostics.Process[] myProc = System.Diagnostics.Process.GetProcessesByName("JamRecorder");
            if (myProc.Length > 1)
            {
                MessageBox.Show("이미 녹화중입니다.");
                Application.Exit();
            }
            else if (is_stop == false)
            {

                // ffmpeg 중복 실행 방지 - 이전것 닫고 새로 실행 
                processList = Process.GetProcessesByName("ffmpeg");
                if (processList.Length > 0)
                {
                    for (int i = 0; i < processList.Length; i++)
                    {
                        AttachConsole(processList[i].Id);
                        SetConsoleCtrlHandler(IntPtr.Zero, true);
                        GenerateConsoleCtrlEvent(0, 0);
                        FreeConsole();
                        Thread.Sleep(1100);
                    }
                }


                // 레지스트리 등록 
                if (IsKeyRegistered() == false)
                {
                    AddReg();
                    registerDLL("screen-capture-recorder.dll", true);
                }
                // 녹화 시작 
                StartRecord();

                // 체크 타이머 작동 
                timer1.Enabled = true;
                timer1.Start();
            }
        }


        private void timer1_Tick(object sender, EventArgs e)
        {
            // ffmpeg 실행 안되면 무조건 종료 
            Process[] processList = Process.GetProcessesByName("ffmpeg");
            if (processList.Length == 0)
            {
                Application.Exit();
            }
        }

        private void contextMenuStrip1_ItemClicked(object sender, ToolStripItemClickedEventArgs e)
        {
            string item;

            item = e.ClickedItem.Text;
            switch (item)
            {
                case "Exit":
                    StopRecord();
                    Application.Exit();
                    break;
            }
        }

        private void toolStripMenuItem1_Click(object sender, EventArgs e)
        {
          
        }
    }
}
