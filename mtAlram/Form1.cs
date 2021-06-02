using Newtonsoft.Json.Linq;
using OpenCvSharp;
using OpenCvSharp.Extensions;
using System;
using System.Diagnostics;
using System.Drawing;
using System.IO;
using System.Net;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using System.Windows.Forms;
using Telegram.Bot;
using Telegram.Bot.Types.Enums;

namespace mtAlram
{
    public partial class Form1 : MetroFramework.Forms.MetroForm
    {
        public Form1()
        {
            InitializeComponent();
            loadKeyValue();
            Bot = new TelegramBotClient(_token.ToString());
        }

        public const int WM_LBUTTONDOWN = 0x0201;
        public const int WM_LBUTTONUP = 0x0202;
        public const int WM_RBUTTONDOWN = 0x0204;
        public const int WM_RBUTTONUP = 0x0205;
        public const int WM_RBUTTONCLICK = 0x0206;
        public const int BM_CLICK = 0x00F5;
        public const int WM_KEYDOWN = 0x100;
        public const int WM_KEYUP = 0x101;
        public const int WM_CHAR = 0x105;
        public const int delayValue = 1000;
        private static StringBuilder _delayValue = new StringBuilder();
        private static StringBuilder _token = new StringBuilder();
        private static StringBuilder _chatId = new StringBuilder();
        private static StringBuilder _ucWindow = new StringBuilder();
        private static StringBuilder _kakaoWindow = new StringBuilder();

        public TelegramBotClient Bot;
        public bool checkprocess;

        [DllImport("User32", EntryPoint = "FindWindow")]
        private static extern IntPtr FindWindow(string lpClassName, string lpWindowName);

        [DllImport("user32.dll")]
        internal static extern bool PrintWindow(IntPtr hWnd, IntPtr hdcBlt, int nFlags);

        [DllImport("user32.dll")]
        public static extern int SendMessage(IntPtr hwnd, int wMsg, int wParam, String lParam);

        [DllImport("user32.dll")]
        public static extern int PostMessage(IntPtr hwnd, int wMsg, int wParam, int lParam);

        [DllImport("User32.dll")]
        public static extern IntPtr FindWindowEx(IntPtr Parent, IntPtr Child, string lpszClass, string lpszWindows);

        [DllImport("User32.dll")]
        private static extern bool SetForegroundWindow(IntPtr hWnd);

        [DllImport("user32.dll")]
        private static extern bool ShowWindowAsync(IntPtr hWnd, int nCmdShow);

        // INI 관련
        [DllImport("kernel32")]
        private static extern long WritePrivateProfileString(string section, string key, string val, string filePath);

        [DllImport("kernel32")]
        private static extern int GetPrivateProfileString(string section, string key, string def, StringBuilder retVal, int size, string filePath);


        private void Form1_Load(object sender, EventArgs e)
        {
            
        }

        //private async void telegramAPIAsync() => Console.WriteLine("Hello my name is {0}", (object)(await new TelegramBotClient(setToken).GetMeAsync(new CancellationToken())).FirstName);

        private static void loadKeyValue()
        {
            // 현재 프로그램 실행 위치
            string path = System.Reflection.Assembly.GetExecutingAssembly().Location;
            path = Path.GetDirectoryName(path) + "\\keyvalue.ini";

            // 쓰기 - long WritePrivateProfileString(string section, string key, string val, string filePath);
            //WritePrivateProfileString("SECTION", "TOKEN", "myToken", path);
            //WritePrivateProfileString("SECTION", "CHAT_ID", "myChatId", path);

            // 읽기 - int GetPrivateProfileString(string section, string key, string def, StringBuilder retVal, int size, string filePath);
            int result1 = GetPrivateProfileString("SECTION", "TOKEN", "def", _token, 64, path); // key 있음
            int result2 = GetPrivateProfileString("SECTION", "CHAT_ID", "def", _chatId, 32, path); // key 없음
            int result3 = GetPrivateProfileString("SECTION", "UCWINDOW", "def", _ucWindow, 64, path); // key 없음
            int result4 = GetPrivateProfileString("SECTION", "KAKAOWINDOW", "def", _kakaoWindow, 64, path); // key 없음
            int result5 = GetPrivateProfileString("SECTION", "DELAY", "def", _delayValue, 64, path); // key 없음

            //int result = GetPrivateProfileString("SECTION", "KEY2", "def", retVal3, 32, "C:\none.ini"); // ini 파일 없음
        }

        //searchIMG에 스크린 이미지와 찾을 이미지를 넣어줄거에요
        public void searchIMG(Bitmap screen_img, Bitmap find_img)
        {
            //스크린 이미지 선언
            using (Mat ScreenMat = BitmapConverter.ToMat(screen_img))
            //찾을 이미지 선언
            using (Mat FindMat = BitmapConverter.ToMat(find_img))
            //스크린 이미지에서 FindMat 이미지를 찾아라
            using (Mat res = ScreenMat.MatchTemplate(FindMat, TemplateMatchModes.CCoeffNormed))
            {
                //찾은 이미지의 유사도를 담을 더블형 최대 최소 값을 선언합니다.
                double minval, maxval = 0;
                //찾은 이미지의 위치를 담을 포인트형을 선업합니다.
                OpenCvSharp.Point minloc, maxloc;
                //찾은 이미지의 유사도 및 위치 값을 받습니다. 
                Cv2.MinMaxLoc(res, out minval, out maxval, out minloc, out maxloc);
                metroListView1.Items.Add("찾은 이미지의 유사도 : " + maxval);

                //이미지를 찾았을 경우 클릭이벤트를 발생!!
                if (maxval >= 0.7)
                {
                    InClick(maxloc.X, maxloc.Y);
                }
            }
        }

        public static int MakeLParam(int x, int y)
        {
            return ((y << 16) | (x & 0xffff));
        }

        public void InClick(int x, int y)
        {
            IntPtr window = FindWindow("Notepad", null);
            if (!(window != IntPtr.Zero))
                return;
            int lParam = MakeLParam(x, y);
            if (!(window != IntPtr.Zero))
                return;
            PostMessage(window, 513, 1, lParam);
            PostMessage(window, 514, 0, lParam);
        }

        //x,y 값을 전달해주면 클릭이벤트를 발생합니다.
        public void sendKakao(Bitmap bmp = null)
        {
            //클릭이벤트를 발생시킬 플레이어를 찾습니다.
            IntPtr findwindow = FindWindow(null, _kakaoWindow.ToString());

            if (findwindow != IntPtr.Zero)
            {
                //플레이어를 찾았을 경우 클릭이벤트를 발생시킬 핸들을 가져옵니다.
                IntPtr hwnd_child = FindWindowEx(findwindow, IntPtr.Zero, "RICHEDIT50W", "");

                //Debug.WriteLine("핸들 : " + hwnd_child);

                if (hwnd_child != IntPtr.Zero)
                {
                    //메세지
                    if (bmp != null)
                    {
                        Clipboard.SetImage(bmp);
                        SetForegroundWindow(hwnd_child);
                        Delay(delayValue);
                        SendKeys.Send("^v");
                        SendKeys.Send("~");
                        Delay(delayValue);
                        IntPtr hwnd_return = FindWindow("SWT_Window0", null);
                        SetForegroundWindow(hwnd_return);
                    }
                    else
                    {
                        SendMessage(hwnd_child, 0x000c, 0, DateTime.Now.ToString("yyyy-MM-dd-HH:mm:ss") + " 메세지 도착");
                    }
                    Delay(delayValue);
                }
            }
        }

        private static Bitmap GetBitmap(IntPtr findwindow)
        {
            Rectangle rectangle = Rectangle.Round(Graphics.FromHwnd(findwindow).VisibleClipBounds);
            Bitmap bitmap = new Bitmap(rectangle.Width, rectangle.Height);
            using (Graphics graphics = Graphics.FromImage(bitmap))
            {
                IntPtr hdc = graphics.GetHdc();
                PrintWindow(findwindow, hdc, 2);
                graphics.ReleaseHdc(hdc);
            }
            return bitmap;
        }

        private static int SendTelegram(string message)
        {
            JObject.Parse(new StreamReader(WebRequest.Create("https://api.telegram.org/bot" + _token.ToString() + "/sendMessage?chat_id=" + _chatId.ToString() + "&text=" + message).GetResponse().GetResponseStream() ?? throw new InvalidOperationException()).ReadToEnd());
            return 0;
        }

        public async Task SendTelegramPhoto(Bitmap bitmap)
        {
            try
            {
                string startupPath = Application.StartupPath;
                bitmap.Save(startupPath + "\\img\\scr.bmp");
                Delay(delayValue);
                FileStream fileStream = File.Open(startupPath + "\\img\\scr.bmp", FileMode.Open);
                Delay(delayValue);
                await Bot.SendPhotoAsync(_chatId.ToString(), fileStream, null, ParseMode.Default, false, 0, null, new CancellationToken());
                Delay(delayValue);
            }
            catch (Exception e)
            {
                Debug.WriteLine(e.Message);
            }
        }

        private static void Delay(int MS)
        {
            DateTime now = DateTime.Now;
            TimeSpan timeSpan = new TimeSpan(0, 0, 0, 0, MS);
            for (DateTime dateTime = now.Add(timeSpan); dateTime >= now; now = DateTime.Now)
                Application.DoEvents();
        }

        public async Task monitorWork()
        {
            while (true)
            {
                //Thread.Sleep(500);
                if (checkprocess)
                {
                    IntPtr window = FindWindow(_ucWindow.ToString(), null);
                    if (window != IntPtr.Zero)
                    {
                        metroListView1.Items.Add(DateTime.Now.ToString("yyyy-MM-dd-HH:mm:ss") + " : 프로세스 찾았습니다.");
                        Bitmap bitmap = GetBitmap(window);
                        pictureBox1.Image = bitmap;
                        await SendTelegramPhoto(bitmap);
                        Delay(Int32.Parse(_delayValue.ToString()));
                    }
                    await Task.Delay(delayValue);
                }
                else
                    break;
            }
        }

        private async void metroButton1_Click(object sender, EventArgs e)
        {
            checkprocess = true;
            metroListView1.Items.Add(DateTime.Now.ToString("yyyy-MM-dd-HH:mm:ss") + " : 모니터링 시작");
            await monitorWork();
            metroListView1.Items.Add(DateTime.Now.ToString("yyyy-MM-dd-HH:mm:ss") + " : 모니터링 종료");
        }

        private void metroButton2_Click(object sender, EventArgs e)
        {
            checkprocess = false;
            Delay(delayValue);
            metroListView1.Items.Clear();
        }

        private void metroButton3_Click(object sender, EventArgs e)
        {
            metroListView1.Items.Add("테스트!!!");
        }

        private void Btn_KeyEvent(object sender, KeyEventArgs e)
        {
            if (e.KeyCode != Keys.Escape)
                return;
            checkprocess = false;
            metroListView1.Items.Add("시스템멈춤!!!");
        }

        private void Form1_Resize(object sender, EventArgs e)
        {
            /*
            if (this.WindowState != FormWindowState.Minimized)
                return;
            this.Visible = false;
            this.ShowIcon = false;
            this.notifyIcon1.Visible = true;
            */

            if (FormWindowState.Minimized == WindowState)
            {
                notifyIcon1.Visible = true; // tray icon 표시
                this.Hide();
            }
            else if (FormWindowState.Normal == WindowState)
            {
                notifyIcon1.Visible = false;
                ShowInTaskbar = true; // 작업 표시줄 표시
            }

        }

        private void NotifyIcon1_DoubleClick(object sender, MouseEventArgs e)
        {
            //this.Visible = true;
            //this.ShowIcon = true;
            notifyIcon1.Visible = false;
            Show();
            IntPtr hwnd_return = FindWindow(null, "팝업모니터링");
            if (hwnd_return != IntPtr.Zero)
            {
                ShowWindowAsync(hwnd_return, 1);
                SetForegroundWindow(hwnd_return);
            }
        }
    }
}
