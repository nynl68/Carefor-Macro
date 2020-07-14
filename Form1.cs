using System;
using System.Drawing;
using System.Runtime.InteropServices;
using System.Threading;
using System.Windows.Forms;

namespace CareForMacro
{
    public partial class MacroForm : Form
    {
        public bool started = false;
        Macro macro;
        Thread macroThread;

        public MacroForm()
        {
            InitializeComponent();
        }

        [DllImport("USER32.DLL")] //핫키등록
        private static extern int RegisterHotKey(int hwnd, int id, int fsModifiers, int vk);

        [DllImport("USER32.DLL")] //핫키제거
        private static extern int UnregisterHotKey(int hwnd, int id);

        private void MacroForm_Load(object sender, EventArgs e)
        {
            RegisterHotKey((int)Handle, 88, 0x0, (int)Keys.F8); // 시작 // (여기로가져와, 니 ID는 0이야, 조합키안써, F8눌러지면)
            RegisterHotKey((int)Handle, 99, 0x0, (int)Keys.F9); // 종료, F9
            RegisterHotKey((int)Handle, 1, 0x0002, (int)Keys.D1); // 컨트롤 1
            RegisterHotKey((int)Handle, 2, 0x0002, (int)Keys.D2); // 컨트롤 2
            RegisterHotKey((int)Handle, 3, 0x0002, (int)Keys.D3); // 컨트롤 3 
            RegisterHotKey((int)Handle, 4, 0x0002, (int)Keys.D4); // 컨트롤 4
            RegisterHotKey((int)Handle, 5, 0x0002, (int)Keys.D5); // 컨트롤 5
            RegisterHotKey((int)Handle, 6, 0x0002, (int)Keys.D6); // 컨트롤 6
            RegisterHotKey((int)Handle, 7, 0x0002, (int)Keys.D7); // 컨트롤 7
            RegisterHotKey((int)Handle, 8, 0x0002, (int)Keys.D8); // 컨트롤 8
            RegisterHotKey((int)Handle, 9, 0x0002, (int)Keys.D9); // 컨트롤 9
            RegisterHotKey((int)Handle, 0, 0x0002, (int)Keys.D0); // 컨트롤 0
            RegisterHotKey((int)Handle, 11, 0x0002, (int)Keys.OemMinus); // - 버튼 -> 투약기록 저장
            RegisterHotKey((int)Handle, 12, 0x0002, (int)Keys.Oemplus); // = 버튼 -> 다음 날짜
        }

        private void MacroForm_FormClosing(object sender, FormClosingEventArgs e)
        {
            UnregisterHotKey((int)Handle, 1); // 이 폼에 ID가 1인 핫키 해제
            UnregisterHotKey((int)Handle, 2);
            UnregisterHotKey((int)Handle, 3);
            UnregisterHotKey((int)Handle, 4);
            UnregisterHotKey((int)Handle, 5);
            UnregisterHotKey((int)Handle, 6);
            UnregisterHotKey((int)Handle, 7);
            UnregisterHotKey((int)Handle, 8);
            UnregisterHotKey((int)Handle, 9);
            UnregisterHotKey((int)Handle, 10);
            UnregisterHotKey((int)Handle, 88);
            UnregisterHotKey((int)Handle, 99);
            UnregisterHotKey((int)Handle, 11);
            UnregisterHotKey((int)Handle, 12);
        }

        protected override void WndProc(ref Message m) // 윈도우프로시저 콜백함수
        {
            base.WndProc(ref m);

            if (m.Msg == 0x0312) // 핫키가 눌러지면 312 정수 메세지를 받게됨
            {
                if (m.WParam == (IntPtr)88) // 그 키의 ID가 88이면 (F8이 눌러지면)
                {
                    if (started)
                    {
                        changeState("매크로가 이미 시작되었습니다.");
                        lb_state.ForeColor = Color.Red;
                        return;
                    }

                    try
                    {
                        Macro.num = Convert.ToInt32(tb_num.Text);
                        Macro.아침식전_시 = Convert.ToInt32(tb1.Text);
                        Macro.아침식전_분 = Convert.ToInt32(tb2.Text);
                        Macro.아침식후_시 = Convert.ToInt32(tb3.Text);
                        Macro.아침식후_분 = Convert.ToInt32(tb4.Text);
                        Macro.점심식후_시 = Convert.ToInt32(tb5.Text);
                        Macro.점심식후_분 = Convert.ToInt32(tb6.Text);
                        Macro.저녁식후_시 = Convert.ToInt32(tb7.Text);
                        Macro.저녁식후_분 = Convert.ToInt32(tb8.Text);
                        Macro.취침전_시 = Convert.ToInt32(tb9.Text);
                        Macro.취침전_분 = Convert.ToInt32(tb0.Text);
                        Macro.interval = Convert.ToInt32(tb_Interval.Text);
                    }
                    catch (FormatException)
                    {
                        lb_state.ForeColor = Color.Red;
                        lb_state.Text = "올바른 숫자를 입력해주세요.";
                        return;
                    }
                    if (Macro.num < 0 || Macro.num > 9999)
                    {
                        lb_state.ForeColor = Color.Red;
                        lb_state.Text = "올바른 숫자를 입력해주세요.";
                        return;
                    }

                    lb_state.ForeColor = Color.Blue;
                    lb_state.Text = "매크로를 시작합니다.";
                    started = true;

                    macro = new Macro(this);
                    macroThread = new Thread(macro.work);

                    macroThread.Start();
                }
                else if (m.WParam == (IntPtr)99) // 그 키의 ID가 99이면 (F9가 눌러지면)
                {
                    try
                    {
                        macroThread.Abort();
                    }
                    catch (Exception)
                    {

                    }
                    finally
                    {
                        if (started == true)
                        {
                            lb_state.ForeColor = Color.Blue;
                            started = false;
                            lb_state.Text = "매크로가 종료되었습니다.";
                        }
                        else
                        {
                            lb_state.ForeColor = Color.Red;
                            lb_state.Text = "매크로가 시작되지 않은 상태입니다.";
                        }
                    }
                }
                else if (m.WParam == (IntPtr)1)
                {
                    Macro.pt_아침식전_시.X = Convert.ToInt32(Cursor.Position.X.ToString());
                    Macro.pt_아침식전_시.Y = Convert.ToInt32(Cursor.Position.Y.ToString());
                    changeState("좌표 1 설정 완료");
                }
                else if (m.WParam == (IntPtr)2)
                {
                    Macro.pt_아침식전_분.X = Convert.ToInt32(Cursor.Position.X.ToString());
                    Macro.pt_아침식전_분.Y = Convert.ToInt32(Cursor.Position.Y.ToString());
                    changeState("좌표 2 설정 완료");
                }
                else if (m.WParam == (IntPtr)3)
                {
                    Macro.pt_아침식후_시.X = Convert.ToInt32(Cursor.Position.X.ToString());
                    Macro.pt_아침식후_시.Y = Convert.ToInt32(Cursor.Position.Y.ToString());
                    changeState("좌표 3 설정 완료");
                }
                else if (m.WParam == (IntPtr)4)
                {
                    Macro.pt_아침식후_분.X = Convert.ToInt32(Cursor.Position.X.ToString());
                    Macro.pt_아침식후_분.Y = Convert.ToInt32(Cursor.Position.Y.ToString());
                    changeState("좌표 4 설정 완료");
                }
                else if (m.WParam == (IntPtr)5)
                {
                    Macro.pt_점심식후_시.X = Convert.ToInt32(Cursor.Position.X.ToString());
                    Macro.pt_점심식후_시.Y = Convert.ToInt32(Cursor.Position.Y.ToString());
                    changeState("좌표 5 설정 완료");
                }
                else if (m.WParam == (IntPtr)6)
                {
                    Macro.pt_점심식후_분.X = Convert.ToInt32(Cursor.Position.X.ToString());
                    Macro.pt_점심식후_분.Y = Convert.ToInt32(Cursor.Position.Y.ToString());
                    changeState("좌표 6 설정 완료");
                }
                else if (m.WParam == (IntPtr)7)
                {
                    Macro.pt_저녁식후_시.X = Convert.ToInt32(Cursor.Position.X.ToString());
                    Macro.pt_저녁식후_시.Y = Convert.ToInt32(Cursor.Position.Y.ToString());
                    changeState("좌표 7 설정 완료");
                }
                else if (m.WParam == (IntPtr)8)
                {
                    Macro.pt_저녁식후_분.X = Convert.ToInt32(Cursor.Position.X.ToString());
                    Macro.pt_저녁식후_분.Y = Convert.ToInt32(Cursor.Position.Y.ToString());
                    changeState("좌표 8 설정 완료");
                }
                else if (m.WParam == (IntPtr)9)
                {
                    Macro.pt_취침전_시.X = Convert.ToInt32(Cursor.Position.X.ToString());
                    Macro.pt_취침전_시.Y = Convert.ToInt32(Cursor.Position.Y.ToString());
                    changeState("좌표 9 설정 완료");
                }
                else if (m.WParam == (IntPtr)0)
                {
                    Macro.pt_취침전_분.X = Convert.ToInt32(Cursor.Position.X.ToString());
                    Macro.pt_취침전_분.Y = Convert.ToInt32(Cursor.Position.Y.ToString());
                    changeState("좌표 0 설정 완료");
                }
                else if (m.WParam == (IntPtr)11)
                {
                    Macro.pt_Minus.X = Convert.ToInt32(Cursor.Position.X.ToString());
                    Macro.pt_Minus.Y = Convert.ToInt32(Cursor.Position.Y.ToString());
                    changeState("좌표 - 설정 완료");
                }
                else if (m.WParam == (IntPtr)12)
                {
                    Macro.pt_Add.X = Convert.ToInt32(Cursor.Position.X.ToString());
                    Macro.pt_Add.Y = Convert.ToInt32(Cursor.Position.Y.ToString());
                    changeState("좌표 = 설정 완료");
                }
            }
        }

        public void changeState(String str)
        {
            lb_state.Text = str;
        }
    }

    class Macro
    {
        public const int LEFT_DOWN = 0x02;
        public const int LEFT_UP = 0x04;

        public static Point pt_아침식전_시 = new Point();
        public static Point pt_아침식전_분 = new Point();
        public static Point pt_아침식후_시 = new Point();
        public static Point pt_아침식후_분 = new Point();
        public static Point pt_점심식후_시 = new Point();
        public static Point pt_점심식후_분 = new Point();
        public static Point pt_저녁식후_시 = new Point();
        public static Point pt_저녁식후_분 = new Point();
        public static Point pt_취침전_시 = new Point();
        public static Point pt_취침전_분 = new Point();
        public static Point pt_Minus = new Point();
        public static Point pt_Add = new Point();

        public static int 아침식전_시 = 0;
        public static int 아침식전_분 = 0;
        public static int 아침식후_시 = 0;
        public static int 아침식후_분 = 0;
        public static int 점심식후_시 = 0;
        public static int 점심식후_분 = 0;
        public static int 저녁식후_시 = 0;
        public static int 저녁식후_분 = 0;
        public static int 취침전_시 = 0;
        public static int 취침전_분 = 0;
        public static int interval = 1000;
        
        public static int num = 0; // 몇 번 반복할 지

        MacroForm parent = null;

        public Macro(MacroForm parent)
        {
            this.parent = parent;
        }

        // Get a handle to an application window.
        [DllImport("USER32.DLL", CharSet = CharSet.Unicode)]
        public static extern IntPtr FindWindow(string lpClassName, string lpWindowName);

        // Activate an application window.
        [DllImport("USER32.DLL")]
        public static extern bool SetForegroundWindow(IntPtr hWnd);

        [DllImport("USER32.DLL")]
        public static extern void mouse_event(int dwFlags, int dx, int dy, int cButtons, int dwExtraInfo);

        [DllImport("USER32.DLL")]
        public static extern int SetCursorPos(int x, int y);

        [DllImport("USER32.DLL")]
        internal static extern bool GetWindowPlacement(int hWnd, ref WINDOWPLACEMENT lpwndpl);

        internal struct WINDOWPLACEMENT
        {
            public int length;
            public int flags;
            public ShowWindowCommands showCmd;
            public Point ptMinPosition;
            public Point ptMaxPosition;
            public Rectangle rcNormalPosition;
        }

        internal enum ShowWindowCommands : int
        {
            Hide = 0,
            Normal = 1,
            Minimized = 2,
            Maximized = 3,
        }

        internal enum WNDSTATE : int
        {
            SW_HIDE = 0,
            SW_SHOWNORMAL = 1,
            SW_NORMAL = 1,
            SW_SHOWMINIMIZED = 2,
            SW_MAXIMIZE = 3,
            SW_SHOWNOACTIVATE = 4,
            SW_SHOW = 5,
            SW_MINIMIZE = 6,
            SW_SHOWMINNOACTIVE = 7,
            SW_SHOWNA = 8,
            SW_RESTORE = 9,
            SW_SHOWDEFAULT = 10,
            SW_MAX = 10
        }

        private static WINDOWPLACEMENT GetPlacement(int hwnd)
        {
            WINDOWPLACEMENT placement = new WINDOWPLACEMENT();
            placement.length = Marshal.SizeOf(placement);
            GetWindowPlacement(hwnd, ref placement);
            return placement;
        }

        private void GetWindowPos(int hwnd, ref int ptrPhwnd, ref int ptrNhwnd, ref Point ptPoint, ref Size szSize)
        {
            WINDOWPLACEMENT wInf = new WINDOWPLACEMENT();
            wInf.length = System.Runtime.InteropServices.Marshal.SizeOf(wInf);
            GetWindowPlacement(hwnd, ref wInf);
            szSize = new Size(wInf.rcNormalPosition.Right - (wInf.rcNormalPosition.Left * 2), wInf.rcNormalPosition.Bottom - (wInf.rcNormalPosition.Top * 2));
            ptPoint = new Point(wInf.rcNormalPosition.Left, wInf.rcNormalPosition.Top);
        }

        public void work()
        {
            // 인터넷 창이 켜져있는지 확인
            IntPtr internetHandle = FindWindow(null, "3-1.간호급여 제공기록 - Internet Explorer");

            // 인터넷 창이 꺼져있으면
            if (internetHandle == IntPtr.Zero)
            {
                parent.changeState("케어포가 켜져 있지 않습니다.");
                return;
            }

            int ptrPhwnd = 0, ptrNhwnd = 0;
            Point ptPoint = new Point();
            Size szSize = new Size();
            GetWindowPos((int)internetHandle, ref ptrPhwnd, ref ptrNhwnd, ref ptPoint, ref szSize); // 인터넷 창 위치와 사이즈 받아오기
            
            // 케어포 프로세스를 포그라운드로 옮기고 명령을 보냄
            SetForegroundWindow(internetHandle);
            
            for (int i = 0; i < num; i++)
            {
                /* 명옥례 비위관 관리 */
                Thread.Sleep(1000); // 1초 딜레이

                oneLine(0, "7", "0", 0, "500");
                oneLine(30, "10", "0", 5, "100");
                oneLine(60, "12", "0", 0, "500");
                oneLine(90, "15", "0", 5, "100");
                oneLine(120, "17", "0", 0, "500");
                oneLine(150, "20", "0", 5, "100");

                doubleClick(pt_아침식전_분.X, pt_아침식전_분.Y); // 저장 버튼
                Thread.Sleep(2000);

                click(pt_아침식후_시.X, pt_아침식후_시.Y); // 다음날 버튼
                Thread.Sleep(1000);


                /* 최종 확인 작업 
                click(pt_아침식전_시.X, pt_아침식전_시.Y); // 저장 버튼
                Thread.Sleep(interval/2);

                click(pt_아침식전_분.X, pt_아침식전_분.Y); // 다음날 버튼
                Thread.Sleep(interval);
                */

                /* -> 변기일 어르신 쿠티아핀정 5월 4일부터 아침식후 빼는 작업

                SetCursorPos(pt_아침식전_분.X, pt_아침식전_분.Y); // 마지막 커서 위치
                Thread.Sleep(50);

                doubleClick(pt_아침식전_분.X, pt_아침식전_분.Y); // 클릭
                Thread.Sleep(interval / 3);

                click(pt_아침식후_시.X, pt_아침식후_시.Y); // 쿠티아핀정 아침 식후 버튼
                Thread.Sleep(interval / 2);

                click(pt_아침식후_분.X, pt_아침식후_분.Y); // 투약기록 저장 버튼
                Thread.Sleep(2*interval / 5);

                SendKeys.SendWait("{Enter}");
                Thread.Sleep(interval / 2);

                click(pt_점심식후_시.X, pt_점심식후_시.Y); // 다음 날짜 버튼
                Thread.Sleep(3 * interval / 4);
                */

                /* -> 취침전 제공자 노미자로 바꾸는 작업 

                click(pt_아침식전_시.X, pt_아침식전_시.Y); // 담당자 "선택" 버튼
                Thread.Sleep(interval);

                doubleClick(pt_점심식후_분.X, pt_점심식후_분.Y); // 스크롤 아래로 가기 버튼
                Thread.Sleep(interval / 4);

                click(pt_아침식전_분.X, pt_아침식전_분.Y); // 노미자 버튼
                Thread.Sleep(interval/2);

                click(pt_아침식후_시.X, pt_아침식후_시.Y); // 담당자 선택 완료 버튼
                Thread.Sleep(interval/3);

                // 이 부분에서 노미자 일정없음으로 나오면 멈춤

                click(pt_아침식후_분.X, pt_아침식후_분.Y); // 투약기록 저장 버튼
                Thread.Sleep(interval/2);

                click(pt_점심식후_시.X, pt_점심식후_시.Y); // 다음 날짜 버튼
                Thread.Sleep(3*interval/4);
                */


                /* -> 투약일지 시간, 분 넣는 작업
                Thread.Sleep(1000); // 1초 딜레이

                doubleClick(pt_아침식전_시.X, pt_아침식전_시.Y); // 아침식전_시 위치
                SendKeys.SendWait(아침식전_시.ToString());
                Thread.Sleep(interval);

                doubleClick(pt_아침식전_분.X, pt_아침식전_분.Y); // 아침식전_분 위치
                SendKeys.SendWait(아침식전_분.ToString());
                Thread.Sleep(interval);

                doubleClick(pt_아침식후_시.X, pt_아침식후_시.Y); // 아침식후_시 위치
                SendKeys.SendWait(아침식후_시.ToString());
                Thread.Sleep(interval);

                doubleClick(pt_아침식후_분.X, pt_아침식후_분.Y); // 아침식후_분 위치
                SendKeys.SendWait(아침식후_분.ToString());
                Thread.Sleep(interval);

                doubleClick(pt_점심식후_시.X, pt_점심식후_시.Y); // 점심식후_시 위치
                SendKeys.SendWait(점심식후_시.ToString());
                Thread.Sleep(interval);

                doubleClick(pt_점심식후_분.X, pt_점심식후_분.Y); // 점심식후_분 위치
                SendKeys.SendWait(점심식후_분.ToString());
                Thread.Sleep(interval);

                doubleClick(pt_저녁식후_시.X, pt_저녁식후_시.Y); // 저녁식후_시 위치
                SendKeys.SendWait(저녁식후_시.ToString());
                Thread.Sleep(interval);

                doubleClick(pt_저녁식후_분.X, pt_저녁식후_분.Y); // 저녁식후_분 위치
                SendKeys.SendWait(저녁식후_분.ToString());
                Thread.Sleep(interval);

                doubleClick(pt_취침전_시.X, pt_취침전_시.Y); // 취침전_시 위치
                SendKeys.SendWait(취침전_시.ToString());
                Thread.Sleep(interval);

                doubleClick(pt_취침전_분.X, pt_취침전_분.Y); // 취침전_분 위치
                SendKeys.SendWait(취침전_분.ToString());
                Thread.Sleep(interval);
                
                click(pt_Minus.X, pt_Minus.Y);
                Thread.Sleep(1000);

                click(pt_Minus.X, pt_Minus.Y);
                Thread.Sleep(500);

                //SendKeys.SendWait("{Enter}");
                //Thread.Sleep(2000);

                click(pt_Add.X, pt_Add.Y);
                Thread.Sleep(4000);

                */

                parent.changeState(i + 1 + "회 반복중!");
            }
            parent.changeState("입력 " + num.ToString() + "회 완료, 매크로가 종료되었습니다.");
            parent.started = false;
        }

        /* 비위관 영양 관리 한 줄 */
        private void oneLine(int y, String hour, String minute, int type, String ml)
        {
            doubleClick(pt_아침식전_시.X, pt_아침식전_시.Y + y); // 식사 시간
            SendKeys.SendWait(hour);
            Thread.Sleep(interval);

            doubleClick(pt_아침식전_시.X + 35, pt_아침식전_시.Y + y); // 식사 분
            SendKeys.SendWait(minute);
            Thread.Sleep(interval);

            click(pt_아침식전_시.X + 100, pt_아침식전_시.Y + y); // 식사 종류
            Thread.Sleep(interval);

            for (int i = 0; i < 9; i++) // 물
            {
                SendKeys.SendWait("{UP}");
                Thread.Sleep(20);
            }

            for (int i = 0; i < type; i++) // type 번째
            {
                SendKeys.SendWait("{DOWN}");
                Thread.Sleep(20);
            }
            Thread.Sleep(interval);


            doubleClick(pt_아침식전_시.X + 180, pt_아침식전_시.Y + y);
            SendKeys.SendWait(ml); // 식사량
            Thread.Sleep(interval);
        }

        private void doubleClick(int x, int y)
        {
            SetCursorPos(x, y); // 커서 위치
            Thread.Sleep(50);
            mouse_event(LEFT_DOWN, 0, 0, 0, 0); // 클릭했다가
            Thread.Sleep(50);
            mouse_event(LEFT_UP, 0, 0, 0, 0); // 뗌
            Thread.Sleep(50);
            mouse_event(LEFT_DOWN, 0, 0, 0, 0); // 클릭했다가
            Thread.Sleep(50);
            mouse_event(LEFT_UP, 0, 0, 0, 0); // 뗌
        }

        private void click(int x, int y)
        {
            SetCursorPos(x, y); // 커서 위치
            Thread.Sleep(50);
            mouse_event(LEFT_DOWN, 0, 0, 0, 0); // 클릭했다가
            Thread.Sleep(50);
            mouse_event(LEFT_UP, 0, 0, 0, 0); // 뗌
        }
    }
}