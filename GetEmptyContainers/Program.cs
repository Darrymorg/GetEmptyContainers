using Microsoft.Office.Interop.Excel;
using System;
using System.Diagnostics;
using System.Runtime.InteropServices;
using System.Text;
using Excel = Microsoft.Office.Interop.Excel;

internal class Program
{
    [DllImport("user32.dll")] private static extern IntPtr GetForegroundWindow();
    [DllImport("user32.dll")] private static extern bool SetForegroundWindow(IntPtr hWnd);
    [DllImport("user32.dll")] private static extern int GetWindowText(IntPtr hWnd, System.Text.StringBuilder lpString, int nMaxCount);
    [DllImport("user32.dll")] private static extern IntPtr FindWindow(string lpClassName, string lpWindowName);
    [DllImport("user32.dll")] private static extern IntPtr FindWindowEx(IntPtr hwndParent, IntPtr hwndChildAfter, string lpszClass, string lpszWindow);
    [DllImport("user32.dll")] private static extern bool PostMessage(IntPtr hWnd, uint Msg, IntPtr wParam, IntPtr lParam);
    [DllImport("user32.dll")] private static extern int GetClassName(IntPtr hWnd, StringBuilder lpString, int nMaxCount);
    [DllImport("user32.dll")] private static extern bool EnumChildWindows(IntPtr hWndParent, EnumChildProc lpEnumFunc, IntPtr lParam);
    [DllImport("user32.dll")] static extern bool SetCursorPos(int X, int Y);
    [DllImport("user32.dll")] public static extern void mouse_event(uint dwFlags, uint dx, uint dy, uint cButtons, UIntPtr dwExtraInfo);
    [DllImport("user32.dll")] public static extern int SendMessage(IntPtr hWnd, int uMsg, int wParam, string lParam);
    [DllImport("user32.dll")] public static extern bool SetWindowText(IntPtr hwnd, String lpString);
    [DllImport("user32.dll")] public static extern void keybd_event(byte virtualKey, byte scanCode, uint flags, IntPtr extraInfo);
    private const uint MOUSEEVENTF_LEFTDOWN = 0x02;
    private const uint MOUSEEVENTF_LEFTUP = 0x04;
    private delegate bool EnumChildProc(IntPtr hWnd, IntPtr lParam);
    public const uint WM_CLICK = 0x00F5;
    public const uint VK_F5 = 0x74;
    const int KEYEVENTF_EXTENDEDKEY = 0x0001;
    const int KEYEVENTF_KEYUP = 0x0002;
    const byte VK_LWIN = 0x5B;
    const byte VK_D = 0x44;
    private static void Main(string[] args)
    {
        #region GetNotEmptyContainers
        //"Стилон 6 / "
        //Mouse move (1817,950)
        //Mouse move (575,73)
        //Mouse move (498,231)
        //Mouse move (355,323)
        //295 388
        //Mouse move (103,330)
        //Mouse move (740,337)
        //Mouse move (521,516)
        //Mouse move (462,616)
        //Mouse move (1003,622)
        int smallLatency = 1500;
        int bigLatency = 2500;

        if (File.Exists("C:\\Users\\User\\Desktop\\d.xls"))
        {
            File.Delete("C:\\Users\\User\\Desktop\\d.xls");
        }
        if (File.Exists("C:\\Users\\User\\Desktop\\dd.xlsx"))
        {
            File.Delete("C:\\Users\\User\\Desktop\\dd.xlsx");
        }
        IntPtr hWnd = FindWindow(null, "Стилон 6 / ");
        SetForegroundWindow(hWnd);

        Thread.Sleep(smallLatency);

        Thread.Sleep(smallLatency);
        SetCursorPos(1817, 950);
        mouse_event(MOUSEEVENTF_LEFTDOWN | MOUSEEVENTF_LEFTUP, 575, 73, 0, 0);
        SetCursorPos(575, 73);
        Thread.Sleep(smallLatency);
        mouse_event(MOUSEEVENTF_LEFTDOWN | MOUSEEVENTF_LEFTUP, 575, 73, 0, 0);
        SetCursorPos(498, 231);
        Thread.Sleep(smallLatency);
        mouse_event(MOUSEEVENTF_LEFTDOWN | MOUSEEVENTF_LEFTUP, 575, 73, 0, 0);
        SetCursorPos(355, 323);
        Thread.Sleep(smallLatency);
        mouse_event(MOUSEEVENTF_LEFTDOWN | MOUSEEVENTF_LEFTUP, 575, 73, 0, 0);
        SetCursorPos(295, 388);
        Thread.Sleep(smallLatency);
        mouse_event(MOUSEEVENTF_LEFTDOWN | MOUSEEVENTF_LEFTUP, 575, 73, 0, 0);
        SetCursorPos(103, 330);
        Thread.Sleep(smallLatency);
        mouse_event(MOUSEEVENTF_LEFTDOWN | MOUSEEVENTF_LEFTUP, 575, 73, 0, 0);
        SetCursorPos(740, 337);
        Thread.Sleep(bigLatency);
        mouse_event(MOUSEEVENTF_LEFTDOWN | MOUSEEVENTF_LEFTUP, 575, 73, 0, 0);
        Thread.Sleep(smallLatency);
        hWnd = FindWindow(null, "Сохранение");
        Thread.Sleep(smallLatency);


        if (hWnd != IntPtr.Zero)
        {
            keybd_event(VK_D, 0, KEYEVENTF_EXTENDEDKEY, IntPtr.Zero); // Press D key
            keybd_event(VK_D, 0, KEYEVENTF_EXTENDEDKEY | KEYEVENTF_KEYUP, IntPtr.Zero); ;
            SetCursorPos(521, 516);
            Thread.Sleep(smallLatency);
            mouse_event(MOUSEEVENTF_LEFTDOWN | MOUSEEVENTF_LEFTUP, 575, 73, 0, 0);
            SetCursorPos(1003, 622);
            Thread.Sleep(smallLatency);
            mouse_event(MOUSEEVENTF_LEFTDOWN | MOUSEEVENTF_LEFTUP, 575, 73, 0, 0);
            Thread.Sleep(smallLatency);
            var save = FindWindowEx(hWnd, IntPtr.Zero, "Button", null);
            if (save != IntPtr.Zero)
            {
                PostMessage(save, WM_CLICK, IntPtr.Zero, IntPtr.Zero);
            }
        }
        else
        {
            Console.WriteLine("Окно не найдено");
        }
        /*while (true)
        {
            IntPtr hWnds = GetForegroundWindow();
            System.Text.StringBuilder windowText = new System.Text.StringBuilder(256);
            GetWindowText(hWnds, windowText, windowText.Capacity);
            Console.WriteLine("Заголовок активно окна: " + windowText + ".");
            System.Threading.Thread.Sleep(5000);
        }*/
        Excel.Application application = null;
        Excel.Workbook workbook = null;
        Excel.Worksheet worksheet = null;
        List<string> notEmptyContainers = null;
        try
        {
            application = new Excel.Application();

            workbook = application.Workbooks.Open("C:\\Users\\User\\Desktop\\d.xls", 0, false, 5, "", "", false, Excel.XlPlatform.xlWindows, "", true, false, 0, true, false, false);
            worksheet = workbook.Sheets[1];
            int numCul = 14;
            Excel.Range usedColumn = worksheet.UsedRange.Columns[numCul];
            System.Array myValues = (System.Array)usedColumn.Cells.Value2;
            notEmptyContainers = myValues.OfType<object>().Select(o => o.ToString()).ToList();
            application.Quit();
            foreach (var item in notEmptyContainers)
            {
                //Console.WriteLine(item);
            }
        }
        catch (Exception)
        {

            throw;
        }
        #endregion
        #region GetAllContainers
        //"Стилон 6 / "
        //Mouse move (282,74)
        //Mouse move (614,232)
       // Mouse move(218,401)

        //Mouse move (1846,159)
        //Mouse move (1551,695)
        //Mouse move (1016,662)
        //Mouse move (1743,120)
        //Input
        //Mouse move (756,517)
        //Mouse move (508,633)

        SetCursorPos(282, 74);
        mouse_event(MOUSEEVENTF_LEFTDOWN | MOUSEEVENTF_LEFTUP, 575, 73, 0, 0);
        SetCursorPos(614, 232);
        Thread.Sleep(smallLatency);
        mouse_event(MOUSEEVENTF_LEFTDOWN | MOUSEEVENTF_LEFTUP, 575, 73, 0, 0);
        SetCursorPos(218, 401);
        Thread.Sleep(smallLatency);
        mouse_event(MOUSEEVENTF_LEFTDOWN | MOUSEEVENTF_LEFTUP, 575, 73, 0, 0);
        SetCursorPos(1551, 695);
        Thread.Sleep(smallLatency);
        mouse_event(MOUSEEVENTF_LEFTDOWN | MOUSEEVENTF_LEFTUP, 575, 73, 0, 0);
        SetCursorPos(1846, 159);
        Thread.Sleep(smallLatency);
        mouse_event(MOUSEEVENTF_LEFTDOWN | MOUSEEVENTF_LEFTUP, 575, 73, 0, 0);
        SetCursorPos(1551, 695);
        Thread.Sleep(smallLatency);
        mouse_event(MOUSEEVENTF_LEFTDOWN | MOUSEEVENTF_LEFTUP, 575, 73, 0, 0);
        SetCursorPos(1016, 662);
        Thread.Sleep(smallLatency);
        mouse_event(MOUSEEVENTF_LEFTDOWN | MOUSEEVENTF_LEFTUP, 575, 73, 0, 0);
        SetCursorPos(1743, 120);
        Thread.Sleep(smallLatency);
        mouse_event(MOUSEEVENTF_LEFTDOWN | MOUSEEVENTF_LEFTUP, 575, 73, 0, 0);
        Thread.Sleep(smallLatency);

        hWnd = FindWindow(null, "Сохранение");
        Thread.Sleep(smallLatency);

        if (hWnd != IntPtr.Zero)
        {
            keybd_event(VK_D, 0, KEYEVENTF_EXTENDEDKEY, IntPtr.Zero); // Press D key
            keybd_event(VK_D, 0, KEYEVENTF_EXTENDEDKEY | KEYEVENTF_KEYUP, IntPtr.Zero);
            keybd_event(VK_D, 0, KEYEVENTF_EXTENDEDKEY, IntPtr.Zero); // Press D key
            keybd_event(VK_D, 0, KEYEVENTF_EXTENDEDKEY | KEYEVENTF_KEYUP, IntPtr.Zero);
            SetCursorPos(756, 517);
            Thread.Sleep(smallLatency);
            mouse_event(MOUSEEVENTF_LEFTDOWN | MOUSEEVENTF_LEFTUP, 575, 73, 0, 0);
            SetCursorPos(508, 633);
            Thread.Sleep(smallLatency);
            mouse_event(MOUSEEVENTF_LEFTDOWN | MOUSEEVENTF_LEFTUP, 575, 73, 0, 0);
            Thread.Sleep(smallLatency);
            var save = FindWindowEx(hWnd, IntPtr.Zero, "Button", null);
            if (save != IntPtr.Zero)
            {
                PostMessage(save, WM_CLICK, IntPtr.Zero, IntPtr.Zero);
            }
        }
        else
        {
            Console.WriteLine("Окно не найдено");
        }
        application = null;
        workbook = null;
        worksheet = null;
        List<string> allContainers = null;
        try
        {
            application = new Excel.Application();

            workbook = application.Workbooks.Open("C:\\Users\\User\\Desktop\\dd.xlsx", 0, false, 5, "", "", false, Excel.XlPlatform.xlWindows, "", true, false, 0, true, false, false);
            worksheet = workbook.Sheets[1];
            int numCul = 5;
            Excel.Range usedColumn = worksheet.UsedRange.Columns[numCul];
            System.Array myValues = (System.Array)usedColumn.Cells.Value2;
            allContainers = myValues.OfType<object>().Select(o => o.ToString()).ToList();            
            application.Quit();
            foreach (var item in allContainers)
            {
                //Console.WriteLine("ALL:"+item);
            }
        }
        catch (Exception)
        {

            throw;
        }
        #endregion
        #region GetEmptyContainers
        List<string> emptyContainers = new List<string>();
        foreach (var item in allContainers)
        {
            if (!notEmptyContainers.Contains(item))
            {
                emptyContainers.Add(item);
            }
        }
        Console.WriteLine("Пустые тары:");
        foreach (var item in emptyContainers)
        {
            Console.WriteLine(item);
        }
        Console.ReadLine();
        #endregion
    }
}