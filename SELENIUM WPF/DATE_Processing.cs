using Microsoft.Win32;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Diagnostics.CodeAnalysis;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using System.Windows;
using System.Windows.Automation;
using System.Windows.Controls;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Threading;
namespace SELENIUM_WPF
{
    public partial class MainWindow : Window
    {

        [DllImport("user32.dll")]
        private static extern bool SetForegroundWindow(IntPtr hWnd);

        [DllImport("user32.dll")]
        private static extern bool ShowWindow(IntPtr hWnd, int nCmdShow);

        private const int SW_RESTORE = 9;


        //  поиск контрола по тексту внутри его
        public ElementRecord Find_By_Text(string name)
        {
            for (int i = 0; i < DATE.Count; i++)
            {
                if (DATE[i].Name == name)
                {
                    return DATE[i];
                }
            }

            return  null;
        }


        //  поиск контрола по имени контрола в приложении(как к нему обращается другое приложение в коде)
        public ElementRecord Find_By_Name(string name)
        {
            for (int i = 0; i < DATE.Count; i++)
            {
                if (DATE[i].AutomationId == name)
                {
                    return DATE[i];
                }
            }

            return null;
        }




        public static bool Activate_Main_Process_W(Process process)
        {
            if (process == null || process.HasExited)
                return false;

            IntPtr hWnd = process.MainWindowHandle;
            if (hWnd == IntPtr.Zero)
                return false;

            ShowWindow(hWnd, SW_RESTORE);       // восстановить, если свернуто
            return SetForegroundWindow(hWnd);   // активировать окно
        }

    }
}
