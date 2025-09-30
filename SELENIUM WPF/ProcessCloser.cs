using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading.Tasks;

namespace SELENIUM_WPF
{
    public static class ProcessCloser
    {
        private const uint WM_CLOSE = 0x0010;
        private const uint CTRL_C_EVENT = 0;

        private delegate bool EnumWindowsProc(IntPtr hWnd, IntPtr lParam);
        private delegate bool ConsoleCtrlDelegate(uint ctrlType);

        [DllImport("user32.dll")]
        private static extern bool EnumWindows(EnumWindowsProc lpEnumFunc, IntPtr lParam);

        [DllImport("user32.dll", SetLastError = true)]
        private static extern uint GetWindowThreadProcessId(IntPtr hWnd, out uint lpdwProcessId);

        [DllImport("user32.dll", SetLastError = true)]
        private static extern bool PostMessage(IntPtr hWnd, uint Msg, IntPtr wParam, IntPtr lParam);

        [DllImport("kernel32.dll", SetLastError = true)]
        private static extern bool AttachConsole(uint dwProcessId);

        [DllImport("kernel32.dll", SetLastError = true)]
        private static extern bool FreeConsole();

        [DllImport("kernel32.dll")]
        private static extern bool SetConsoleCtrlHandler(ConsoleCtrlDelegate handler, bool add);

        [DllImport("kernel32.dll", SetLastError = true)]
        private static extern bool GenerateConsoleCtrlEvent(uint dwCtrlEvent, uint dwProcessGroupId);

        /// <summary>
        /// Пытается аккуратно завершить процесс. 
        /// </summary>
        /// <param name="proc">Процесс для закрытия.</param>
        /// <param name="timeoutMs">Таймаут в миллисекундах.</param>
        public static void CloseProcessGracefully(Process proc, int timeoutMs)
        {
            if (proc == null)
            {
                //Debug.WriteLine("ProcessCloser: передан null вместо Process");
                return;
            }

            try
            {

                if (proc.HasExited)
                {
                    //Debug.WriteLine($"ProcessCloser: процесс {proc.Id} уже завершен");
                    return;
                }

                // Собираем все окна процесса
                List<IntPtr> windows = null;
                try
                {
                    windows = GetWindowsForProcess(proc.Id);
                }
                catch (Exception ex)
                {
                    //Debug.WriteLine($"ProcessCloser: ошибка при перечислении окон процесса {proc.Id}: {ex}");
                    windows = new List<IntPtr>();
                }

                // Если это GUI-процесс с несколькими окнами
                if (windows.Count > 1)
                {
                    foreach (var hWnd in windows)
                    {
                        try
                        {
                            PostMessage(hWnd, WM_CLOSE, IntPtr.Zero, IntPtr.Zero);
                        }
                        catch (Exception ex)
                        {
                            //Debug.WriteLine($"ProcessCloser: не удалось послать WM_CLOSE окну {hWnd}: {ex}");
                        }
                    }
                }
                // Одинокое окно — CloseMainWindow
                else if (windows.Count == 1)
                {
                    try
                    {
                        proc.CloseMainWindow();
                    }
                    catch (Exception ex)
                    {
                        //Debug.WriteLine($"ProcessCloser: CloseMainWindow не сработал для процесса {proc.Id}: {ex}");
                    }
                }
                // Ни одного окна — пробуем Ctrl+C
                else
                {
                    bool ctrlSent = false;
                    try
                    {
                        ctrlSent = TrySendCtrlC(proc.Id);
                    }
                    catch (Exception ex)
                    {
                        //Debug.WriteLine($"ProcessCloser: ошибка отправки Ctrl+C процессу {proc.Id}: {ex}");
                    }

                    if (ctrlSent)
                    {
                        try
                        {
                            proc.WaitForExit(timeoutMs);
                        }
                        catch (Exception ex)
                        {
                            //Debug.WriteLine($"ProcessCloser: WaitForExit прервано для {proc.Id}: {ex}");
                        }

                        if (proc.HasExited)
                            return;
                    }

                    // если все попытки не помогли — принудим процесс завершиться
                    try
                    {
                        if (!proc.HasExited)
                            proc.Kill();
                    }
                    catch (Exception ex)
                    {
                        //Debug.WriteLine($"ProcessCloser: Kill не сработал для {proc.Id}: {ex}");
                    }

                    return;
                }

                // Ждём завершения GUI-процесса
                try
                {
                    proc.WaitForExit(timeoutMs);
                }
                catch (Exception ex)
                {
                    //Debug.WriteLine($"ProcessCloser: ожидание завершения процесса {proc.Id} упало: {ex}");
                }

                if (!proc.HasExited)
                {
                    try
                    {
                        proc.Kill();
                    }
                    catch (Exception ex)
                    {
                        //Debug.WriteLine($"ProcessCloser: финальное Kill не сработало для {proc.Id}: {ex}");
                    }
                }
            }
            catch (Exception ex)
            {
                // Перехватываем всё, чтобы код не упал
                //Debug.WriteLine($"ProcessCloser: неожиданное исключение: {ex}");
            }
        }

        private static List<IntPtr> GetWindowsForProcess(int pid)
        {
            var result = new List<IntPtr>();
            try
            {
                EnumWindows((hWnd, _) =>
                {
                    try
                    {
                        GetWindowThreadProcessId(hWnd, out uint windowPid);
                        if (windowPid == pid)
                            result.Add(hWnd);
                    }
                    catch
                    {
                        // проигнорировать ошибки получения PID окна
                    }
                    return true;
                }, IntPtr.Zero);
            }
            catch (Exception ex)
            {
                //Debug.WriteLine($"ProcessCloser: EnumWindows упал для PID {pid}: {ex}");
            }
            return result;
        }

        private static bool TrySendCtrlC(int pid)
        {
            try
            {
                if (!AttachConsole((uint)pid))
                {
                    //Debug.WriteLine($"ProcessCloser: AttachConsole({pid}) вернул false");
                    return false;
                }

                // Отключаем свой обработчик, чтобы не поймать Ctrl+C сами
                SetConsoleCtrlHandler(null, true);

                bool sent = GenerateConsoleCtrlEvent(CTRL_C_EVENT, 0);
                Thread.Sleep(200);

                FreeConsole();
                SetConsoleCtrlHandler(null, false);

                return sent;
            }
            catch (Exception ex)
            {
                //Debug.WriteLine($"ProcessCloser: TrySendCtrlC упал для PID {pid}: {ex}");
                return false;
            }
        }
    }
}
