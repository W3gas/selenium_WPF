using SELENIUM_WPF;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Runtime.InteropServices;
using System.Windows;
using System.Windows.Automation;

namespace SELENIUM_WPF
{
    public class WindowInfo
    {
        public IntPtr Handle { get; set; }
        public string WindowTitle { get; set; }
        public DateTime CreationTime { get; set; }
        public AutomationElement Element { get; set; }
        public System.Windows.Rect Bounds { get; set; }
    }

    public static class WindowUITool
    {
        #region Win32 API

        [DllImport("user32.dll")]
        private static extern bool EnumChildWindows(IntPtr hWndParent, EnumWindowsProc lpEnumFunc, IntPtr lParam);

        [DllImport("user32.dll")]
        private static extern bool EnumWindows(EnumWindowsProc enumProc, IntPtr lParam);

        [DllImport("user32.dll")]
        private static extern uint GetWindowThreadProcessId(IntPtr hWnd, out uint lpdwProcessId);

        [DllImport("user32.dll", CharSet = CharSet.Unicode)]
        private static extern int GetWindowText(IntPtr hWnd, System.Text.StringBuilder lpString, int nMaxCount);

        [DllImport("user32.dll")]
        private static extern int GetWindowTextLength(IntPtr hWnd);

        [DllImport("user32.dll")]
        private static extern bool IsWindow(IntPtr hWnd);

        [DllImport("user32.dll")]
        private static extern bool IsWindowVisible(IntPtr hWnd);

        [DllImport("user32.dll")]
        private static extern bool SetForegroundWindow(IntPtr hWnd);

        [DllImport("user32.dll")]
        private static extern bool ShowWindow(IntPtr hWnd, int nCmdShow);

        [DllImport("user32.dll")]
        private static extern bool IsIconic(IntPtr hWnd);

        private delegate bool EnumWindowsProc(IntPtr hWnd, IntPtr lParam);

        private const int SW_RESTORE = 9;
        private const int SW_SHOW = 5;

        #endregion


        /// <summary>
        /// Получает дочерние окна процесса, находит нужное по имени и сканирует его UI элементы
        /// </summary>
        /// <param name="process">Процесс для поиска дочерних окон</param>
        /// <param name="windowName">Имя окна для поиска</param>
        /// <param name="makeWindowActive">Активировать найденное окно</param>
        /// <returns>Список ElementRecord с UI элементами окна (всегда возвращает список, даже пустой)</returns>
        public static List<ElementRecord> CaptureWindowUI(Process process, string windowName, bool makeWindowActive = true)
        {
            try
            {
                // Валидация входных параметров
                if (process == null)
                {
                    ShowError("Процесс не может быть null", "Ошибка параметров");
                    return new List<ElementRecord>();
                }

                if (string.IsNullOrEmpty(windowName))
                {
                    ShowError("Имя окна не может быть пустым", "Ошибка параметров");
                    return new List<ElementRecord>();
                }

                // Проверка доступности процесса
                if (process.HasExited)
                {
                    ShowError($"Процесс {process.ProcessName} (PID: {process.Id}) уже завершен", "Ошибка процесса");
                    return new List<ElementRecord>();
                }

                //  все дочерние окна процесса
                var childWindows = GetChildWindowsOfProcess(process);

                if (childWindows == null || !childWindows.Any())
                {
                    ShowError(
                        $"Не найдено ни одного дочернего окна для процесса:\n{process.ProcessName} (PID: {process.Id})",
                        "Окна не найдены"
                    );
                    return new List<ElementRecord>();
                }

                //  окна с нужным именем
                var matchingWindows = childWindows
                    .Where(w => w.WindowTitle.Contains(windowName, StringComparison.OrdinalIgnoreCase))
                    .ToList();

                if (!matchingWindows.Any())
                {
                    var availableWindows = string.Join("\n",
                        childWindows.Take(10).Select(w => $"• {w.WindowTitle}"));

                    ShowError(
                        $"Не найдено окно с именем: '{windowName}'\n\n",
                        "Окно не найдено"
                    );
                    return new List<ElementRecord>();
                }

              
                var targetWindow = matchingWindows
                    .OrderByDescending(w => w.Handle.ToInt64())
                    .First();

                
                if (makeWindowActive)
                {
                    if (!ActivateWindow(targetWindow.Handle))
                    {
                        ShowWarning(
                            $"Не удалось активировать окно:\n'{targetWindow.WindowTitle}'\n\nПродолжаем без активации...",
                            "Предупреждение"
                        );
                    }
                }

               
                return ScanWindowUI(targetWindow.Element, targetWindow.WindowTitle);
            }
            catch (InvalidOperationException ex)
            {
                ShowError(
                    $"Ошибка доступа к процессу:\n{ex.Message}",
                    "Ошибка операции"
                );
                return new List<ElementRecord>();
            }
            catch (UnauthorizedAccessException ex)
            {
                ShowError(
                    $"Недостаточно прав для доступа к процессу:\n{ex.Message}",
                    "Ошибка доступа"
                );
                return new List<ElementRecord>();
            }
            catch (Exception ex)
            {
                ShowError(
                    $"Непредвиденная ошибка при захвате UI окна:\n{ex.Message}",
                    "Критическая ошибка"
                );
                return new List<ElementRecord>();
            }
        }


        /// <summary>
        /// Получает все дочерние окна указанного процесса
        /// </summary>
        private static List<WindowInfo> GetChildWindowsOfProcess(Process process)
        {
            var windows = new List<WindowInfo>();

            try
            {
                var processId = (uint)process.Id;
                var topLevelWindows = new List<IntPtr>();

               
                EnumWindows((hWnd, lParam) =>
                {
                    try
                    {
                        GetWindowThreadProcessId(hWnd, out uint windowProcessId);
                        if (windowProcessId == processId && IsWindowVisible(hWnd))
                        {
                            topLevelWindows.Add(hWnd);
                        }
                    }
                    catch
                    {
                        
                    }
                    return true;
                }, IntPtr.Zero);

                
                foreach (var parentWindow in topLevelWindows)
                {
                    try
                    {
                       
                        ProcessWindow(parentWindow, windows);

                       
                        EnumChildWindows(parentWindow, (hWnd, lParam) =>
                        {
                            try
                            {
                                if (IsWindow(hWnd) && IsWindowVisible(hWnd))
                                {
                                    ProcessWindow(hWnd, windows);
                                }
                            }
                            catch
                            {
                               
                            }
                            return true;
                        }, IntPtr.Zero);
                    }
                    catch
                    {
                        
                        continue;
                    }
                }

                return windows;
            }
            catch (Exception ex)
            {
                ShowError(
                    $"Ошибка при получении списка окон процесса:\n{ex.Message}",
                    "Ошибка перечисления окон"
                );
                return windows;
            }
        }


        /// <summary>
        /// Обрабатывает одно окно и добавляет его в список если подходит
        /// </summary>
        private static void ProcessWindow(IntPtr hWnd, List<WindowInfo> windows)
        {
            try
            {
                var windowTitle = GetWindowTitle(hWnd);

               
                if (string.IsNullOrWhiteSpace(windowTitle) ||
                    windowTitle.Length < 2 ||
                    windowTitle.StartsWith("Default IME") ||
                    windowTitle.StartsWith("MSCTFIME UI"))
                {
                    return;
                }

               
                AutomationElement element = null;
                System.Windows.Rect bounds = System.Windows.Rect.Empty;

                try
                {
                    element = AutomationElement.FromHandle(hWnd);
                    if (element != null)
                    {
                        bounds = element.Current.BoundingRectangle;
                        
                        var _ = element.Current.Name;
                    }
                    else
                    {
                        return;
                    }
                }
                catch (ElementNotAvailableException)
                {
                    return;
                }
                catch
                {
                    return;
                }

                windows.Add(new WindowInfo
                {
                    Handle = hWnd,
                    WindowTitle = windowTitle,
                    CreationTime = GetWindowCreationTime(hWnd),
                    Element = element,
                    Bounds = bounds
                });
            }
            catch
            {
               
            }
        }


        /// <summary>
        /// Активирует окно (выводит на передний план)
        /// </summary>
        private static bool ActivateWindow(IntPtr hWnd)
        {
            try
            {
                // Если окно свернуто
                if (IsIconic(hWnd))
                {
                    if (!ShowWindow(hWnd, SW_RESTORE))
                        return false;
                }
                else
                {
                    if (!ShowWindow(hWnd, SW_SHOW))
                        return false;
                }

              
                return SetForegroundWindow(hWnd);
            }
            catch
            {
                return false;
            }
        }


        /// <summary>
        /// Получает заголовок окна
        /// </summary>
        private static string GetWindowTitle(IntPtr hWnd)
        {
            try
            {
                int length = GetWindowTextLength(hWnd);
                if (length == 0) return string.Empty;

                var builder = new System.Text.StringBuilder(length + 1);
                GetWindowText(hWnd, builder, builder.Capacity);
                return builder.ToString();
            }
            catch
            {
                return string.Empty;
            }
        }


        /// <summary>
        /// Приблизительное время создания окна
        /// </summary>
        private static DateTime GetWindowCreationTime(IntPtr hWnd)
        {
            try
            {
                return DateTime.Now.AddMilliseconds(-(Environment.TickCount - Math.Abs(hWnd.ToInt64() % 1000000)));
            }
            catch
            {
                return DateTime.Now;
            }
        }


        /// <summary>
        /// Сканирует UI элементы окна с помощью UI_Scanner
        /// </summary>
        private static List<ElementRecord> ScanWindowUI(AutomationElement windowElement, string windowTitle)
        {
            if (windowElement == null)
            {
                ShowError(
                    "Элемент окна недоступен для сканирования UI",
                    "Ошибка UI Automation"
                );
                return new List<ElementRecord>();
            }

            try
            {
                
                var records = UI_Scanner.SnapshotControls(windowElement);

                if (records == null)
                {
                    ShowWarning(
                        $"UI_Scanner не вернул элементы для окна:\n'{windowTitle}'",
                        "Предупреждение сканирования"
                    );
                    return new List<ElementRecord>();
                }

                if (records.Count == 0)
                {
                    ShowInfo(
                        $"В окне '{windowTitle}' не найдено UI элементов",
                        "Информация"
                    );
                    return records;
                }

               
                try
                {
                    UI_Scanner.EvaluateAvailability(windowElement, records);
                }
                catch (Exception ex)
                {
                    ShowWarning(
                        $"Ошибка при оценке доступности элементов:\n{ex.Message}\n\nПродолжаем с базовыми данными...",
                        "Предупреждение"
                    );
                }

                return records;
            }
            catch (ElementNotAvailableException)
            {
                ShowError(
                    $"Окно '{windowTitle}' стало недоступным во время сканирования",
                    "Элемент недоступен"
                );
                return new List<ElementRecord>();
            }
            catch (Exception ex)
            {
                ShowError(
                    $"Ошибка при сканировании UI окна '{windowTitle}':\n{ex.Message}",
                    "Ошибка сканирования"
                );
                return new List<ElementRecord>();
            }
        }



        #region Методы отображения сообщений


        private static void ShowError(string message, string title)
        {
            System.Windows.MessageBox.Show(
                messageBoxText: $"\t{message}",
                caption: title,
                button: MessageBoxButton.OK,
                icon: MessageBoxImage.Error
            );
        }


        private static void ShowWarning(string message, string title)
        {
            System.Windows.MessageBox.Show(
                messageBoxText: $"\t{message}",
                caption: title,
                button: MessageBoxButton.OK,
                icon: MessageBoxImage.Warning
            );
        }


        private static void ShowInfo(string message, string title)
        {
            System.Windows.MessageBox.Show(
                messageBoxText: $"\t{message}",
                caption: title,
                button: MessageBoxButton.OK,
                icon: MessageBoxImage.Information
            );
        }

        #endregion
    }
}