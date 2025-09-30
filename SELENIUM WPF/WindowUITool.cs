using SELENIUM_WPF;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Runtime.InteropServices;
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
        /// <returns>Список ElementRecord с UI элементами окна</returns>
        public static List<ElementRecord> CaptureWindowUI(Process process, string windowName, bool makeWindowActive = false)
        {
            if (process == null)
                throw new ArgumentNullException(nameof(process));

            if (string.IsNullOrEmpty(windowName))
                throw new ArgumentException("Window name cannot be null or empty", nameof(windowName));

            try
            {
                // Получаем все дочерние окна процесса
                var childWindows = GetChildWindowsOfProcess(process);

                if (!childWindows.Any())
                {
                    Console.WriteLine($"No child windows found for process {process.ProcessName} (PID: {process.Id})");
                    return new List<ElementRecord>();
                }

                // Ищем окна с нужным именем
                var matchingWindows = childWindows
                    .Where(w => w.WindowTitle.Contains(windowName, StringComparison.OrdinalIgnoreCase))
                    .ToList();

                if (!matchingWindows.Any())
                {
                    Console.WriteLine($"No windows found with name containing '{windowName}'");
                    Console.WriteLine("Available windows:");
                    foreach (var window in childWindows.Take(10)) // Показываем первые 10 для примера
                    {
                        Console.WriteLine($"  - '{window.WindowTitle}' (Handle: {window.Handle})");
                    }
                    return new List<ElementRecord>();
                }

                // Берем последнее созданное окно (самое новое)
                var targetWindow = matchingWindows
                    .OrderByDescending(w => w.CreationTime)
                    .First();

                Console.WriteLine($"Found target window: '{targetWindow.WindowTitle}' (Handle: {targetWindow.Handle})");

                // Активируем окно, если требуется
                if (makeWindowActive)
                {
                    ActivateWindow(targetWindow.Handle);
                }

                // Сканируем UI элементы с помощью вашего UI_Scanner
                return ScanWindowUI(targetWindow.Element);
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error in CaptureWindowUI: {ex.Message}");
                Console.WriteLine($"Stack trace: {ex.StackTrace}");
                return new List<ElementRecord>();
            }
        }

        /// <summary>
        /// Получает все дочерние окна указанного процесса
        /// </summary>
        private static List<WindowInfo> GetChildWindowsOfProcess(Process process)
        {
            var windows = new List<WindowInfo>();
            var processId = (uint)process.Id;

            try
            {
                // Сначала получаем все окна верхнего уровня принадлежащие процессу
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
                    catch (Exception ex)
                    {
                        Console.WriteLine($"Error processing window {hWnd}: {ex.Message}");
                    }
                    return true;
                }, IntPtr.Zero);

                Console.WriteLine($"Found {topLevelWindows.Count} top-level windows for process {process.ProcessName}");

                // Для каждого окна верхнего уровня получаем его дочерние окна
                foreach (var parentWindow in topLevelWindows)
                {
                    // Добавляем само родительское окно
                    ProcessWindow(parentWindow, windows);

                    // Добавляем дочерние окна
                    EnumChildWindows(parentWindow, (hWnd, lParam) =>
                    {
                        try
                        {
                            if (IsWindow(hWnd) && IsWindowVisible(hWnd))
                            {
                                ProcessWindow(hWnd, windows);
                            }
                        }
                        catch (Exception ex)
                        {
                            Console.WriteLine($"Error processing child window {hWnd}: {ex.Message}");
                        }
                        return true;
                    }, IntPtr.Zero);
                }

                Console.WriteLine($"Total windows found: {windows.Count}");
                return windows;
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error in GetChildWindowsOfProcess: {ex.Message}");
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

                // Пропускаем окна без заголовка или с системными заголовками
                if (string.IsNullOrWhiteSpace(windowTitle) ||
                    windowTitle.Length < 2 ||
                    windowTitle.StartsWith("Default IME") ||
                    windowTitle.StartsWith("MSCTFIME UI"))
                {
                    return;
                }

                // КРИТИЧЕСКИ ВАЖНО: создаем AutomationElement с обработкой ошибок
                AutomationElement element = null;
                System.Windows.Rect bounds = System.Windows.Rect.Empty;

                try
                {
                    element = AutomationElement.FromHandle(hWnd);
                    if (element != null)
                    {
                        // Получаем границы через UI Automation
                        bounds = element.Current.BoundingRectangle;

                        // Проверяем, что element действительно работает
                        var _ = element.Current.Name; // Это вызовет исключение если element не валиден
                    }
                }
                catch (ElementNotAvailableException)
                {
                    Console.WriteLine($"Window {hWnd} ({windowTitle}) not available for UI Automation");
                    return;
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"Cannot create AutomationElement for window {hWnd} ({windowTitle}): {ex.Message}");
                    return;
                }

                if (element == null)
                {
                    Console.WriteLine($"Failed to create AutomationElement for window {hWnd} ({windowTitle})");
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

                Console.WriteLine($"Added window: '{windowTitle}' (Handle: {hWnd})");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error processing window {hWnd}: {ex.Message}");
            }
        }

        /// <summary>
        /// Активирует окно (выводит на передний план)
        /// </summary>
        private static void ActivateWindow(IntPtr hWnd)
        {
            try
            {
                // Если окно свернуто, восстанавливаем его
                if (IsIconic(hWnd))
                {
                    ShowWindow(hWnd, SW_RESTORE);
                }
                else
                {
                    ShowWindow(hWnd, SW_SHOW);
                }

                // Выводим на передний план
                SetForegroundWindow(hWnd);

                Console.WriteLine("Window activated and brought to foreground");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error activating window: {ex.Message}");
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
            // Handle окна обычно увеличивается со временем
            // Используем это для приблизительной сортировки по времени создания
            return DateTime.Now.AddMilliseconds(-(Environment.TickCount - Math.Abs(hWnd.ToInt64() % 1000000)));
        }

        /// <summary>
        /// Сканирует UI элементы окна с помощью UI_Scanner
        /// </summary>
        private static List<ElementRecord> ScanWindowUI(AutomationElement windowElement)
        {
            if (windowElement == null)
            {
                Console.WriteLine("Window element is null, cannot scan UI");
                return new List<ElementRecord>();
            }

            try
            {
                Console.WriteLine("Starting UI scan...");

                // ВАЖНО: передаем валидный AutomationElement в ваш UI_Scanner
                var records = UI_Scanner.SnapshotControls(windowElement);

                Console.WriteLine($"UI_Scanner.SnapshotControls returned {records?.Count ?? 0} elements");

                if (records == null)
                {
                    Console.WriteLine("UI_Scanner.SnapshotControls returned null");
                    return new List<ElementRecord>();
                }

                if (records.Count == 0)
                {
                    Console.WriteLine("No UI elements found by UI_Scanner");
                    return records;
                }

                // Оцениваем доступность элементов
                Console.WriteLine("Evaluating element availability...");
                UI_Scanner.EvaluateAvailability(windowElement, records);

                // Выводим статистику
                var interactableCount = records.Count(r => r.InteractableNow);
                Console.WriteLine($"UI scan completed:");
                Console.WriteLine($"  Total elements: {records.Count}");
                Console.WriteLine($"  Interactable elements: {interactableCount}");

                // Выводим первые несколько интерактивных элементов для отладки
                var interactableElements = records.Where(r => r.InteractableNow).Take(5);
                foreach (var element in interactableElements)
                {
                    Console.WriteLine($"  Interactive: {element.ControlType} '{element.Name}' at ({element.CenterPoint.X},{element.CenterPoint.Y})");
                }

                return records;
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error scanning UI: {ex.Message}");
                Console.WriteLine($"Stack trace: {ex.StackTrace}");
                return new List<ElementRecord>();
            }
        }

      
    }

    
    
}