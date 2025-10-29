using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading.Tasks;

namespace SELENIUM_WPF
{
    public class Mouse_Emulator
    {
        [StructLayout(LayoutKind.Sequential)]
        struct INPUT
        {
            public uint type;
            public MOUSEINPUT mi;
        }

        [StructLayout(LayoutKind.Sequential)]
        struct MOUSEINPUT
        {
            public int dx;
            public int dy;
            public uint mouseData;
            public uint dwFlags;
            public uint time;
            public IntPtr dwExtraInfo;
        }

        const uint INPUT_MOUSE = 0;
        const uint MOUSEEVENTF_LEFTDOWN = 0x0002;
        const uint MOUSEEVENTF_LEFTUP = 0x0004;
        const uint MOUSEEVENTF_RIGHTDOWN = 0x0008;
        const uint MOUSEEVENTF_RIGHTUP = 0x0010;
        const uint MOUSEEVENTF_MIDDLEDOWN = 0x0020;
        const uint MOUSEEVENTF_MIDDLEUP = 0x0040;
        const uint MOUSEEVENTF_MOVE = 0x0001;
        const uint MOUSEEVENTF_ABSOLUTE = 0x8000;

        [DllImport("user32.dll", SetLastError = true)]
        static extern uint SendInput(uint nInputs, INPUT[] pInputs, int cbSize);

        [DllImport("user32.dll")]
        static extern int GetSystemMetrics(int nIndex);

        const int SM_CXSCREEN = 0; // Ширина экрана
        const int SM_CYSCREEN = 1; // Высота экрана

        public enum MouseButton
        {
            Left,
            Right,
            Middle
        }

        public static void MouseClick(MouseButton button = MouseButton.Left)
        {
            uint downFlag = 0, upFlag = 0;
            switch (button)
            {
                case MouseButton.Left:
                    downFlag = MOUSEEVENTF_LEFTDOWN;
                    upFlag = MOUSEEVENTF_LEFTUP;
                    break;
                case MouseButton.Right:
                    downFlag = MOUSEEVENTF_RIGHTDOWN;
                    upFlag = MOUSEEVENTF_RIGHTUP;
                    break;
                case MouseButton.Middle:
                    downFlag = MOUSEEVENTF_MIDDLEDOWN;
                    upFlag = MOUSEEVENTF_MIDDLEUP;
                    break;
            }

            INPUT[] inputs = new INPUT[2];
            // Нажатие
            inputs[0].type = INPUT_MOUSE;
            inputs[0].mi.dwFlags = downFlag;
            // Отпускание
            inputs[1].type = INPUT_MOUSE;
            inputs[1].mi.dwFlags = upFlag;

            SendInput((uint)inputs.Length, inputs, Marshal.SizeOf(typeof(INPUT)));
        }




        public static void MouseDown(MouseButton button = MouseButton.Left)
        {
            uint downFlag = 0, upFlag = 0;
            switch (button)
            {
                case MouseButton.Left:
                    downFlag = MOUSEEVENTF_LEFTDOWN;
                    upFlag = MOUSEEVENTF_LEFTUP;
                    break;
                case MouseButton.Right:
                    downFlag = MOUSEEVENTF_RIGHTDOWN;
                    upFlag = MOUSEEVENTF_RIGHTUP;
                    break;
                case MouseButton.Middle:
                    downFlag = MOUSEEVENTF_MIDDLEDOWN;
                    upFlag = MOUSEEVENTF_MIDDLEUP;
                    break;
            }

            INPUT[] inputs = new INPUT[1];
            // Нажатие
            inputs[0].type = INPUT_MOUSE;
            inputs[0].mi.dwFlags = downFlag;
            

            SendInput((uint)inputs.Length, inputs, Marshal.SizeOf(typeof(INPUT)));
        }



        public static void MouseUp(MouseButton button = MouseButton.Left)
        {
            uint downFlag = 0, upFlag = 0;
            switch (button)
            {
                case MouseButton.Left:
                    downFlag = MOUSEEVENTF_LEFTDOWN;
                    upFlag = MOUSEEVENTF_LEFTUP;
                    break;
                case MouseButton.Right:
                    downFlag = MOUSEEVENTF_RIGHTDOWN;
                    upFlag = MOUSEEVENTF_RIGHTUP;
                    break;
                case MouseButton.Middle:
                    downFlag = MOUSEEVENTF_MIDDLEDOWN;
                    upFlag = MOUSEEVENTF_MIDDLEUP;
                    break;
            }

            INPUT[] inputs = new INPUT[1];
            
            // Отпускание
            inputs[0].type = INPUT_MOUSE;
            inputs[0].mi.dwFlags = upFlag;

            SendInput((uint)inputs.Length, inputs, Marshal.SizeOf(typeof(INPUT)));
        }



        /// <summary>
        /// Перемещает курсор мыши в указанную позицию на экране
        /// </summary>
        /// <param name="x">Координата X (горизонтальная позиция в пикселях)</param>
        /// <param name="y">Координата Y (вертикальная позиция в пикселях)</param>
        public static void MoveCursor(int x, int y)
        {
            // Получаем размеры экрана
            int screenWidth = GetSystemMetrics(SM_CXSCREEN);
            int screenHeight = GetSystemMetrics(SM_CYSCREEN);

            // Преобразуем координаты в формат Windows (0-65535)
            // Windows использует нормализованные координаты для абсолютного позиционирования
            int normalizedX = (x * 65535) / screenWidth;
            int normalizedY = (y * 65535) / screenHeight;

            INPUT[] inputs = new INPUT[1];
            inputs[0].type = INPUT_MOUSE;
            inputs[0].mi.dx = normalizedX;
            inputs[0].mi.dy = normalizedY;
            inputs[0].mi.dwFlags = MOUSEEVENTF_MOVE | MOUSEEVENTF_ABSOLUTE;
            inputs[0].mi.mouseData = 0;
            inputs[0].mi.time = 0;
            inputs[0].mi.dwExtraInfo = IntPtr.Zero;

            SendInput((uint)inputs.Length, inputs, Marshal.SizeOf(typeof(INPUT)));
        }

        /// <summary>
        /// Комбинированная функция: перемещает курсор и делает клик
        /// </summary>
        /// <param name="x">Координата X</param>
        /// <param name="y">Координата Y</param>
        /// <param name="button">Кнопка мыши для клика</param>
        public static void MoveAndClick(int x, int y, MouseButton button = MouseButton.Left)
        {
            MoveCursor(x, y);
            System.Threading.Thread.Sleep(10); // Небольшая задержка для надежности
            MouseClick(button);
        }
    }




public static class Keyboard_Emulator
    {
        // Правильная структура INPUT с union
        [StructLayout(LayoutKind.Sequential)]
        struct INPUT
        {
            public uint type;
            public INPUTUNION U;
        }

        [StructLayout(LayoutKind.Explicit)]
        struct INPUTUNION
        {
            [FieldOffset(0)] public MOUSEINPUT mi;
            [FieldOffset(0)] public KEYBDINPUT ki;
            [FieldOffset(0)] public HARDWAREINPUT hi;
        }

        [StructLayout(LayoutKind.Sequential)]
        struct MOUSEINPUT
        {
            public int dx;
            public int dy;
            public uint mouseData;
            public uint dwFlags;
            public uint time;
            public UIntPtr dwExtraInfo;
        }

        [StructLayout(LayoutKind.Sequential)]
        struct KEYBDINPUT
        {
            public ushort wVk;
            public ushort wScan;
            public uint dwFlags;
            public uint time;
            public UIntPtr dwExtraInfo;
        }

        [StructLayout(LayoutKind.Sequential)]
        struct HARDWAREINPUT
        {
            public uint uMsg;
            public ushort wParamL;
            public ushort wParamH;
        }

        const uint INPUT_KEYBOARD = 1;
        const uint KEYEVENTF_EXTENDEDKEY = 0x0001;
        const uint KEYEVENTF_KEYUP = 0x0002;
        const uint KEYEVENTF_UNICODE = 0x0004;
        const uint KEYEVENTF_SCANCODE = 0x0008;
        const uint MAPVK_VK_TO_VSC = 0x0;

        [DllImport("user32.dll", SetLastError = true)]
        static extern uint SendInput(uint nInputs, INPUT[] pInputs, int cbSize);

        [DllImport("user32.dll")]
        static extern short VkKeyScan(char ch);

        [DllImport("user32.dll")]
        static extern short VkKeyScanEx(char ch, IntPtr dwhkl);

        [DllImport("user32.dll")]
        static extern IntPtr GetKeyboardLayout(uint idThread);

        [DllImport("user32.dll")]
        static extern uint MapVirtualKeyEx(uint uCode, uint uMapType, IntPtr dwhkl);

        // Специальные клавиши
        public const byte VK_CANCEL = 0x03;
        public const byte VK_BACK = 0x08;
        public const byte VK_TAB = 0x09;
        public const byte VK_CLEAR = 0x0C;
        public const byte VK_RETURN = 0x0D;  //  ЕНТЕР
        public const byte VK_SHIFT = 0x10;
        public const byte VK_CONTROL = 0x11;
        public const byte VK_MENU = 0x12;
        public const byte VK_PAUSE = 0x13;
        public const byte VK_CAPITAL = 0x14;
        public const byte VK_ESCAPE = 0x1B;
        public const byte VK_SPACE = 0x20;

        // Клавиши навигации
        public const byte VK_PRIOR = 0x21;
        public const byte VK_NEXT = 0x22;
        public const byte VK_END = 0x23;
        public const byte VK_HOME = 0x24;
        public const byte VK_LEFT = 0x25;
        public const byte VK_UP = 0x26;
        public const byte VK_RIGHT = 0x27;
        public const byte VK_DOWN = 0x28;
        public const byte VK_SELECT = 0x29;
        public const byte VK_PRINT = 0x2A;
        public const byte VK_EXECUTE = 0x2B;
        public const byte VK_SNAPSHOT = 0x2C;
        public const byte VK_INSERT = 0x2D;
        public const byte VK_DELETE = 0x2E;
        public const byte VK_HELP = 0x2F;

        // Цифры
        public const byte VK_0 = 0x30;
        public const byte VK_1 = 0x31;
        public const byte VK_2 = 0x32;
        public const byte VK_3 = 0x33;
        public const byte VK_4 = 0x34;
        public const byte VK_5 = 0x35;
        public const byte VK_6 = 0x36;
        public const byte VK_7 = 0x37;
        public const byte VK_8 = 0x38;
        public const byte VK_9 = 0x39;

        // Буквы
        public const byte VK_A = 0x41;
        public const byte VK_B = 0x42;
        public const byte VK_C = 0x43;
        public const byte VK_D = 0x44;
        public const byte VK_E = 0x45;
        public const byte VK_F = 0x46;
        public const byte VK_G = 0x47;
        public const byte VK_H = 0x48;
        public const byte VK_I = 0x49;
        public const byte VK_J = 0x4A;
        public const byte VK_K = 0x4B;
        public const byte VK_L = 0x4C;
        public const byte VK_M = 0x4D;
        public const byte VK_N = 0x4E;
        public const byte VK_O = 0x4F;
        public const byte VK_P = 0x50;
        public const byte VK_Q = 0x51;
        public const byte VK_R = 0x52;
        public const byte VK_S = 0x53;
        public const byte VK_T = 0x54;
        public const byte VK_U = 0x55;
        public const byte VK_V = 0x56;
        public const byte VK_W = 0x57;
        public const byte VK_X = 0x58;
        public const byte VK_Y = 0x59;
        public const byte VK_Z = 0x5A;

        // Windows клавиши
        public const byte VK_LWIN = 0x5B;
        public const byte VK_RWIN = 0x5C;
        public const byte VK_APPS = 0x5D;

        // Цифровая клавиатура
        public const byte VK_NUMPAD0 = 0x60;
        public const byte VK_NUMPAD1 = 0x61;
        public const byte VK_NUMPAD2 = 0x62;
        public const byte VK_NUMPAD3 = 0x63;
        public const byte VK_NUMPAD4 = 0x64;
        public const byte VK_NUMPAD5 = 0x65;
        public const byte VK_NUMPAD6 = 0x66;
        public const byte VK_NUMPAD7 = 0x67;
        public const byte VK_NUMPAD8 = 0x68;
        public const byte VK_NUMPAD9 = 0x69;
        public const byte VK_MULTIPLY = 0x6A;
        public const byte VK_ADD = 0x6B;
        public const byte VK_SEPARATOR = 0x6C;
        public const byte VK_SUBTRACT = 0x6D;
        public const byte VK_DECIMAL = 0x6E;
        public const byte VK_DIVIDE = 0x6F;

        // Функциональные клавиши
        public const byte VK_F1 = 0x70;
        public const byte VK_F2 = 0x71;
        public const byte VK_F3 = 0x72;
        public const byte VK_F4 = 0x73;
        public const byte VK_F5 = 0x74;
        public const byte VK_F6 = 0x75;
        public const byte VK_F7 = 0x76;
        public const byte VK_F8 = 0x77;
        public const byte VK_F9 = 0x78;
        public const byte VK_F10 = 0x79;
        public const byte VK_F11 = 0x7A;
        public const byte VK_F12 = 0x7B;
        public const byte VK_F13 = 0x7C;
        public const byte VK_F14 = 0x7D;
        public const byte VK_F15 = 0x7E;
        public const byte VK_F16 = 0x7F;
        public const byte VK_F17 = 0x80;
        public const byte VK_F18 = 0x81;
        public const byte VK_F19 = 0x82;
        public const byte VK_F20 = 0x83;
        public const byte VK_F21 = 0x84;
        public const byte VK_F22 = 0x85;
        public const byte VK_F23 = 0x86;
        public const byte VK_F24 = 0x87;

        // Lock клавиши
        public const byte VK_NUMLOCK = 0x90;
        public const byte VK_SCROLL = 0x91;

        // Модификаторы
        public const byte VK_LSHIFT = 0xA0;
        public const byte VK_RSHIFT = 0xA1;
        public const byte VK_LCONTROL = 0xA2;
        public const byte VK_RCONTROL = 0xA3;
        public const byte VK_LMENU = 0xA4;
        public const byte VK_RMENU = 0xA5;

        // Браузер и мультимедиа
        public const byte VK_BROWSER_BACK = 0xA6;
        public const byte VK_BROWSER_FORWARD = 0xA7;
        public const byte VK_BROWSER_REFRESH = 0xA8;
        public const byte VK_BROWSER_STOP = 0xA9;
        public const byte VK_BROWSER_SEARCH = 0xAA;
        public const byte VK_BROWSER_FAVORITES = 0xAB;
        public const byte VK_BROWSER_HOME = 0xAC;
        public const byte VK_VOLUME_MUTE = 0xAD;
        public const byte VK_VOLUME_DOWN = 0xAE;
        public const byte VK_VOLUME_UP = 0xAF;
        public const byte VK_MEDIA_NEXT_TRACK = 0xB0;
        public const byte VK_MEDIA_PREV_TRACK = 0xB1;
        public const byte VK_MEDIA_STOP = 0xB2;
        public const byte VK_MEDIA_PLAY_PAUSE = 0xB3;
        public const byte VK_LAUNCH_MAIL = 0xB4;
        public const byte VK_LAUNCH_MEDIA_SELECT = 0xB5;
        public const byte VK_LAUNCH_APP1 = 0xB6;
        public const byte VK_LAUNCH_APP2 = 0xB7;

        // OEM клавиши
        public const byte VK_OEM_1 = 0xBA;
        public const byte VK_OEM_PLUS = 0xBB;
        public const byte VK_OEM_COMMA = 0xBC;
        public const byte VK_OEM_MINUS = 0xBD;
        public const byte VK_OEM_PERIOD = 0xBE;
        public const byte VK_OEM_2 = 0xBF;
        public const byte VK_OEM_3 = 0xC0;
        public const byte VK_OEM_4 = 0xDB;
        public const byte VK_OEM_5 = 0xDC;
        public const byte VK_OEM_6 = 0xDD;
        public const byte VK_OEM_7 = 0xDE;
        public const byte VK_OEM_8 = 0xDF;
        public const byte VK_OEM_102 = 0xE2;

        // Специальные коды
        public const byte VK_PROCESSKEY = 0xE5;
        public const byte VK_PACKET = 0xE7;
        public const byte VK_ATTN = 0xF6;
        public const byte VK_CRSEL = 0xF7;
        public const byte VK_EXSEL = 0xF8;
        public const byte VK_EREOF = 0xF9;
        public const byte VK_PLAY = 0xFA;
        public const byte VK_ZOOM = 0xFB;
        public const byte VK_NONAME = 0xFC;
        public const byte VK_PA1 = 0xFD;
        public const byte VK_OEM_CLEAR = 0xFE;

        // Список расширенных клавиш
        private static readonly HashSet<byte> ExtendedKeys = new HashSet<byte>
        {
            VK_UP, VK_DOWN, VK_LEFT, VK_RIGHT,
            VK_HOME, VK_END, VK_PRIOR, VK_NEXT,
            VK_INSERT, VK_DELETE,
            VK_RCONTROL, VK_RMENU,
            VK_DIVIDE, VK_NUMLOCK
        };

        // Вспомогательный метод для получения скан-кода
        private static ushort VkToScan(byte vk)
        {
            var hkl = GetKeyboardLayout(0);
            return (ushort)MapVirtualKeyEx(vk, MAPVK_VK_TO_VSC, hkl);
        }

        /// <summary>
        /// Нажать клавишу (без отпускания)
        /// </summary>
        public static void KeyDown(byte keyCode)
        {
            ushort scan = VkToScan(keyCode);
            uint flags = KEYEVENTF_SCANCODE;
            if (ExtendedKeys.Contains(keyCode))
                flags |= KEYEVENTF_EXTENDEDKEY;

            INPUT[] inputs = new INPUT[1];
            inputs[0].type = INPUT_KEYBOARD;
            inputs[0].U = new INPUTUNION
            {
                ki = new KEYBDINPUT
                {
                    wVk = 0,
                    wScan = scan,
                    dwFlags = flags,
                    time = 0,
                    dwExtraInfo = UIntPtr.Zero
                }
            };

            SendInput((uint)inputs.Length, inputs, Marshal.SizeOf(typeof(INPUT)));
        }

        /// <summary>
        /// Отпустить клавишу
        /// </summary>
        public static void KeyUp(byte keyCode)
        {
            ushort scan = VkToScan(keyCode);
            uint flags = KEYEVENTF_SCANCODE | KEYEVENTF_KEYUP;
            if (ExtendedKeys.Contains(keyCode))
                flags |= KEYEVENTF_EXTENDEDKEY;

            INPUT[] inputs = new INPUT[1];
            inputs[0].type = INPUT_KEYBOARD;
            inputs[0].U = new INPUTUNION
            {
                ki = new KEYBDINPUT
                {
                    wVk = 0,
                    wScan = scan,
                    dwFlags = flags,
                    time = 0,
                    dwExtraInfo = UIntPtr.Zero
                }
            };

            SendInput((uint)inputs.Length, inputs, Marshal.SizeOf(typeof(INPUT)));
        }

        /// <summary>
        /// Нажать и отпустить клавишу
        /// </summary>
        public static void KeyPress(byte keyCode, int delayMs = 10)
        {
            KeyDown(keyCode);
            Thread.Sleep(delayMs);
            KeyUp(keyCode);
        }

        /// <summary>
        /// Нажать комбинацию клавиш (например: Ctrl+C)
        /// </summary>
        public static void KeyCombination(params byte[] keyCodes)
        {
            // Нажимаем все клавиши по порядку
            foreach (var key in keyCodes)
            {
                KeyDown(key);
                Thread.Sleep(10);
            }

            Thread.Sleep(80);

            // Отпускаем в обратном порядке
            for (int i = keyCodes.Length - 1; i >= 0; i--)
            {
                KeyUp(keyCodes[i]);
                Thread.Sleep(10);
            }
        }

        /// <summary>
        /// Ввести текст посимвольно (работает с любой раскладкой через Unicode)
        /// </summary>
        public static void TypeText(string text, int delayMs = 10)
        {
            foreach (char c in text)
            {
                TypeCharUnicode(c);
                Thread.Sleep(delayMs);
            }
        }

        /// <summary>
        /// Ввести один символ через Unicode (работает с любой раскладкой)
        /// </summary>
        private static void TypeCharUnicode(char character)
        {
            INPUT[] inputs = new INPUT[2];

            // Нажатие
            inputs[0].type = INPUT_KEYBOARD;
            inputs[0].U = new INPUTUNION
            {
                ki = new KEYBDINPUT
                {
                    wVk = 0,
                    wScan = character,
                    dwFlags = KEYEVENTF_UNICODE,
                    time = 0,
                    dwExtraInfo = UIntPtr.Zero
                }
            };

            // Отпускание
            inputs[1].type = INPUT_KEYBOARD;
            inputs[1].U = new INPUTUNION
            {
                ki = new KEYBDINPUT
                {
                    wVk = 0,
                    wScan = character,
                    dwFlags = KEYEVENTF_UNICODE | KEYEVENTF_KEYUP,
                    time = 0,
                    dwExtraInfo = UIntPtr.Zero
                }
            };

            SendInput((uint)inputs.Length, inputs, Marshal.SizeOf(typeof(INPUT)));
        }

        /// <summary>
        /// Переключить раскладку клавиатуры (Alt+Shift или Ctrl+Shift в зависимости от настроек системы)
        /// </summary>
        public static void SwitchKeyboardLayout()
        {
            // Пробуем Alt+Shift (стандартная комбинация для переключения раскладки)
            KeyCombination(VK_LMENU, VK_SHIFT);
            Thread.Sleep(100); // Даём время системе переключить раскладку
        }

        /// <summary>
        /// Ввести текст с использованием виртуальных клавиш и автоматическим переключением раскладки
        /// </summary>
        public static void TypeTextVK(string text, int delayMs = 10)
        {
            foreach (char c in text)
            {
                short vkAndShift = VkKeyScan(c);
                int attemptCount = 0;
                const int maxAttempts = 10; // Максимум 10 попыток (чтобы не зациклиться)

                // Цикл поиска нужной раскладки
                while (vkAndShift == -1 && attemptCount < maxAttempts)
                {
                    // Символ не найден в текущей раскладке - переключаем раскладку
                    SwitchKeyboardLayout();

                    // Проверяем снова
                    vkAndShift = VkKeyScan(c);
                    attemptCount++;
                }

                // Если после всех попыток символ не найден - используем Unicode метод
                if (vkAndShift == -1)
                {
                    TypeCharUnicode(c);
                }
                else
                {
                    // Извлекаем виртуальный код (младший байт)
                    byte vk = (byte)(vkAndShift & 0xFF);

                    // Извлекаем состояние модификаторов (старший байт)
                    byte shiftState = (byte)(vkAndShift >> 8);

                    // Проверяем какие модификаторы нужны
                    bool needShift = (shiftState & 1) != 0;  // Бит 0
                    bool needCtrl = (shiftState & 2) != 0;   // Бит 1
                    bool needAlt = (shiftState & 4) != 0;    // Бит 2

                    // Нажимаем модификаторы
                    if (needShift) KeyDown(VK_SHIFT);
                    if (needCtrl) KeyDown(VK_CONTROL);
                    if (needAlt) KeyDown(VK_MENU);

                    // Нажимаем саму клавишу
                    KeyPress(vk);

                    // Отпускаем модификаторы (в обратном порядке)
                    if (needAlt) KeyUp(VK_MENU);
                    if (needCtrl) KeyUp(VK_CONTROL);
                    if (needShift) KeyUp(VK_SHIFT);
                }

                Thread.Sleep(delayMs);
            }
        }





        
        public static readonly Dictionary<string, byte> VirtualKeyMap = new Dictionary<string, byte>
        {
            { nameof(VK_CANCEL), VK_CANCEL },
            { nameof(VK_BACK), VK_BACK },
            { nameof(VK_TAB), VK_TAB },
            { nameof(VK_CLEAR), VK_CLEAR },
            { nameof(VK_RETURN), VK_RETURN },
            { nameof(VK_SHIFT), VK_SHIFT },
            { nameof(VK_CONTROL), VK_CONTROL },
            { nameof(VK_MENU), VK_MENU },
            { nameof(VK_PAUSE), VK_PAUSE },
            { nameof(VK_CAPITAL), VK_CAPITAL },
            { nameof(VK_ESCAPE), VK_ESCAPE },
            { nameof(VK_SPACE), VK_SPACE },

            { nameof(VK_PRIOR), VK_PRIOR },
            { nameof(VK_NEXT), VK_NEXT },
            { nameof(VK_END), VK_END },
            { nameof(VK_HOME), VK_HOME },
            { nameof(VK_LEFT), VK_LEFT },
            { nameof(VK_UP), VK_UP },
            { nameof(VK_RIGHT), VK_RIGHT },
            { nameof(VK_DOWN), VK_DOWN },
            { nameof(VK_SELECT), VK_SELECT },
            { nameof(VK_PRINT), VK_PRINT },
            { nameof(VK_EXECUTE), VK_EXECUTE },
            { nameof(VK_SNAPSHOT), VK_SNAPSHOT },
            { nameof(VK_INSERT), VK_INSERT },
            { nameof(VK_DELETE), VK_DELETE },
            { nameof(VK_HELP), VK_HELP },

            { nameof(VK_0), VK_0 },
            { nameof(VK_1), VK_1 },
            { nameof(VK_2), VK_2 },
            { nameof(VK_3), VK_3 },
            { nameof(VK_4), VK_4 },
            { nameof(VK_5), VK_5 },
            { nameof(VK_6), VK_6 },
            { nameof(VK_7), VK_7 },
            { nameof(VK_8), VK_8 },
            { nameof(VK_9), VK_9 },

            { nameof(VK_A), VK_A },
            { nameof(VK_B), VK_B },
            { nameof(VK_C), VK_C },
            { nameof(VK_D), VK_D },
            { nameof(VK_E), VK_E },
            { nameof(VK_F), VK_F },
            { nameof(VK_G), VK_G },
            { nameof(VK_H), VK_H },
            { nameof(VK_I), VK_I },
            { nameof(VK_J), VK_J },
            { nameof(VK_K), VK_K },
            { nameof(VK_L), VK_L },
            { nameof(VK_M), VK_M },
            { nameof(VK_N), VK_N },
            { nameof(VK_O), VK_O },
            { nameof(VK_P), VK_P },
            { nameof(VK_Q), VK_Q },
            { nameof(VK_R), VK_R },
            { nameof(VK_S), VK_S },
            { nameof(VK_T), VK_T },
            { nameof(VK_U), VK_U },
            { nameof(VK_V), VK_V },
            { nameof(VK_W), VK_W },
            { nameof(VK_X), VK_X },
            { nameof(VK_Y), VK_Y },
            { nameof(VK_Z), VK_Z },

            { nameof(VK_LWIN), VK_LWIN },
            { nameof(VK_RWIN), VK_RWIN },
            { nameof(VK_APPS), VK_APPS },

            { nameof(VK_NUMPAD0), VK_NUMPAD0 },
            { nameof(VK_NUMPAD1), VK_NUMPAD1 },
            { nameof(VK_NUMPAD2), VK_NUMPAD2 },
            { nameof(VK_NUMPAD3), VK_NUMPAD3 },
            { nameof(VK_NUMPAD4), VK_NUMPAD4 },
            { nameof(VK_NUMPAD5), VK_NUMPAD5 },
            { nameof(VK_NUMPAD6), VK_NUMPAD6 },
            { nameof(VK_NUMPAD7), VK_NUMPAD7 },
            { nameof(VK_NUMPAD8), VK_NUMPAD8 },
            { nameof(VK_NUMPAD9), VK_NUMPAD9 },
            { nameof(VK_MULTIPLY), VK_MULTIPLY },
            { nameof(VK_ADD), VK_ADD },
            { nameof(VK_SEPARATOR), VK_SEPARATOR },
            { nameof(VK_SUBTRACT), VK_SUBTRACT },
            { nameof(VK_DECIMAL), VK_DECIMAL },
            { nameof(VK_DIVIDE), VK_DIVIDE },

            { nameof(VK_F1), VK_F1 },
            { nameof(VK_F2), VK_F2 },
            { nameof(VK_F3), VK_F3 },
            { nameof(VK_F4), VK_F4 },
            { nameof(VK_F5), VK_F5 },
            { nameof(VK_F6), VK_F6 },
            { nameof(VK_F7), VK_F7 },
            { nameof(VK_F8), VK_F8 },
            { nameof(VK_F9), VK_F9 },
            { nameof(VK_F10), VK_F10 },
            { nameof(VK_F11), VK_F11 },
            { nameof(VK_F12), VK_F12 },
            { nameof(VK_F13), VK_F13 },
            { nameof(VK_F14), VK_F14 },
            { nameof(VK_F15), VK_F15 },
            { nameof(VK_F16), VK_F16 },
            { nameof(VK_F17), VK_F17 },
            { nameof(VK_F18), VK_F18 },
            { nameof(VK_F19), VK_F19 },
            { nameof(VK_F20), VK_F20 },
            { nameof(VK_F21), VK_F21 },
            { nameof(VK_F22), VK_F22 },
            { nameof(VK_F23), VK_F23 },
            { nameof(VK_F24), VK_F24 },

            { nameof(VK_NUMLOCK), VK_NUMLOCK },
            { nameof(VK_SCROLL), VK_SCROLL },

            { nameof(VK_LSHIFT), VK_LSHIFT },
            { nameof(VK_RSHIFT), VK_RSHIFT },
            { nameof(VK_LCONTROL), VK_LCONTROL },
            { nameof(VK_RCONTROL), VK_RCONTROL },
            { nameof(VK_LMENU), VK_LMENU },
            { nameof(VK_RMENU), VK_RMENU },

            { nameof(VK_BROWSER_BACK), VK_BROWSER_BACK },
            { nameof(VK_BROWSER_FORWARD), VK_BROWSER_FORWARD },
            { nameof(VK_BROWSER_REFRESH), VK_BROWSER_REFRESH },
            { nameof(VK_BROWSER_STOP), VK_BROWSER_STOP },
            { nameof(VK_BROWSER_SEARCH), VK_BROWSER_SEARCH },
            { nameof(VK_BROWSER_FAVORITES), VK_BROWSER_FAVORITES },
            { nameof(VK_BROWSER_HOME), VK_BROWSER_HOME },
            { nameof(VK_VOLUME_MUTE), VK_VOLUME_MUTE },
            { nameof(VK_VOLUME_DOWN), VK_VOLUME_DOWN },
            { nameof(VK_VOLUME_UP), VK_VOLUME_UP },
            { nameof(VK_MEDIA_NEXT_TRACK), VK_MEDIA_NEXT_TRACK },
            { nameof(VK_MEDIA_PREV_TRACK), VK_MEDIA_PREV_TRACK },
            { nameof(VK_MEDIA_STOP), VK_MEDIA_STOP },
            { nameof(VK_MEDIA_PLAY_PAUSE), VK_MEDIA_PLAY_PAUSE },
            { nameof(VK_LAUNCH_MAIL), VK_LAUNCH_MAIL },
            { nameof(VK_LAUNCH_MEDIA_SELECT), VK_LAUNCH_MEDIA_SELECT },
            { nameof(VK_LAUNCH_APP1), VK_LAUNCH_APP1 },
            { nameof(VK_LAUNCH_APP2), VK_LAUNCH_APP2 },

            { nameof(VK_OEM_1), VK_OEM_1 },
            { nameof(VK_OEM_PLUS), VK_OEM_PLUS },
            { nameof(VK_OEM_COMMA), VK_OEM_COMMA },
            { nameof(VK_OEM_MINUS), VK_OEM_MINUS },
            { nameof(VK_OEM_PERIOD), VK_OEM_PERIOD },
            { nameof(VK_OEM_2), VK_OEM_2 },
            { nameof(VK_OEM_3), VK_OEM_3 },
            { nameof(VK_OEM_4), VK_OEM_4 },
            { nameof(VK_OEM_5), VK_OEM_5 },
            { nameof(VK_OEM_6), VK_OEM_6 },
            { nameof(VK_OEM_7), VK_OEM_7 },
            { nameof(VK_OEM_8), VK_OEM_8 },
            { nameof(VK_OEM_102), VK_OEM_102 },

            { nameof(VK_PROCESSKEY), VK_PROCESSKEY },
            { nameof(VK_PACKET), VK_PACKET },
            { nameof(VK_ATTN), VK_ATTN },
            { nameof(VK_CRSEL), VK_CRSEL },
            { nameof(VK_EXSEL), VK_EXSEL },
            { nameof(VK_EREOF), VK_EREOF },
            { nameof(VK_PLAY), VK_PLAY },
            { nameof(VK_ZOOM), VK_ZOOM },
            { nameof(VK_NONAME), VK_NONAME },
            { nameof(VK_PA1), VK_PA1 },
            { nameof(VK_OEM_CLEAR), VK_OEM_CLEAR }
        };




    }


}