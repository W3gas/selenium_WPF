using Microsoft.Win32;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Diagnostics.CodeAnalysis;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using System.Security.Cryptography.X509Certificates;
using System.Windows;
using System.Windows.Automation;
using System.Windows.Controls;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Threading;
namespace SELENIUM_WPF
{

    public struct Command_Text
    {
        public string Full_Command;
        public string Command_Name;
        public List<string> args;

        public Command_Text()
        {
            this.Full_Command = "";
            this.Command_Name = "";
            this.args = new List<string>();
        }

        public Command_Text(string full, string name, List<string> Args)
        {
            this.Full_Command = full;
            this.Command_Name = name;
            this.args = Args;
        }
    }


    public partial class MainWindow : Window
    {

        [DllImport("user32.dll")]
        private static extern bool SetForegroundWindow(IntPtr hWnd);

        [DllImport("user32.dll")]
        private static extern bool ShowWindow(IntPtr hWnd, int nCmdShow);

        private const int SW_RESTORE = 9;

        public List <Command_Text> command_s = new List <Command_Text>();


        









        //  поиск контрола по тексту внутри его
        public ElementRecord Find_By_Name(string name)
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
        public ElementRecord Find_By_Id(string name)
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



        //  активация главного окна процесса
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


        public void Analyze_Code()
        {
            command_s.Clear();
            string code_Input = "" + code_TB.Text + '\n';
            string command = "";

            int sz = 0;
            int line_Number = 0;

            if (code_Input.Trim() == "")
            {
                return;
            }

            while (sz < code_Input.Length)
            {
                line_Number++;
                command = "";

                while (sz < code_Input.Length && code_Input[sz] != '\n' && code_Input[sz] != ';')
                {
                    command += code_Input[sz];
                    sz++;
                }

                if (sz >= code_Input.Length)
                {
                    continue;
                }
                else
                {
                    if (code_Input[sz] == ';')
                    {
                        while (sz < code_Input.Length && code_Input[sz] != '\n')
                        {
                            sz++;
                        }
                    }

                    command = command.Trim();
                    if (command == "")
                    {
                        sz++;
                        continue;
                    }

                    int open_Index = command.IndexOf('(');
                    int close_Index = command.IndexOf(')');
                    if (open_Index == -1 || close_Index == -1 || close_Index < open_Index)
                    {
                        System.Windows.MessageBox.Show(
                            owner: this,
                            messageBoxText: "\tОшибка синтаксиса:\n отсутствуют скобки в строке с номером " + line_Number,
                            caption: "Парсер псевдокода",
                            button: MessageBoxButton.OK,
                            icon: MessageBoxImage.Error
                        );
                        return;
                    }

                    string command_Name = command.Substring(0, open_Index).Trim();                                     // сама команда
                    string arguments_Text = command.Substring(open_Index + 1, close_Index - open_Index - 1);








                    



                    if (command_Name == "TextEnter")
                    {
                        //  начало
                        if (arguments_Text.Length == 0 || arguments_Text[0] != '"')
                        {
                            System.Windows.MessageBox.Show(
                                owner: this,
                                messageBoxText: $"Текст в команде TextEnter должен начинаться с двойной кавычки (строка {line_Number})",
                                caption: "Парсер псевдокода",
                                button: MessageBoxButton.OK,
                                icon: MessageBoxImage.Error
                            );
                            return;
                        }

                        
                        string textContent = "";
                        int i = 1; 


                        //  запись текста ("" = ")
                        while (i < arguments_Text.Length)
                        {
                            char c = arguments_Text[i];

                            if (c == '"' && i + 1 < arguments_Text.Length && arguments_Text[i + 1] == '"')
                            {
                                textContent += '"';
                                i += 2;
                            }
                            else if (c == '"')
                            {
                                i++; 
                                break;
                            }
                            else
                            {
                                textContent += c;
                                i++;
                            }
                        }

                        
                        if (i > arguments_Text.Length)
                        {
                            System.Windows.MessageBox.Show(
                                owner: this,
                                messageBoxText: $"Незакрытая кавычка в команде TextEnter (строка {line_Number})",
                                caption: "Парсер псевдокода",
                                button: MessageBoxButton.OK,
                                icon: MessageBoxImage.Error
                            );
                            return;
                        }

                        // Остаток строки после закрывающей кавычки
                        string rest = "";
                        if (i < arguments_Text.Length)
                        {
                            rest = arguments_Text.Substring(i);
                        }

                        // Список дополнительных аргументов
                        List<string> otherArgs = new List<string>();

                        if (rest.Length > 0)
                        {
                            if (rest[0] != ',')
                            {
                                System.Windows.MessageBox.Show(
                                    owner: this,
                                    messageBoxText: $"После текста в TextEnter должен следовать запятая и другие аргументы (строка {line_Number})",
                                    caption: "Парсер псевдокода",
                                    button: MessageBoxButton.OK,
                                    icon: MessageBoxImage.Error
                                );
                                return;
                            }

                            
                            string argsPart = rest.Substring(1);
                            string[] rawArgs = argsPart.Split(new char[] { ',' });

                            foreach (string arg in rawArgs)
                            {
                                otherArgs.Add(arg.Trim());
                            }
                        }

                        
                        List<string> finalArguments = new List<string>();
                        finalArguments.Add(textContent);
                        finalArguments.AddRange(otherArgs);

                        
                        Command_Text ntemp = new Command_Text(command, command_Name, finalArguments);
                        command_s.Add(ntemp);

                        sz++;
                        continue;
                    }










                    List<string> argument_List = arguments_Text
                        .Split(new[] { ',' }, StringSplitOptions.RemoveEmptyEntries)
                        .Select(arg => arg.Trim())
                        .ToList();










                    if (command_Name == "MouseDown")
                    {
                        boolKeys.Is_LeftMouse_Down = true;
                    }
                    if (command_Name == "MouseRightDown")
                    {
                        boolKeys.Is_RightMouse_Down = true;
                    }
                    if (command_Name == "MouseMiddleDown")
                    {
                        boolKeys.Is_MiddleMouse_Down = true;
                    }



                    if (command_Name == "MouseUp")
                    {
                        boolKeys.Is_LeftMouse_Down = false;
                    }
                    if (command_Name == "MouseRightUp")
                    {
                        boolKeys.Is_RightMouse_Down = false;
                    }
                    if (command_Name == "MouseMiddleUp")
                    {
                        boolKeys.Is_MiddleMouse_Down = false;
                    }



                   
                    if (command_Name == "KeyDown")
                    {
                        foreach (string rawArg in argument_List)
                        {
                            string keyName = rawArg.Trim().ToUpperInvariant();

                            switch (keyName)
                            {
                                case "VK_CANCEL": boolKeys.Is_CANCEL_Down = true; break;
                                case "VK_BACK": boolKeys.Is_BACK_Down = true; break;
                                case "VK_TAB": boolKeys.Is_TAB_Down = true; break;
                                case "VK_CLEAR": boolKeys.Is_CLEAR_Down = true; break;
                                case "VK_RETURN": boolKeys.Is_RETURN_Down = true; break;
                                case "VK_SHIFT": boolKeys.Is_SHIFT_Down = true; break;
                                case "VK_CONTROL": boolKeys.Is_CONTROL_Down = true; break;
                                case "VK_MENU": boolKeys.Is_MENU_Down = true; break;
                                case "VK_PAUSE": boolKeys.Is_PAUSE_Down = true; break;
                                case "VK_CAPITAL": boolKeys.Is_CAPITAL_Down = true; break;
                                case "VK_ESCAPE": boolKeys.Is_ESCAPE_Down = true; break;
                                case "VK_SPACE": boolKeys.Is_SPACE_Down = true; break;
                                case "VK_PRIOR": boolKeys.Is_PRIOR_Down = true; break;
                                case "VK_NEXT": boolKeys.Is_NEXT_Down = true; break;
                                case "VK_END": boolKeys.Is_END_Down = true; break;
                                case "VK_HOME": boolKeys.Is_HOME_Down = true; break;
                                case "VK_LEFT": boolKeys.Is_LEFT_Down = true; break;
                                case "VK_UP": boolKeys.Is_UP_Down = true; break;
                                case "VK_RIGHT": boolKeys.Is_RIGHT_Down = true; break;
                                case "VK_DOWN": boolKeys.Is_DOWN_Down = true; break;
                                case "VK_SELECT": boolKeys.Is_SELECT_Down = true; break;
                                case "VK_PRINT": boolKeys.Is_PRINT_Down = true; break;
                                case "VK_EXECUTE": boolKeys.Is_EXECUTE_Down = true; break;
                                case "VK_SNAPSHOT": boolKeys.Is_SNAPSHOT_Down = true; break;
                                case "VK_INSERT": boolKeys.Is_INSERT_Down = true; break;
                                case "VK_DELETE": boolKeys.Is_DELETE_Down = true; break;
                                case "VK_HELP": boolKeys.Is_HELP_Down = true; break;
                                case "VK_0": boolKeys.Is_0_Down = true; break;
                                case "VK_1": boolKeys.Is_1_Down = true; break;
                                case "VK_2": boolKeys.Is_2_Down = true; break;
                                case "VK_3": boolKeys.Is_3_Down = true; break;
                                case "VK_4": boolKeys.Is_4_Down = true; break;
                                case "VK_5": boolKeys.Is_5_Down = true; break;
                                case "VK_6": boolKeys.Is_6_Down = true; break;
                                case "VK_7": boolKeys.Is_7_Down = true; break;
                                case "VK_8": boolKeys.Is_8_Down = true; break;
                                case "VK_9": boolKeys.Is_9_Down = true; break;
                                case "VK_A": boolKeys.Is_A_Down = true; break;
                                case "VK_B": boolKeys.Is_B_Down = true; break;
                                case "VK_C": boolKeys.Is_C_Down = true; break;
                                case "VK_D": boolKeys.Is_D_Down = true; break;
                                case "VK_E": boolKeys.Is_E_Down = true; break;
                                case "VK_F": boolKeys.Is_F_Down = true; break;
                                case "VK_G": boolKeys.Is_G_Down = true; break;
                                case "VK_H": boolKeys.Is_H_Down = true; break;
                                case "VK_I": boolKeys.Is_I_Down = true; break;
                                case "VK_J": boolKeys.Is_J_Down = true; break;
                                case "VK_K": boolKeys.Is_K_Down = true; break;
                                case "VK_L": boolKeys.Is_L_Down = true; break;
                                case "VK_M": boolKeys.Is_M_Down = true; break;
                                case "VK_N": boolKeys.Is_N_Down = true; break;
                                case "VK_O": boolKeys.Is_O_Down = true; break;
                                case "VK_P": boolKeys.Is_P_Down = true; break;
                                case "VK_Q": boolKeys.Is_Q_Down = true; break;
                                case "VK_R": boolKeys.Is_R_Down = true; break;
                                case "VK_S": boolKeys.Is_S_Down = true; break;
                                case "VK_T": boolKeys.Is_T_Down = true; break;
                                case "VK_U": boolKeys.Is_U_Down = true; break;
                                case "VK_V": boolKeys.Is_V_Down = true; break;
                                case "VK_W": boolKeys.Is_W_Down = true; break;
                                case "VK_X": boolKeys.Is_X_Down = true; break;
                                case "VK_Y": boolKeys.Is_Y_Down = true; break;
                                case "VK_Z": boolKeys.Is_Z_Down = true; break;
                                case "VK_LWIN": boolKeys.Is_LWIN_Down = true; break;
                                case "VK_RWIN": boolKeys.Is_RWIN_Down = true; break;
                                case "VK_APPS": boolKeys.Is_APPS_Down = true; break;
                                case "VK_NUMPAD0": boolKeys.Is_NUMPAD0_Down = true; break;
                                case "VK_NUMPAD1": boolKeys.Is_NUMPAD1_Down = true; break;
                                case "VK_NUMPAD2": boolKeys.Is_NUMPAD2_Down = true; break;
                                case "VK_NUMPAD3": boolKeys.Is_NUMPAD3_Down = true; break;
                                case "VK_NUMPAD4": boolKeys.Is_NUMPAD4_Down = true; break;
                                case "VK_NUMPAD5": boolKeys.Is_NUMPAD5_Down = true; break;
                                case "VK_NUMPAD6": boolKeys.Is_NUMPAD6_Down = true; break;
                                case "VK_NUMPAD7": boolKeys.Is_NUMPAD7_Down = true; break;
                                case "VK_NUMPAD8": boolKeys.Is_NUMPAD8_Down = true; break;
                                case "VK_NUMPAD9": boolKeys.Is_NUMPAD9_Down = true; break;
                                case "VK_MULTIPLY": boolKeys.Is_MULTIPLY_Down = true; break;
                                case "VK_ADD": boolKeys.Is_ADD_Down = true; break;
                                case "VK_SEPARATOR": boolKeys.Is_SEPARATOR_Down = true; break;
                                case "VK_SUBTRACT": boolKeys.Is_SUBTRACT_Down = true; break;
                                case "VK_DECIMAL": boolKeys.Is_DECIMAL_Down = true; break;
                                case "VK_DIVIDE": boolKeys.Is_DIVIDE_Down = true; break;
                                case "VK_F1": boolKeys.Is_F1_Down = true; break;
                                case "VK_F2": boolKeys.Is_F2_Down = true; break;
                                case "VK_F3": boolKeys.Is_F3_Down = true; break;
                                case "VK_F4": boolKeys.Is_F4_Down = true; break;
                                case "VK_F5": boolKeys.Is_F5_Down = true; break;
                                case "VK_F6": boolKeys.Is_F6_Down = true; break;
                                case "VK_F7": boolKeys.Is_F7_Down = true; break;
                                case "VK_F8": boolKeys.Is_F8_Down = true; break;
                                case "VK_F9": boolKeys.Is_F9_Down = true; break;
                                case "VK_F10": boolKeys.Is_F10_Down = true; break;
                                case "VK_F11": boolKeys.Is_F11_Down = true; break;
                                case "VK_F12": boolKeys.Is_F12_Down = true; break;
                                case "VK_F13": boolKeys.Is_F13_Down = true; break;
                                case "VK_F14": boolKeys.Is_F14_Down = true; break;
                                case "VK_F15": boolKeys.Is_F15_Down = true; break;
                                case "VK_F16": boolKeys.Is_F16_Down = true; break;
                                case "VK_F17": boolKeys.Is_F17_Down = true; break;
                                case "VK_F18": boolKeys.Is_F18_Down = true; break;
                                case "VK_F19": boolKeys.Is_F19_Down = true; break;
                                case "VK_F20": boolKeys.Is_F20_Down = true; break;
                                case "VK_F21": boolKeys.Is_F21_Down = true; break;
                                case "VK_F22": boolKeys.Is_F22_Down = true; break;
                                case "VK_F23": boolKeys.Is_F23_Down = true; break;
                                case "VK_F24": boolKeys.Is_F24_Down = true; break;
                                case "VK_NUMLOCK": boolKeys.Is_NUMLOCK_Down = true; break;
                                case "VK_SCROLL": boolKeys.Is_SCROLL_Down = true; break;
                                case "VK_LSHIFT": boolKeys.Is_LSHIFT_Down = true; break;
                                case "VK_RSHIFT": boolKeys.Is_RSHIFT_Down = true; break;
                                case "VK_LCONTROL": boolKeys.Is_LCONTROL_Down = true; break;
                                case "VK_RCONTROL": boolKeys.Is_RCONTROL_Down = true; break;
                                case "VK_LMENU": boolKeys.Is_LMENU_Down = true; break;
                                case "VK_RMENU": boolKeys.Is_RMENU_Down = true; break;
                                case "VK_BROWSER_BACK": boolKeys.Is_BROWSER_BACK_Down = true; break;
                                case "VK_BROWSER_FORWARD": boolKeys.Is_BROWSER_FORWARD_Down = true; break;
                                case "VK_BROWSER_REFRESH": boolKeys.Is_BROWSER_REFRESH_Down = true; break;
                                case "VK_BROWSER_STOP": boolKeys.Is_BROWSER_STOP_Down = true; break;
                                case "VK_BROWSER_SEARCH": boolKeys.Is_BROWSER_SEARCH_Down = true; break;
                                case "VK_BROWSER_FAVORITES": boolKeys.Is_BROWSER_FAVORITES_Down = true; break;
                                case "VK_BROWSER_HOME": boolKeys.Is_BROWSER_HOME_Down = true; break;
                                case "VK_VOLUME_MUTE": boolKeys.Is_VOLUME_MUTE_Down = true; break;
                                case "VK_VOLUME_DOWN": boolKeys.Is_VOLUME_DOWN_Down = true; break;
                                case "VK_VOLUME_UP": boolKeys.Is_VOLUME_UP_Down = true; break;
                                case "VK_MEDIA_NEXT_TRACK": boolKeys.Is_MEDIA_NEXT_TRACK_Down = true; break;
                                case "VK_MEDIA_PREV_TRACK": boolKeys.Is_MEDIA_PREV_TRACK_Down = true; break;
                                case "VK_MEDIA_STOP": boolKeys.Is_MEDIA_STOP_Down = true; break;
                                case "VK_MEDIA_PLAY_PAUSE": boolKeys.Is_MEDIA_PLAY_PAUSE_Down = true; break;
                                case "VK_LAUNCH_MAIL": boolKeys.Is_LAUNCH_MAIL_Down = true; break;
                                case "VK_LAUNCH_MEDIA_SELECT": boolKeys.Is_LAUNCH_MEDIA_SELECT_Down = true; break;
                                case "VK_LAUNCH_APP1": boolKeys.Is_LAUNCH_APP1_Down = true; break;
                                case "VK_LAUNCH_APP2": boolKeys.Is_LAUNCH_APP2_Down = true; break;
                                case "VK_OEM_1": boolKeys.Is_OEM_1_Down = true; break;
                                case "VK_OEM_PLUS": boolKeys.Is_OEM_PLUS_Down = true; break;
                                case "VK_OEM_COMMA": boolKeys.Is_OEM_COMMA_Down = true; break;
                                case "VK_OEM_MINUS": boolKeys.Is_OEM_MINUS_Down = true; break;
                                case "VK_OEM_PERIOD": boolKeys.Is_OEM_PERIOD_Down = true; break;
                                case "VK_OEM_2": boolKeys.Is_OEM_2_Down = true; break;
                                case "VK_OEM_3": boolKeys.Is_OEM_3_Down = true; break;
                                case "VK_OEM_4": boolKeys.Is_OEM_4_Down = true; break;
                                case "VK_OEM_5": boolKeys.Is_OEM_5_Down = true; break;
                                case "VK_OEM_6": boolKeys.Is_OEM_6_Down = true; break;
                                case "VK_OEM_7": boolKeys.Is_OEM_7_Down = true; break;
                                case "VK_OEM_8": boolKeys.Is_OEM_8_Down = true; break;
                                case "VK_OEM_102": boolKeys.Is_OEM_102_Down = true; break;
                                case "VK_PROCESSKEY": boolKeys.Is_PROCESSKEY_Down = true; break;
                                case "VK_PACKET": boolKeys.Is_PACKET_Down = true; break;
                                case "VK_ATTN": boolKeys.Is_ATTN_Down = true; break;
                                case "VK_CRSEL": boolKeys.Is_CRSEL_Down = true; break;
                                case "VK_EXSEL": boolKeys.Is_EXSEL_Down = true; break;
                                case "VK_EREOF": boolKeys.Is_EREOF_Down = true; break;
                                case "VK_PLAY": boolKeys.Is_PLAY_Down = true; break;
                                case "VK_ZOOM": boolKeys.Is_ZOOM_Down = true; break;
                                case "VK_NONAME": boolKeys.Is_NONAME_Down = true; break;
                                case "VK_PA1": boolKeys.Is_PA1_Down = true; break;
                                case "VK_OEM_CLEAR": boolKeys.Is_OEM_CLEAR_Down = true; break;
                            }
                        }
                    }

                    
                    else if (command_Name == "KeyUp")
                    {
                        foreach (string rawArg in argument_List)
                        {
                            string keyName = rawArg.Trim().ToUpperInvariant();

                            switch (keyName)
                            {
                                case "VK_CANCEL": boolKeys.Is_CANCEL_Down = false; break;
                                case "VK_BACK": boolKeys.Is_BACK_Down = false; break;
                                case "VK_TAB": boolKeys.Is_TAB_Down = false; break;
                                case "VK_CLEAR": boolKeys.Is_CLEAR_Down = false; break;
                                case "VK_RETURN": boolKeys.Is_RETURN_Down = false; break;
                                case "VK_SHIFT": boolKeys.Is_SHIFT_Down = false; break;
                                case "VK_CONTROL": boolKeys.Is_CONTROL_Down = false; break;
                                case "VK_MENU": boolKeys.Is_MENU_Down = false; break;
                                case "VK_PAUSE": boolKeys.Is_PAUSE_Down = false; break;
                                case "VK_CAPITAL": boolKeys.Is_CAPITAL_Down = false; break;
                                case "VK_ESCAPE": boolKeys.Is_ESCAPE_Down = false; break;
                                case "VK_SPACE": boolKeys.Is_SPACE_Down = false; break;
                                case "VK_PRIOR": boolKeys.Is_PRIOR_Down = false; break;
                                case "VK_NEXT": boolKeys.Is_NEXT_Down = false; break;
                                case "VK_END": boolKeys.Is_END_Down = false; break;
                                case "VK_HOME": boolKeys.Is_HOME_Down = false; break;
                                case "VK_LEFT": boolKeys.Is_LEFT_Down = false; break;
                                case "VK_UP": boolKeys.Is_UP_Down = false; break;
                                case "VK_RIGHT": boolKeys.Is_RIGHT_Down = false; break;
                                case "VK_DOWN": boolKeys.Is_DOWN_Down = false; break;
                                case "VK_SELECT": boolKeys.Is_SELECT_Down = false; break;
                                case "VK_PRINT": boolKeys.Is_PRINT_Down = false; break;
                                case "VK_EXECUTE": boolKeys.Is_EXECUTE_Down = false; break;
                                case "VK_SNAPSHOT": boolKeys.Is_SNAPSHOT_Down = false; break;
                                case "VK_INSERT": boolKeys.Is_INSERT_Down = false; break;
                                case "VK_DELETE": boolKeys.Is_DELETE_Down = false; break;
                                case "VK_HELP": boolKeys.Is_HELP_Down = false; break;
                                case "VK_0": boolKeys.Is_0_Down = false; break;
                                case "VK_1": boolKeys.Is_1_Down = false; break;
                                case "VK_2": boolKeys.Is_2_Down = false; break;
                                case "VK_3": boolKeys.Is_3_Down = false; break;
                                case "VK_4": boolKeys.Is_4_Down = false; break;
                                case "VK_5": boolKeys.Is_5_Down = false; break;
                                case "VK_6": boolKeys.Is_6_Down = false; break;
                                case "VK_7": boolKeys.Is_7_Down = false; break;
                                case "VK_8": boolKeys.Is_8_Down = false; break;
                                case "VK_9": boolKeys.Is_9_Down = false; break;
                                case "VK_A": boolKeys.Is_A_Down = false; break;
                                case "VK_B": boolKeys.Is_B_Down = false; break;
                                case "VK_C": boolKeys.Is_C_Down = false; break;
                                case "VK_D": boolKeys.Is_D_Down = false; break;
                                case "VK_E": boolKeys.Is_E_Down = false; break;
                                case "VK_F": boolKeys.Is_F_Down = false; break;
                                case "VK_G": boolKeys.Is_G_Down = false; break;
                                case "VK_H": boolKeys.Is_H_Down = false; break;
                                case "VK_I": boolKeys.Is_I_Down = false; break;
                                case "VK_J": boolKeys.Is_J_Down = false; break;
                                case "VK_K": boolKeys.Is_K_Down = false; break;
                                case "VK_L": boolKeys.Is_L_Down = false; break;
                                case "VK_M": boolKeys.Is_M_Down = false; break;
                                case "VK_N": boolKeys.Is_N_Down = false; break;
                                case "VK_O": boolKeys.Is_O_Down = false; break;
                                case "VK_P": boolKeys.Is_P_Down = false; break;
                                case "VK_Q": boolKeys.Is_Q_Down = false; break;
                                case "VK_R": boolKeys.Is_R_Down = false; break;
                                case "VK_S": boolKeys.Is_S_Down = false; break;
                                case "VK_T": boolKeys.Is_T_Down = false; break;
                                case "VK_U": boolKeys.Is_U_Down = false; break;
                                case "VK_V": boolKeys.Is_V_Down = false; break;
                                case "VK_W": boolKeys.Is_W_Down = false; break;
                                case "VK_X": boolKeys.Is_X_Down = false; break;
                                case "VK_Y": boolKeys.Is_Y_Down = false; break;
                                case "VK_Z": boolKeys.Is_Z_Down = false; break;
                                case "VK_LWIN": boolKeys.Is_LWIN_Down = false; break;
                                case "VK_RWIN": boolKeys.Is_RWIN_Down = false; break;
                                case "VK_APPS": boolKeys.Is_APPS_Down = false; break;
                                case "VK_NUMPAD0": boolKeys.Is_NUMPAD0_Down = false; break;
                                case "VK_NUMPAD1": boolKeys.Is_NUMPAD1_Down = false; break;
                                case "VK_NUMPAD2": boolKeys.Is_NUMPAD2_Down = false; break;
                                case "VK_NUMPAD3": boolKeys.Is_NUMPAD3_Down = false; break;
                                case "VK_NUMPAD4": boolKeys.Is_NUMPAD4_Down = false; break;
                                case "VK_NUMPAD5": boolKeys.Is_NUMPAD5_Down = false; break;
                                case "VK_NUMPAD6": boolKeys.Is_NUMPAD6_Down = false; break;
                                case "VK_NUMPAD7": boolKeys.Is_NUMPAD7_Down = false; break;
                                case "VK_NUMPAD8": boolKeys.Is_NUMPAD8_Down = false; break;
                                case "VK_NUMPAD9": boolKeys.Is_NUMPAD9_Down = false; break;
                                case "VK_MULTIPLY": boolKeys.Is_MULTIPLY_Down = false; break;
                                case "VK_ADD": boolKeys.Is_ADD_Down = false; break;
                                case "VK_SEPARATOR": boolKeys.Is_SEPARATOR_Down = false; break;
                                case "VK_SUBTRACT": boolKeys.Is_SUBTRACT_Down = false; break;
                                case "VK_DECIMAL": boolKeys.Is_DECIMAL_Down = false; break;
                                case "VK_DIVIDE": boolKeys.Is_DIVIDE_Down = false; break;
                                case "VK_F1": boolKeys.Is_F1_Down = false; break;
                                case "VK_F2": boolKeys.Is_F2_Down = false; break;
                                case "VK_F3": boolKeys.Is_F3_Down = false; break;
                                case "VK_F4": boolKeys.Is_F4_Down = false; break;
                                case "VK_F5": boolKeys.Is_F5_Down = false; break;
                                case "VK_F6": boolKeys.Is_F6_Down = false; break;
                                case "VK_F7": boolKeys.Is_F7_Down = false; break;
                                case "VK_F8": boolKeys.Is_F8_Down = false; break;
                                case "VK_F9": boolKeys.Is_F9_Down = false; break;
                                case "VK_F10": boolKeys.Is_F10_Down = false; break;
                                case "VK_F11": boolKeys.Is_F11_Down = false; break;
                                case "VK_F12": boolKeys.Is_F12_Down = false; break;
                                case "VK_F13": boolKeys.Is_F13_Down = false; break;
                                case "VK_F14": boolKeys.Is_F14_Down = false; break;
                                case "VK_F15": boolKeys.Is_F15_Down = false; break;
                                case "VK_F16": boolKeys.Is_F16_Down = false; break;
                                case "VK_F17": boolKeys.Is_F17_Down = false; break;
                                case "VK_F18": boolKeys.Is_F18_Down = false; break;
                                case "VK_F19": boolKeys.Is_F19_Down = false; break;
                                case "VK_F20": boolKeys.Is_F20_Down = false; break;
                                case "VK_F21": boolKeys.Is_F21_Down = false; break;
                                case "VK_F22": boolKeys.Is_F22_Down = false; break;
                                case "VK_F23": boolKeys.Is_F23_Down = false; break;
                                case "VK_F24": boolKeys.Is_F24_Down = false; break;
                                case "VK_NUMLOCK": boolKeys.Is_NUMLOCK_Down = false; break;
                                case "VK_SCROLL": boolKeys.Is_SCROLL_Down = false; break;
                                case "VK_LSHIFT": boolKeys.Is_LSHIFT_Down = false; break;
                                case "VK_RSHIFT": boolKeys.Is_RSHIFT_Down = false; break;
                                case "VK_LCONTROL": boolKeys.Is_LCONTROL_Down = false; break;
                                case "VK_RCONTROL": boolKeys.Is_RCONTROL_Down = false; break;
                                case "VK_LMENU": boolKeys.Is_LMENU_Down = false; break;
                                case "VK_RMENU": boolKeys.Is_RMENU_Down = false; break;
                                case "VK_BROWSER_BACK": boolKeys.Is_BROWSER_BACK_Down = false; break;
                                case "VK_BROWSER_FORWARD": boolKeys.Is_BROWSER_FORWARD_Down = false; break;
                                case "VK_BROWSER_REFRESH": boolKeys.Is_BROWSER_REFRESH_Down = false; break;
                                case "VK_BROWSER_STOP": boolKeys.Is_BROWSER_STOP_Down = false; break;
                                case "VK_BROWSER_SEARCH": boolKeys.Is_BROWSER_SEARCH_Down = false; break;
                                case "VK_BROWSER_FAVORITES": boolKeys.Is_BROWSER_FAVORITES_Down = false; break;
                                case "VK_BROWSER_HOME": boolKeys.Is_BROWSER_HOME_Down = false; break;
                                case "VK_VOLUME_MUTE": boolKeys.Is_VOLUME_MUTE_Down = false; break;
                                case "VK_VOLUME_DOWN": boolKeys.Is_VOLUME_DOWN_Down = false; break;
                                case "VK_VOLUME_UP": boolKeys.Is_VOLUME_UP_Down = false; break;
                                case "VK_MEDIA_NEXT_TRACK": boolKeys.Is_MEDIA_NEXT_TRACK_Down = false; break;
                                case "VK_MEDIA_PREV_TRACK": boolKeys.Is_MEDIA_PREV_TRACK_Down = false; break;
                                case "VK_MEDIA_STOP": boolKeys.Is_MEDIA_STOP_Down = false; break;
                                case "VK_MEDIA_PLAY_PAUSE": boolKeys.Is_MEDIA_PLAY_PAUSE_Down = false; break;
                                case "VK_LAUNCH_MAIL": boolKeys.Is_LAUNCH_MAIL_Down = false; break;
                                case "VK_LAUNCH_MEDIA_SELECT": boolKeys.Is_LAUNCH_MEDIA_SELECT_Down = false; break;
                                case "VK_LAUNCH_APP1": boolKeys.Is_LAUNCH_APP1_Down = false; break;
                                case "VK_LAUNCH_APP2": boolKeys.Is_LAUNCH_APP2_Down = false; break;
                                case "VK_OEM_1": boolKeys.Is_OEM_1_Down = false; break;
                                case "VK_OEM_PLUS": boolKeys.Is_OEM_PLUS_Down = false; break;
                                case "VK_OEM_COMMA": boolKeys.Is_OEM_COMMA_Down = false; break;
                                case "VK_OEM_MINUS": boolKeys.Is_OEM_MINUS_Down = false; break;
                                case "VK_OEM_PERIOD": boolKeys.Is_OEM_PERIOD_Down = false; break;
                                case "VK_OEM_2": boolKeys.Is_OEM_2_Down = false; break;
                                case "VK_OEM_3": boolKeys.Is_OEM_3_Down = false; break;
                                case "VK_OEM_4": boolKeys.Is_OEM_4_Down = false; break;
                                case "VK_OEM_5": boolKeys.Is_OEM_5_Down = false; break;
                                case "VK_OEM_6": boolKeys.Is_OEM_6_Down = false; break;
                                case "VK_OEM_7": boolKeys.Is_OEM_7_Down = false; break;
                                case "VK_OEM_8": boolKeys.Is_OEM_8_Down = false; break;
                                case "VK_OEM_102": boolKeys.Is_OEM_102_Down = false; break;
                                case "VK_PROCESSKEY": boolKeys.Is_PROCESSKEY_Down = false; break;
                                case "VK_PACKET": boolKeys.Is_PACKET_Down = false; break;
                                case "VK_ATTN": boolKeys.Is_ATTN_Down = false; break;
                                case "VK_CRSEL": boolKeys.Is_CRSEL_Down = false; break;
                                case "VK_EXSEL": boolKeys.Is_EXSEL_Down = false; break;
                                case "VK_EREOF": boolKeys.Is_EREOF_Down = false; break;
                                case "VK_PLAY": boolKeys.Is_PLAY_Down = false; break;
                                case "VK_ZOOM": boolKeys.Is_ZOOM_Down = false; break;
                                case "VK_NONAME": boolKeys.Is_NONAME_Down = false; break;
                                case "VK_PA1": boolKeys.Is_PA1_Down = false; break;
                                case "VK_OEM_CLEAR": boolKeys.Is_OEM_CLEAR_Down = false; break;
                            }
                        }
                    }






                    Command_Text temp = new Command_Text(command, command_Name, argument_List);
                    
                    command_s.Add(temp);
                    

                    sz++;
                }
            }

            
        }


    }
}
