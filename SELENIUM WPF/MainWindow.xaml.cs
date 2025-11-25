using ICSharpCode.AvalonEdit.Document;
using System.ComponentModel.Design;
using System.Diagnostics;
using System.Globalization;
using System.IO;
using System.Windows;
using System.Windows.Automation;
using System.Windows.Controls;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Threading;
using System.Xml.Linq;


namespace SELENIUM_WPF
{





    public partial class MainWindow : Window
    {

        //                                                                 ФИКС ПОДСКАЗОК
        /// ////////////////////////////////////////////////////////////////////////////


        //                                                           наведение мыши на поджсказку
        public void clue_Patch_L_Enter(object sender, EventArgs e)
        {
            //browse_TB.BorderBrush = new SolidColorBrush(Color.FromRgb(126, 180, 234));
            var brush = new SolidColorBrush(System.Windows.Media.Color.FromRgb(126, 180, 234));
            browse_TB.BorderBrush = brush;
        }



        //                                                           убирание мыши на поджсказку
        public void clue_Patch_L_Leave(object sender, EventArgs e)
        {
            browse_TB.ClearValue(System.Windows.Controls.Control.BorderBrushProperty);
        }


        //                                                           нажатие мыши на поджсказку
        private void Clue_Patch_L_MouseLeftButtonDown(object sender, MouseButtonEventArgs e)
        {
            browse_TB.Focus();
        }


        /// ////////////////////////////////////////////////////////////////////////////






        //                                                                 UPDATE
        /// ////////////////////////////////////////////////////////////////////////////


        public void Timer_Main_Tick(object? sender, EventArgs e)
        {
            if(!Is_Run_Update)
            {
                return;
            }

            //  проверка поля ввода для удаления подсказки
            //*****************************************
            if (browse_TB.Text.Length != 0)
            {
                clue_Patch_L.Visibility = Visibility.Hidden;
            }
            else
            {
                clue_Patch_L.Visibility = Visibility.Visible;
            }
            //*****************************************




            //  отслеживание состояния процесса
            //*****************************************
            if (Main_Proc == null)
            {
                if (!Is_Closed)
                {
                    status_L.Foreground = System.Windows.Media.Brushes.Red;
                    status_L.Content = "NONE";
                    start_B.IsEnabled = true;

                }
            }
            else
            {
                if (!Main_Proc.HasExited)
                {
                    status_L.Foreground = System.Windows.Media.Brushes.Fuchsia;
                    status_L.Content = "WORKING";
                    Is_Working = true;
                    Is_Closed = false;
                    start_B.IsEnabled = false;
                }
                else
                {
                    int code = Main_Proc.ExitCode;
                    if (code == 0)
                    {
                        status_L.Content = "CLOSED";
                        status_L.Foreground = System.Windows.Media.Brushes.Green;
                        Is_Working = false;
                        Is_Closed = true;
                        Main_Proc = null;
                        start_B.IsEnabled = true;
                    }
                    else
                    {
                        status_L.Foreground = System.Windows.Media.Brushes.Red;
                        status_L.Content = "NONE";
                        Is_Closed = false;
                        Is_Working = false;
                        Main_Proc = null;
                        if(code == -1)
                        {
                            crash_L.Visibility = Visibility.Visible;
                            crash_Code_L.Content = Convert.ToString(code) + "   PROCESS  KILL";
                            crash_Code_L.Visibility = Visibility.Visible;
                            start_B.IsEnabled = true;
                        }
                        else if(code == -1073741510)
                        {
                            crash_L.Visibility = Visibility.Visible;
                            crash_Code_L.Content = Convert.ToString(code) + "   CTRL+C  EVENT";
                            crash_Code_L.Visibility = Visibility.Visible;
                            start_B.IsEnabled = true;
                        }
                        else
                        {
                            crash_L.Visibility = Visibility.Visible;
                            crash_Code_L.Content = Convert.ToString(code);
                            crash_Code_L.Visibility = Visibility.Visible;
                            start_B.IsEnabled = true;
                        }
                            

                    }


                }
            }
            //*****************************************
        }





        /// ////////////////////////////////////////////////////////////////////////////










        //                                                                 INIT
        /// ////////////////////////////////////////////////////////////////////////////


        private static readonly HashSet<string> Extensions_ = new HashSet<string>(
       StringComparer.OrdinalIgnoreCase) { ".exe", ".msi", ".bat", ".cmd", ".com", ".ps1", ".vbs", ".js", ".jar", ".lnk" };   //   разрешённые расширения для файлов

        public Process Main_Proc = null;  //  Текущий процесс (запущенное приложение)
        AutomationElement Main_Root = null;  //  Корневой элемент текущего проццессса


        public List<ElementRecord> DATE = new List<ElementRecord>();  // cписок доступных в данный момент элементов


        public int Time_Out_APP = 5000;  //  время на закрытие приложения (мc) --------------------
        public int Time_Load_APP = 3000;  //  время на открытие приложения (мc) +++++++++++++++++++

        public string Patch_APP = "";  //  путь к исполнительному файлу

        public bool Is_Closed = false;  //  закрыт ли процесс корректно
        public bool Is_Working = false;  //  работает ли процесс
        public bool Is_Run_Update = false;      //  работает ли апдейт
        public bool Is_Console_App = false;      //  наличие GUI
       


        //public string This_Window_Name = "second_Form";  //  окно для обработки и поиска
        //public string This_Control_Name = "_TB";  // контрол для поиска и работы

       

        public ElementRecord Current_Control = new ElementRecord();

        public Dictionary<string, Func<List<string>, string, bool>> Commands = new Dictionary<string, Func<List<string>, string, bool>>();













        public MainWindow()
        {
            InitializeComponent();

            



        //  ссылка на окно для доступа 
        //******************************************
        GUI_Access.Main_W = this;
        //******************************************


            //  основной таймер запуск
            //******************************************
            Is_Run_Update = true;

            var timer_Main = new DispatcherTimer
            {
                Interval = TimeSpan.FromMicroseconds(1)
            };
            timer_Main.Tick += Timer_Main_Tick;
            timer_Main.Start();
            //******************************************


            //  привязка функций к подсказке пути
            //******************************************
            clue_Patch_L.MouseEnter += clue_Patch_L_Enter;
            clue_Patch_L.MouseLeave += clue_Patch_L_Leave;
            clue_Patch_L.MouseLeftButtonDown += Clue_Patch_L_MouseLeftButtonDown;
            //******************************************


            //  привязка функций к текст. полю пути
            //******************************************
            browse_TB.TextChanged += Browse_TB_TextChanged;
            //******************************************


            // привязка функций к кнопке пути
            //******************************************
            browse_B.Click += browse_B_click;
            //******************************************


            // привязка функций к кнопке старта
            //******************************************
            start_B.Click += start_B_Click;
            //******************************************


            // сброс лейблов к дефолту
            //******************************************
            status_L.Content = "NONE";
            status_L.Foreground = System.Windows.Media.Brushes.Red;

            crash_L.Visibility = Visibility.Hidden;
            crash_Code_L.Visibility = Visibility.Hidden;

            error_APP_L.Visibility = Visibility.Hidden;

            clue_Patch_L.Foreground = System.Windows.Media.Brushes.LightGray;
            //******************************************



            //  настройка вкладок и текстбокса для кода
            //*******************************************************888
            // --- Стилизация TextEditor для кода ---
            code_TB.BorderThickness = new Thickness(0);
            code_TB.Background = System.Windows.Media.Brushes.White;
            code_TB.Padding = new Thickness(5, 5, 0, 5);

            code_TB.FontFamily = new System.Windows.Media.FontFamily("JetBrains Mono");
            code_TB.FontWeight = FontWeights.Regular;
            code_TB.FontStyle = FontStyles.Normal;
            code_TB.FontSize = 14;

            // Прокрутка и отсутствие переноса
            code_TB.HorizontalScrollBarVisibility = ScrollBarVisibility.Auto;
            code_TB.VerticalScrollBarVisibility = ScrollBarVisibility.Auto;
            code_TB.WordWrap = false;

            // Нумерация строк
            code_TB.ShowLineNumbers = true;

            code_TB.LineNumbersForeground = System.Windows.Media.Brushes.Gray;




            // --- Стилизация вкладки "Главная" ---
            tab_Main.Padding = new Thickness(16, 8, 16, 8);
            tab_Main.Background = System.Windows.Media.Brushes.Gainsboro;
            tab_Main.BorderBrush = System.Windows.Media.Brushes.DarkGray;
            tab_Main.BorderThickness = new Thickness(1);
            tab_Main.Foreground = System.Windows.Media.Brushes.Black;
            tab_Main.FontWeight = FontWeights.SemiBold;
            tab_Main.Header = "Главная";

            // --- Стилизация вкладки "Код" ---
            tab_Code.Padding = new Thickness(16, 8, 16, 8);
            tab_Code.Background = System.Windows.Media.Brushes.Gainsboro;
            tab_Code.BorderBrush = System.Windows.Media.Brushes.DarkGray;
            tab_Code.BorderThickness = new Thickness(1);
            tab_Code.Foreground = System.Windows.Media.Brushes.Black;
            tab_Code.FontWeight = FontWeights.SemiBold;
            tab_Code.Header = "Код";

            // --- Подсветка выбранной вкладки ---
            this.Loaded += (s, e) =>
            {
                var updateTabs = new Action(() =>
                {
                    if (tab_Main.IsSelected)
                    {
                        tab_Main.Background = System.Windows.Media.Brushes.White;
                        tab_Main.BorderBrush = System.Windows.Media.Brushes.DodgerBlue;
                    }
                    else
                    {
                        tab_Main.Background = System.Windows.Media.Brushes.Gainsboro;
                        tab_Main.BorderBrush = System.Windows.Media.Brushes.DarkGray;
                    }

                    if (tab_Code.IsSelected)
                    {
                        tab_Code.Background = System.Windows.Media.Brushes.White;
                        tab_Code.BorderBrush = System.Windows.Media.Brushes.DodgerBlue;
                    }
                    else
                    {
                        tab_Code.Background = System.Windows.Media.Brushes.Gainsboro;
                        tab_Code.BorderBrush = System.Windows.Media.Brushes.DarkGray;
                    }
                });

                tab_Main.GotFocus += (s2, e2) => updateTabs();
                tab_Code.GotFocus += (s2, e2) => updateTabs();
                updateTabs();
            };
            //*******************************************************888




        }

        //AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA
        private void Stop_APP_Click(object sender, RoutedEventArgs e)
        {                                                                
            STOP_APP();
        }
        //AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA


        /// ////////////////////////////////////////////////////////////////////////////








        //                                                         выбрать путь к исполнительному файлу
        private void browse_B_click(object sender, EventArgs e)
        {

            // настройка диалога
            //******************************************
            var dlg = new Microsoft.Win32.OpenFileDialog
            {
                Filter = "Исполняемые файлы (*.exe;*.bat;*.cmd;*.msi;*.com;*.ps1;*.vbs;*.js;*.jar;*.lnk)|*.exe;*.bat;*.cmd;*.msi;*.com;*.ps1;*.vbs;*.js;*.jar;*.lnk",
                Title = "Выберите исполнительный файл программы"
            };
            //******************************************

            bool? result = dlg.ShowDialog();
            if (result == true)
            {
                Patch_APP = dlg.FileName;
                browse_TB.Text = Patch_APP;
            }
        }


        private void Browse_TB_TextChanged(object sender, TextChangedEventArgs e)
        {
           
            Patch_APP = browse_TB.Text;
            Patch_APP = Patch_APP.Trim();

           
            var formatted = new FormattedText(
                browse_TB.Text,
                CultureInfo.CurrentCulture,
                System.Windows.FlowDirection.LeftToRight, 
                new Typeface(browse_TB.FontFamily, browse_TB.FontStyle, browse_TB.FontWeight, browse_TB.FontStretch),
                browse_TB.FontSize,
                System.Windows.Media.Brushes.Black,       
                new NumberSubstitution(),
                1.0);

            double overflow = Math.Max(0, formatted.Width - browse_TB.ActualWidth);

            
            browseScroll.Maximum = overflow;
            browseScroll.ViewportSize = browse_TB.ActualWidth;

            browseScroll.Value = browse_TB.HorizontalOffset;
        }

       
        private void browseScroll_Scroll(object sender, System.Windows.Controls.Primitives.ScrollEventArgs e)
        {
            browse_TB.ScrollToHorizontalOffset(e.NewValue);
        }








        //                                                                          ЗАПУСКАТОР
        private void start_B_Click(object sender, EventArgs e)
        {
            string ext = System.IO.Path.GetExtension(Patch_APP);

            Time_Load_APP = int.Parse(time_Open_TB.Text) * 1000;
            Time_Out_APP = int.Parse(time_Stop_TB.Text) * 1000;

            if (browse_TB.Text == "")
            {
                error_APP_L.Content = "Выберите или введите путь к исполнительному файлу!";
                error_APP_L.Visibility = Visibility.Visible;
                return;
            }
            else if (!File.Exists(Patch_APP))
            {
                error_APP_L.Content = $"Путь к файлу {Patch_APP} не найден!";
                error_APP_L.Visibility = Visibility.Visible;
                status_L.Content = "NONE";
                status_L.Foreground = System.Windows.Media.Brushes.Red;

                return;
            }
            else if (!Extensions_.Contains(ext))
            {
                error_APP_L.Content = $"неверное расширение {ext}";
                error_APP_L.Visibility = Visibility.Visible;
                status_L.Content = "NONE";
                status_L.Foreground = System.Windows.Media.Brushes.Red;

                return;
            }
            else
            {
                error_APP_L.Visibility = Visibility.Hidden;

                var psi = new ProcessStartInfo
                {
                    FileName = Patch_APP,

                    UseShellExecute = true,
                    //Verb = "runas",                         // запрос прав админа
                    Arguments = "",                         // при необходимости можно передать аргументы
                    WindowStyle = ProcessWindowStyle.Normal
                };
                try
                {

                    Main_Proc = Process.Start(psi);

                    
                    
                    


                    crash_L.Visibility = Visibility.Hidden;
                    crash_Code_L.Visibility = Visibility.Hidden;
                    start_B.IsEnabled = false;

                    
                    if (Main_Proc == null)
                    {
                        Is_Run_Update = true;
                        error_APP_L.Content = $"Ошибка! Процесс {Patch_APP} не запущен!";
                        error_APP_L.Visibility = Visibility.Visible;
                        status_L.Content = "NONE";
                        status_L.Foreground = System.Windows.Media.Brushes.Red;
                        //   КАКИМ ТО БОКОМ ЗАКРЫВАТЬ ПРОВОДНИК ЕСЛИ  ЯРЛЫК УКАЗЫВАЕТ НА ПАПКУ
                        return;
                    }

                    Try_Get_GUI();
                }
                catch
                {
                    Is_Run_Update = true;
                    error_APP_L.Content = $"Ошибка при запуске файла {Patch_APP}";
                    error_APP_L.Visibility = Visibility.Visible;
                    status_L.Content = "NONE";
                    status_L.Foreground = System.Windows.Media.Brushes.Red;
                    try
                    {
                        Main_Proc.Kill();
                    }
                    catch
                    {
                        Is_Run_Update = true;
                        System.Windows.MessageBox.Show(
                            owner: this,
                            messageBoxText: "FATAL ERROR",
                            caption: "Selenium_WPF",
                            button: MessageBoxButton.OK,
                            icon: MessageBoxImage.Error
                        );
                        return;
                    }
                }

            }

        }


        //                                                                  попытка первичного чтоения GUI
        public void Try_Get_GUI()
        {
            //  начальное получение доступа к GUI
            //$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$
            if (!GUI_Access.Is_Console_APP(Patch_APP))   //  проверка на консоль
            {
                Is_Console_App = false;
                Is_Run_Update = false;
                status_L.Content = $"TRYING TO GET GUI ACCESS IN {Convert.ToString(Time_Load_APP / 1000)} SECONDS";
                status_L.Foreground = System.Windows.Media.Brushes.Fuchsia;
                
                
                //  дать приложению перерисовать картинку
                //*******************************************************************
                System.Windows.Application.Current.Dispatcher.Invoke(
                    DispatcherPriority.Render,
                    new Action(() => { })
                );
                //*******************************************************************

                Main_Root = GUI_Access.Get_Root(Main_Proc, Time_Load_APP);

                Is_Run_Update = true;



                if (Main_Root == null)
                {
                    error_APP_L.Content = $"Ошибка при зполучении корневого элемента  {Main_Proc.ProcessName}  \r\n Автоматическое закрытие процесса";
                    error_APP_L.Visibility = Visibility.Visible;
                    Main_Proc.Kill();
                    return;
                }
                try
                {
                    DATE = new List<ElementRecord>();
                    DATE = UI_Scanner.SnapshotControls(Main_Root);
                    Console.WriteLine();
                }
                catch
                {
                    error_APP_L.Content = $"Ошибка при обработке корневого элемента  {Main_Proc.ProcessName}  \r\n Автоматическое закрытие процесса";
                    error_APP_L.Visibility = Visibility.Visible;
                    Main_Proc.Kill();
                    return;
                }
            }
            else
            {
                Is_Console_App = true;

                System.Windows.MessageBox.Show(
                            owner: this,
                            messageBoxText: "Тестируемое приложение является консольным! Пожалуйста, добавьте задержку перед выполнением кода, а после его запуска вручную перейдите на окно тестируемого приложения",
                            caption: "Обработчик GUI",
                            button: MessageBoxButton.OK,
                            icon: MessageBoxImage.Information
                );
            }
            //$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$


            
        }


        //                                                             завершить внешний процесс
        public async void STOP_APP()
        {

            ProcessCloser.CloseProcessGracefully(Main_Proc, Time_Out_APP);
            start_B.IsEnabled = true;

        }




        //                                                                       тесты и запуск кода
        private void Run_Code_Click(object sender, RoutedEventArgs e)
        {

            //***************************************************************словарь
            Dictionary<string, Func<List<string>, string, bool>> Commands = new Dictionary<string, Func<List<string>, string, bool>>
            {
                //+++++++++++++++++++++++++++++++++++++++
                {
                    "Stay", (args, command_T) =>
                    {
                        int normal_arg = 0;
                        int lineNumber = 0;

                        int offset = code_TB.Document.Text.IndexOf(command_T);
                        if (offset >= 0)
                        {
                            DocumentLine line = code_TB.Document.GetLineByOffset(offset);
                            lineNumber = line.LineNumber;
                        }

                        if (args.Count != 1)
                        {
                            System.Windows.MessageBox.Show(
                                 owner: this,
                                messageBoxText: $"\tОшибка синтаксиса:\n неверное количество аргументов в строке с номером {lineNumber}\nОжидалось 1, а встречено {args.Count}",
                                caption: "Парсер псевдокода",
                                button: MessageBoxButton.OK,
                                icon: MessageBoxImage.Error
                            );
                            return false;
                        }
                        if(!int.TryParse(args[0], out normal_arg))
                        {
                            System.Windows.MessageBox.Show(
                                 owner: this,
                                messageBoxText: $"\tОшибка синтаксиса:\n неверный формат аргумента в строке с номером {lineNumber}\nОжидалось int",
                                caption: "Парсер псевдокода",
                                button: MessageBoxButton.OK,
                                icon: MessageBoxImage.Error
                            );
                            return false;
                        }

                        Thread.Sleep (normal_arg * 1000);
                        return true;
                    }
                },
                //+++++++++++++++++++++++++++++++++++++++

                //+++++++++++++++++++++++++++++++++++++++
                {
                    "Find_Window", (args, command_T) =>
                    {

                        int lineNumber = 0;

                        int offset = code_TB.Document.Text.IndexOf(command_T);
                        if (offset >= 0)
                        {
                            DocumentLine line = code_TB.Document.GetLineByOffset(offset);
                            lineNumber = line.LineNumber;
                        }

                        if (args.Count != 1)
                        {
                            System.Windows.MessageBox.Show(
                                 owner: this,
                                messageBoxText: $"\tОшибка синтаксиса:\n неверное количество аргументов в строке с номером {lineNumber}\nОжидалось 1, а встречено {args.Count}",
                                caption: "Парсер псевдокода",
                                button: MessageBoxButton.OK,
                                icon: MessageBoxImage.Error
                            );
                            return false;
                        }

                        try
                        {
                            DATE = new List<ElementRecord>();
                            Thread.Sleep (200);
                            DATE = WindowUITool.CaptureWindowUI(Main_Proc, args[0]);
                            Thread.Sleep (200);
                        }
                        catch
                        {
                            error_APP_L.Content = $"Ошибка при обработке корневого элемента окна  {args[0]}";
                            error_APP_L.Visibility = Visibility.Visible;
                            return false;
                        }


                        return true;
                    }
                },
                //+++++++++++++++++++++++++++++++++++++++

                //+++++++++++++++++++++++++++++++++++++++
                {
                    "Activate_MainWindow", (args, command_T) =>
                    {

                        int lineNumber = 0;

                        int offset = code_TB.Document.Text.IndexOf(command_T);
                        if (offset >= 0)
                        {
                            DocumentLine line = code_TB.Document.GetLineByOffset(offset);
                            lineNumber = line.LineNumber;
                        }

                        if (args.Count != 0)
                        {
                            System.Windows.MessageBox.Show(
                                 owner: this,
                                messageBoxText: $"\tОшибка синтаксиса:\n неверное количество аргументов в строке с номером {lineNumber}\nОжидалось 0, а встречено {args.Count}",
                                caption: "Парсер псевдокода",
                                button: MessageBoxButton.OK,
                                icon: MessageBoxImage.Error
                            );
                            return false;
                        }

                        try
                        {
                            Activate_Main_Process_W(Main_Proc);
                            Thread.Sleep(200);
                            DATE = new List<ElementRecord>();
                            DATE = UI_Scanner.SnapshotControls(Main_Root);
                        }
                        catch
                        {
                            error_APP_L.Content = $"Ошибка при обработке корневого элемента главного окна  {Main_Proc.ProcessName}  \r\n Автоматическое закрытие процесса";
                            error_APP_L.Visibility = Visibility.Visible;
                            Main_Proc.Kill();
                            return false;
                        }


                        return true;
                    }
                },
                //+++++++++++++++++++++++++++++++++++++++

                //+++++++++++++++++++++++++++++++++++++++
                {
                    "Stop_App", (args, command_T) =>
                    {

                        int lineNumber = 0;

                        int offset = code_TB.Document.Text.IndexOf(command_T);
                        if (offset >= 0)
                        {
                            DocumentLine line = code_TB.Document.GetLineByOffset(offset);
                            lineNumber = line.LineNumber;
                        }

                        if (args.Count != 0)
                        {
                            System.Windows.MessageBox.Show(
                                 owner: this,
                                messageBoxText: $"\tОшибка синтаксиса:\n неверное количество аргументов в строке с номером {lineNumber}\nОжидалось 0, а встречено {args.Count}",
                                caption: "Парсер псевдокода",
                                button: MessageBoxButton.OK,
                                icon: MessageBoxImage.Error
                            );
                            return false;
                        }


                        STOP_APP();

                        return true;
                    }
                },
                //+++++++++++++++++++++++++++++++++++++++

                //+++++++++++++++++++++++++++++++++++++++
                {
                    "Kill_App", (args, command_T) =>
                    {

                        int lineNumber = 0;

                        int offset = code_TB.Document.Text.IndexOf(command_T);
                        if (offset >= 0)
                        {
                            DocumentLine line = code_TB.Document.GetLineByOffset(offset);
                            lineNumber = line.LineNumber;
                        }

                        if (args.Count != 0)
                        {
                            System.Windows.MessageBox.Show(
                                 owner: this,
                                messageBoxText: $"\tОшибка синтаксиса:\n неверное количество аргументов в строке с номером {lineNumber}\nОжидалось 0, а встречено {args.Count}",
                                caption: "Парсер псевдокода",
                                button: MessageBoxButton.OK,
                                icon: MessageBoxImage.Error
                            );
                            return false;
                        }


                        Main_Proc.Kill();

                        return true;
                    }
                },
                //+++++++++++++++++++++++++++++++++++++++

                //+++++++++++++++++++++++++++++++++++++++
                {
                    "FindBy_Name", (args, command_T) =>
                    {

                        int lineNumber = 0;

                        int offset = code_TB.Document.Text.IndexOf(command_T);
                        if (offset >= 0)
                        {
                            DocumentLine line = code_TB.Document.GetLineByOffset(offset);
                            lineNumber = line.LineNumber;
                        }

                        if (args.Count != 1)
                        {
                            System.Windows.MessageBox.Show(
                                 owner: this,
                                messageBoxText: $"\tОшибка синтаксиса:\n неверное количество аргументов в строке с номером {lineNumber}\nОжидалось 1, а встречено {args.Count}",
                                caption: "Парсер псевдокода",
                                button: MessageBoxButton.OK,
                                icon: MessageBoxImage.Error
                            );
                            return false;
                        }

                        Current_Control = Find_By_Name(args[0]);
                        if (Current_Control == null)
                        {
                            System.Windows.MessageBox.Show(
                                owner: this,
                                messageBoxText: $"\tОшибка поиска:\n контрол {args[0]} не найден",
                                caption: "Парсер псевдокода",
                                button: MessageBoxButton.OK,
                                icon: MessageBoxImage.Error
                            );
                            return false;
                        }
                        return true;
                    }
                },
                //+++++++++++++++++++++++++++++++++++++++

                //+++++++++++++++++++++++++++++++++++++++
                {
                    "FindBy_Id", (args, command_T) =>
                    {

                        int lineNumber = 0;

                        int offset = code_TB.Document.Text.IndexOf(command_T);
                        if (offset >= 0)
                        {
                            DocumentLine line = code_TB.Document.GetLineByOffset(offset);
                            lineNumber = line.LineNumber;
                        }

                        if (args.Count != 1)
                        {
                            System.Windows.MessageBox.Show(
                                 owner: this,
                                messageBoxText: $"\tОшибка синтаксиса:\n неверное количество аргументов в строке с номером {lineNumber}\nОжидалось 1, а встречено {args.Count}",
                                caption: "Парсер псевдокода",
                                button: MessageBoxButton.OK,
                                icon: MessageBoxImage.Error
                            );
                            return false;
                        }

                        Current_Control = Find_By_Id(args[0]);
                        if (Current_Control == null)
                        {
                            System.Windows.MessageBox.Show(
                                owner: this,
                                messageBoxText: $"\tОшибка поиска:\n контрол {args[0]} не найден",
                                caption: "Парсер псевдокода",
                                button: MessageBoxButton.OK,
                                icon: MessageBoxImage.Error
                            );
                            return false;
                        }
                        return true;
                    }
                },
                //+++++++++++++++++++++++++++++++++++++++

                //+++++++++++++++++++++++++++++++++++++++
                {
                    "MouseMove", (args, command_T) =>
                    {
                        int x=0, y=0;
                        int lineNumber = 0;

                        int offset = code_TB.Document.Text.IndexOf(command_T);
                        if (offset >= 0)
                        {
                            DocumentLine line = code_TB.Document.GetLineByOffset(offset);
                            lineNumber = line.LineNumber;
                        }

                        if (args.Count != 1 && args.Count != 2)
                        {
                            System.Windows.MessageBox.Show(
                                 owner: this,
                                messageBoxText: $"\tОшибка синтаксиса:\n неверное количество аргументов в строке с номером {lineNumber}\nОжидалось 1 или 2, а встречено {args.Count}",
                                caption: "Парсер псевдокода",
                                button: MessageBoxButton.OK,
                                icon: MessageBoxImage.Error
                            );
                            return false;
                        }
                        if(args.Count == 2)
                        {
                            if(!int.TryParse(args[0], out x))
                            {
                                System.Windows.MessageBox.Show(
                                     owner: this,
                                    messageBoxText: $"\tОшибка синтаксиса:\n неверный формат аргумента в строке с номером {lineNumber}\nОжидалось int",
                                    caption: "Парсер псевдокода",
                                    button: MessageBoxButton.OK,
                                    icon: MessageBoxImage.Error
                                );
                                return false;
                            }
                            if(!int.TryParse(args[1], out y))
                            {
                                System.Windows.MessageBox.Show(
                                     owner: this,
                                    messageBoxText: $"\tОшибка синтаксиса:\n неверный формат аргумента в строке с номером {lineNumber}\nОжидалось int",
                                    caption: "Парсер псевдокода",
                                    button: MessageBoxButton.OK,
                                    icon: MessageBoxImage.Error
                                );
                                return false;
                            }

                            Mouse_Emulator.MoveCursor(x, y);
                        }
                        if (args.Count == 1)
                        {
                            if(args[0] != "Current_Control")
                            {
                                System.Windows.MessageBox.Show(
                                     owner: this,
                                    messageBoxText: $"\tОшибка синтаксиса:\n неверный формат аргумента в строке с номером {lineNumber}\nОжидалось Current_Control",
                                    caption: "Парсер псевдокода",
                                    button: MessageBoxButton.OK,
                                    icon: MessageBoxImage.Error
                                );
                                return false;
                            }
                            Mouse_Emulator.MoveCursor(Current_Control.CenterPoint.X, Current_Control.CenterPoint.Y);


                        }



                        return true;
                    }
                },
                //+++++++++++++++++++++++++++++++++++++++

                //+++++++++++++++++++++++++++++++++++++++
                {
                    "MouseClick", (args, command_T) =>
                    {
                        int x=0, y=0;
                        int lineNumber = 0;

                        int offset = code_TB.Document.Text.IndexOf(command_T);
                        if (offset >= 0)
                        {
                            DocumentLine line = code_TB.Document.GetLineByOffset(offset);
                            lineNumber = line.LineNumber;
                        }

                        if (args.Count != 1 && args.Count != 2 && args.Count !=0)
                        {
                            System.Windows.MessageBox.Show(
                                 owner: this,
                                messageBoxText: $"\tОшибка синтаксиса:\n неверное количество аргументов в строке с номером {lineNumber}\nОжидалось 1, 2 или 0, а встречено {args.Count}",
                                caption: "Парсер псевдокода",
                                button: MessageBoxButton.OK,
                                icon: MessageBoxImage.Error
                            );
                            return false;
                        }
                        if(args.Count == 2)
                        {
                            if(!int.TryParse(args[0], out x))
                            {
                                System.Windows.MessageBox.Show(
                                     owner: this,
                                    messageBoxText: $"\tОшибка синтаксиса:\n неверный формат аргумента в строке с номером {lineNumber}\nОжидалось int",
                                    caption: "Парсер псевдокода",
                                    button: MessageBoxButton.OK,
                                    icon: MessageBoxImage.Error
                                );
                                return false;
                            }
                            if(!int.TryParse(args[1], out y))
                            {
                                System.Windows.MessageBox.Show(
                                     owner: this,
                                    messageBoxText: $"\tОшибка синтаксиса:\n неверный формат аргумента в строке с номером {lineNumber}\nОжидалось int",
                                    caption: "Парсер псевдокода",
                                    button: MessageBoxButton.OK,
                                    icon: MessageBoxImage.Error
                                );
                                return false;
                            }

                            Mouse_Emulator.MoveAndClick(x, y);
                        }
                        if (args.Count == 1)
                        {
                            if(args[0] != "Current_Control")
                            {
                                System.Windows.MessageBox.Show(
                                     owner: this,
                                    messageBoxText: $"\tОшибка синтаксиса:\n неверный формат аргумента в строке с номером {lineNumber}\nОжидалось Current_Control",
                                    caption: "Парсер псевдокода",
                                    button: MessageBoxButton.OK,
                                    icon: MessageBoxImage.Error
                                );
                                return false;
                            }
                            Mouse_Emulator.MoveAndClick(Current_Control.CenterPoint.X, Current_Control.CenterPoint.Y);


                        }
                        if(args.Count == 0)
                        {
                            Mouse_Emulator.MouseClick();
                        }




                        return true;
                    }
                },
                //+++++++++++++++++++++++++++++++++++++++

                //+++++++++++++++++++++++++++++++++++++++
                {
                    "MouseRightClick", (args, command_T) =>
                    {
                        int x=0, y=0;
                        int lineNumber = 0;

                        int offset = code_TB.Document.Text.IndexOf(command_T);
                        if (offset >= 0)
                        {
                            DocumentLine line = code_TB.Document.GetLineByOffset(offset);
                            lineNumber = line.LineNumber;
                        }

                        if (args.Count != 1 && args.Count != 2 && args.Count !=0)
                        {
                            System.Windows.MessageBox.Show(
                                 owner: this,
                                messageBoxText: $"\tОшибка синтаксиса:\n неверное количество аргументов в строке с номером {lineNumber}\nОжидалось 1, 2 или 0, а встречено {args.Count}",
                                caption: "Парсер псевдокода",
                                button: MessageBoxButton.OK,
                                icon: MessageBoxImage.Error
                            );
                            return false;
                        }
                        if(args.Count == 2)
                        {
                            if(!int.TryParse(args[0], out x))
                            {
                                System.Windows.MessageBox.Show(
                                     owner: this,
                                    messageBoxText: $"\tОшибка синтаксиса:\n неверный формат аргумента в строке с номером {lineNumber}\nОжидалось int",
                                    caption: "Парсер псевдокода",
                                    button: MessageBoxButton.OK,
                                    icon: MessageBoxImage.Error
                                );
                                return false;
                            }
                            if(!int.TryParse(args[1], out y))
                            {
                                System.Windows.MessageBox.Show(
                                     owner: this,
                                    messageBoxText: $"\tОшибка синтаксиса:\n неверный формат аргумента в строке с номером {lineNumber}\nОжидалось int",
                                    caption: "Парсер псевдокода",
                                    button: MessageBoxButton.OK,
                                    icon: MessageBoxImage.Error
                                );
                                return false;
                            }

                            Mouse_Emulator.MoveAndClick(x, y, Mouse_Emulator.MouseButton.Right);
                        }
                        if (args.Count == 1)
                        {
                            if(args[0] != "Current_Control")
                            {
                                System.Windows.MessageBox.Show(
                                     owner: this,
                                    messageBoxText: $"\tОшибка синтаксиса:\n неверный формат аргумента в строке с номером {lineNumber}\nОжидалось Current_Control",
                                    caption: "Парсер псевдокода",
                                    button: MessageBoxButton.OK,
                                    icon: MessageBoxImage.Error
                                );
                                return false;
                            }
                            Mouse_Emulator.MoveAndClick(Current_Control.CenterPoint.X, Current_Control.CenterPoint.Y, Mouse_Emulator.MouseButton.Right);


                        }
                        if(args.Count == 0)
                        {
                            Mouse_Emulator.MouseClick(Mouse_Emulator.MouseButton.Right);
                        }




                        return true;
                    }
                },
                //+++++++++++++++++++++++++++++++++++++++

                //+++++++++++++++++++++++++++++++++++++++
                {
                    "MouseMiddleClick", (args, command_T) =>
                    {
                        int x=0, y=0;
                        int lineNumber = 0;

                        int offset = code_TB.Document.Text.IndexOf(command_T);
                        if (offset >= 0)
                        {
                            DocumentLine line = code_TB.Document.GetLineByOffset(offset);
                            lineNumber = line.LineNumber;
                        }

                        if (args.Count != 1 && args.Count != 2 && args.Count !=0)
                        {
                            System.Windows.MessageBox.Show(
                                 owner: this,
                                messageBoxText: $"\tОшибка синтаксиса:\n неверное количество аргументов в строке с номером {lineNumber}\nОжидалось 1, 2 или 0, а встречено {args.Count}",
                                caption: "Парсер псевдокода",
                                button: MessageBoxButton.OK,
                                icon: MessageBoxImage.Error
                            );
                            return false;
                        }
                        if(args.Count == 2)
                        {
                            if(!int.TryParse(args[0], out x))
                            {
                                System.Windows.MessageBox.Show(
                                     owner: this,
                                    messageBoxText: $"\tОшибка синтаксиса:\n неверный формат аргумента в строке с номером {lineNumber}\nОжидалось int",
                                    caption: "Парсер псевдокода",
                                    button: MessageBoxButton.OK,
                                    icon: MessageBoxImage.Error
                                );
                                return false;
                            }
                            if(!int.TryParse(args[1], out y))
                            {
                                System.Windows.MessageBox.Show(
                                     owner: this,
                                    messageBoxText: $"\tОшибка синтаксиса:\n неверный формат аргумента в строке с номером {lineNumber}\nОжидалось int",
                                    caption: "Парсер псевдокода",
                                    button: MessageBoxButton.OK,
                                    icon: MessageBoxImage.Error
                                );
                                return false;
                            }

                            Mouse_Emulator.MoveAndClick(x, y, Mouse_Emulator.MouseButton.Middle);
                        }
                        if (args.Count == 1)
                        {
                            if(args[0] != "Current_Control")
                            {
                                System.Windows.MessageBox.Show(
                                     owner: this,
                                    messageBoxText: $"\tОшибка синтаксиса:\n неверный формат аргумента в строке с номером {lineNumber}\nОжидалось Current_Control",
                                    caption: "Парсер псевдокода",
                                    button: MessageBoxButton.OK,
                                    icon: MessageBoxImage.Error
                                );
                                return false;
                            }
                            Mouse_Emulator.MoveAndClick(Current_Control.CenterPoint.X, Current_Control.CenterPoint.Y, Mouse_Emulator.MouseButton.Middle);


                        }
                        if(args.Count == 0)
                        {
                            Mouse_Emulator.MouseClick(Mouse_Emulator.MouseButton.Middle);
                        }




                        return true;
                    }
                },
                //+++++++++++++++++++++++++++++++++++++++

                //+++++++++++++++++++++++++++++++++++++++
                {
                    "MouseDown", (args, command_T) =>
                    {

                        int lineNumber = 0;

                        int offset = code_TB.Document.Text.IndexOf(command_T);
                        if (offset >= 0)
                        {
                            DocumentLine line = code_TB.Document.GetLineByOffset(offset);
                            lineNumber = line.LineNumber;
                        }

                        if (args.Count != 0)
                        {
                            System.Windows.MessageBox.Show(
                                 owner: this,
                                messageBoxText: $"\tОшибка синтаксиса:\n неверное количество аргументов в строке с номером {lineNumber}\nОжидалось 0, а встречено {args.Count}",
                                caption: "Парсер псевдокода",
                                button: MessageBoxButton.OK,
                                icon: MessageBoxImage.Error
                            );
                            return false;
                        }


                        Mouse_Emulator.MouseDown();

                        return true;
                    }
                },
                //+++++++++++++++++++++++++++++++++++++++

                //+++++++++++++++++++++++++++++++++++++++
                {
                    "MouseUp", (args, command_T) =>
                    {

                        int lineNumber = 0;

                        int offset = code_TB.Document.Text.IndexOf(command_T);
                        if (offset >= 0)
                        {
                            DocumentLine line = code_TB.Document.GetLineByOffset(offset);
                            lineNumber = line.LineNumber;
                        }

                        if (args.Count != 0)
                        {
                            System.Windows.MessageBox.Show(
                                 owner: this,
                                messageBoxText: $"\tОшибка синтаксиса:\n неверное количество аргументов в строке с номером {lineNumber}\nОжидалось 0, а встречено {args.Count}",
                                caption: "Парсер псевдокода",
                                button: MessageBoxButton.OK,
                                icon: MessageBoxImage.Error
                            );
                            return false;
                        }


                        Mouse_Emulator.MouseUp();

                        return true;
                    }
                },
                //+++++++++++++++++++++++++++++++++++++++

                //+++++++++++++++++++++++++++++++++++++++
                {
                    "MouseRightDown", (args, command_T) =>
                    {

                        int lineNumber = 0;

                        int offset = code_TB.Document.Text.IndexOf(command_T);
                        if (offset >= 0)
                        {
                            DocumentLine line = code_TB.Document.GetLineByOffset(offset);
                            lineNumber = line.LineNumber;
                        }

                        if (args.Count != 0)
                        {
                            System.Windows.MessageBox.Show(
                                 owner: this,
                                messageBoxText: $"\tОшибка синтаксиса:\n неверное количество аргументов в строке с номером {lineNumber}\nОжидалось 0, а встречено {args.Count}",
                                caption: "Парсер псевдокода",
                                button: MessageBoxButton.OK,
                                icon: MessageBoxImage.Error
                            );
                            return false;
                        }


                        Mouse_Emulator.MouseDown(Mouse_Emulator.MouseButton.Right);

                        return true;
                    }
                },
                //+++++++++++++++++++++++++++++++++++++++

                //+++++++++++++++++++++++++++++++++++++++
                {
                    "MouseRightUp", (args, command_T) =>
                    {

                        int lineNumber = 0;

                        int offset = code_TB.Document.Text.IndexOf(command_T);
                        if (offset >= 0)
                        {
                            DocumentLine line = code_TB.Document.GetLineByOffset(offset);
                            lineNumber = line.LineNumber;
                        }

                        if (args.Count != 0)
                        {
                            System.Windows.MessageBox.Show(
                                 owner: this,
                                messageBoxText: $"\tОшибка синтаксиса:\n неверное количество аргументов в строке с номером {lineNumber}\nОжидалось 0, а встречено {args.Count}",
                                caption: "Парсер псевдокода",
                                button: MessageBoxButton.OK,
                                icon: MessageBoxImage.Error
                            );
                            return false;
                        }


                        Mouse_Emulator.MouseUp(Mouse_Emulator.MouseButton.Right);

                        return true;
                    }
                },
                //+++++++++++++++++++++++++++++++++++++++

                //+++++++++++++++++++++++++++++++++++++++
                {
                    "MouseMiddleDown", (args, command_T) =>
                    {

                        int lineNumber = 0;

                        int offset = code_TB.Document.Text.IndexOf(command_T);
                        if (offset >= 0)
                        {
                            DocumentLine line = code_TB.Document.GetLineByOffset(offset);
                            lineNumber = line.LineNumber;
                        }

                        if (args.Count != 0)
                        {
                            System.Windows.MessageBox.Show(
                                 owner: this,
                                messageBoxText: $"\tОшибка синтаксиса:\n неверное количество аргументов в строке с номером {lineNumber}\nОжидалось 0, а встречено {args.Count}",
                                caption: "Парсер псевдокода",
                                button: MessageBoxButton.OK,
                                icon: MessageBoxImage.Error
                            );
                            return false;
                        }


                        Mouse_Emulator.MouseDown(Mouse_Emulator.MouseButton.Middle);

                        return true;
                    }
                },
                //+++++++++++++++++++++++++++++++++++++++

                //+++++++++++++++++++++++++++++++++++++++
                {
                    "MouseMiddleUp", (args, command_T) =>
                    {

                        int lineNumber = 0;

                        int offset = code_TB.Document.Text.IndexOf(command_T);
                        if (offset >= 0)
                        {
                            DocumentLine line = code_TB.Document.GetLineByOffset(offset);
                            lineNumber = line.LineNumber;
                        }

                        if (args.Count != 0)
                        {
                            System.Windows.MessageBox.Show(
                                 owner: this,
                                messageBoxText: $"\tОшибка синтаксиса:\n неверное количество аргументов в строке с номером {lineNumber}\nОжидалось 0, а встречено {args.Count}",
                                caption: "Парсер псевдокода",
                                button: MessageBoxButton.OK,
                                icon: MessageBoxImage.Error
                            );
                            return false;
                        }


                        Mouse_Emulator.MouseUp(Mouse_Emulator.MouseButton.Middle);

                        return true;
                    }
                },
                //+++++++++++++++++++++++++++++++++++++++

                //+++++++++++++++++++++++++++++++++++++++
                {
                    "KeyClick", (args, command_T) =>
                    {

                        int lineNumber = 0;

                        int offset = code_TB.Document.Text.IndexOf(command_T);
                        if (offset >= 0)
                        {
                            DocumentLine line = code_TB.Document.GetLineByOffset(offset);
                            lineNumber = line.LineNumber;
                        }

                        if (args.Count <1)
                        {
                            System.Windows.MessageBox.Show(
                                 owner: this,
                                messageBoxText: $"\tОшибка синтаксиса:\n неверное количество аргументов в строке с номером {lineNumber}\nОжидалось 1 и более, а встречено {args.Count}",
                                caption: "Парсер псевдокода",
                                button: MessageBoxButton.OK,
                                icon: MessageBoxImage.Error
                            );
                            return false;
                        }
                        else
                        {
                            foreach (string key in args)
                            {
                                if (Keyboard_Emulator.VirtualKeyMap.TryGetValue(key, out byte keyCode))
                                {
                                    Keyboard_Emulator.KeyPress(keyCode);
                                }
                                else
                                {
                                    System.Windows.MessageBox.Show(
                                         owner: this,
                                        messageBoxText: $"\tОшибка синтаксиса:\nнеизвестная клавиша в строке {lineNumber}",
                                        caption: "Парсер псевдокода",
                                        button: MessageBoxButton.OK,
                                        icon: MessageBoxImage.Error
                                    );
                                    return false;
                                }
                            }
                        }
                        return true;
                    }
                },
                //+++++++++++++++++++++++++++++++++++++++


                //+++++++++++++++++++++++++++++++++++++++
                {
                    "KeyDown", (args, command_T) =>
                    {

                        int lineNumber = 0;

                        int offset = code_TB.Document.Text.IndexOf(command_T);
                        if (offset >= 0)
                        {
                            DocumentLine line = code_TB.Document.GetLineByOffset(offset);
                            lineNumber = line.LineNumber;
                        }

                        if (args.Count <1)
                        {
                            System.Windows.MessageBox.Show(
                                 owner: this,
                                messageBoxText: $"\tОшибка синтаксиса:\n неверное количество аргументов в строке с номером {lineNumber}\nОжидалось 1 и более, а встречено {args.Count}",
                                caption: "Парсер псевдокода",
                                button: MessageBoxButton.OK,
                                icon: MessageBoxImage.Error
                            );
                            return false;
                        }
                        else
                        {
                            foreach (string key in args)
                            {
                                if (Keyboard_Emulator.VirtualKeyMap.TryGetValue(key, out byte keyCode))
                                {
                                    Keyboard_Emulator.KeyDown(keyCode);
                                }
                                else
                                {
                                    System.Windows.MessageBox.Show(
                                         owner: this,
                                        messageBoxText: $"\tОшибка синтаксиса:\nнеизвестная клавиша в строке {lineNumber}",
                                        caption: "Парсер псевдокода",
                                        button: MessageBoxButton.OK,
                                        icon: MessageBoxImage.Error
                                    );
                                    return false;
                                }
                            }
                        }
                        return true;
                    }
                },
                //+++++++++++++++++++++++++++++++++++++++

                //+++++++++++++++++++++++++++++++++++++++
                {
                    "KeyUp", (args, command_T) =>
                    {

                        int lineNumber = 0;

                        int offset = code_TB.Document.Text.IndexOf(command_T);
                        if (offset >= 0)
                        {
                            DocumentLine line = code_TB.Document.GetLineByOffset(offset);
                            lineNumber = line.LineNumber;
                        }

                        if (args.Count <1)
                        {
                            System.Windows.MessageBox.Show(
                                 owner: this,
                                messageBoxText: $"\tОшибка синтаксиса:\n неверное количество аргументов в строке с номером {lineNumber}\nОжидалось 1 и более, а встречено {args.Count}",
                                caption: "Парсер псевдокода",
                                button: MessageBoxButton.OK,
                                icon: MessageBoxImage.Error
                            );
                            return false;
                        }
                        else
                        {
                            foreach (string key in args)
                            {
                                if (Keyboard_Emulator.VirtualKeyMap.TryGetValue(key, out byte keyCode))
                                {
                                    Keyboard_Emulator.KeyDown(keyCode);
                                }
                                else
                                {
                                    System.Windows.MessageBox.Show(
                                         owner: this,
                                        messageBoxText: $"\tОшибка синтаксиса:\nнеизвестная клавиша в строке {lineNumber}",
                                        caption: "Парсер псевдокода",
                                        button: MessageBoxButton.OK,
                                        icon: MessageBoxImage.Error
                                    );
                                    return false;
                                }
                            }
                        }
                        return true;
                    }
                },
                //+++++++++++++++++++++++++++++++++++++++

                //+++++++++++++++++++++++++++++++++++++++
                {
                    "TextEnter", (args, command_T) =>
                    {

                        int lineNumber = 0;
                        int delay = 0;

                        int offset = code_TB.Document.Text.IndexOf(command_T);
                        if (offset >= 0)
                        {
                            DocumentLine line = code_TB.Document.GetLineByOffset(offset);
                            lineNumber = line.LineNumber;
                        }

                        if (args.Count != 1 && args.Count != 2)
                        {
                            System.Windows.MessageBox.Show(
                                 owner: this,
                                messageBoxText: $"\tОшибка синтаксиса:\n неверное количество аргументов в строке с номером {lineNumber}\nОжидалось 1 или 2, а встречено {args.Count}",
                                caption: "Парсер псевдокода",
                                button: MessageBoxButton.OK,
                                icon: MessageBoxImage.Error
                            );
                            return false;
                        }
                        if(args.Count == 1)
                        {
                            Keyboard_Emulator.TypeText(args[0]);
                        }
                        else if(args.Count == 2)
                        {
                            if(!int.TryParse(args[1], out delay))
                            {
                                System.Windows.MessageBox.Show(
                                     owner: this,
                                    messageBoxText: $"\tОшибка синтаксиса:\n неверный формат аргумента в строке с номером {lineNumber}\nОжидалось int",
                                    caption: "Парсер псевдокода",
                                    button: MessageBoxButton.OK,
                                    icon: MessageBoxImage.Error
                                );
                                return false;
                            }
                            Keyboard_Emulator.TypeText(args[0],delay);
                        }
                        return true;
                    }
                },
                //+++++++++++++++++++++++++++++++++++++++

                //+++++++++++++++++++++++++++++++++++++++
                {
                    "SwitchKeyboardLayout", (args, command_T) =>
                    {

                        int lineNumber = 0;

                        int offset = code_TB.Document.Text.IndexOf(command_T);
                        if (offset >= 0)
                        {
                            DocumentLine line = code_TB.Document.GetLineByOffset(offset);
                            lineNumber = line.LineNumber;
                        }

                        if (args.Count != 0)
                        {
                            System.Windows.MessageBox.Show(
                                 owner: this,
                                messageBoxText: $"\tОшибка синтаксиса:\n неверное количество аргументов в строке с номером {lineNumber}\nОжидалось 0, а встречено {args.Count}",
                                caption: "Парсер псевдокода",
                                button: MessageBoxButton.OK,
                                icon: MessageBoxImage.Error
                            );
                            return false;
                        }


                        Keyboard_Emulator.SwitchKeyboardLayout();

                        return true;
                    }
                },
                //+++++++++++++++++++++++++++++++++++++++


                

            };
            //***************************************************************словарь





            Analyze_Code();                                                    //  считывание кода



            
            List<string> pressedItems = new List<string>();

            // Кнопки мыши
            if (boolKeys.Is_LeftMouse_Down) pressedItems.Add("Левая кнопка мыши");
            if (boolKeys.Is_RightMouse_Down) pressedItems.Add("Правая кнопка мыши");
            if (boolKeys.Is_MiddleMouse_Down) pressedItems.Add("Средняя кнопка мыши");

            // Клавиши
            if (boolKeys.Is_CANCEL_Down) pressedItems.Add("CANCEL");
            if (boolKeys.Is_BACK_Down) pressedItems.Add("BACKSPACE");
            if (boolKeys.Is_TAB_Down) pressedItems.Add("TAB");
            if (boolKeys.Is_CLEAR_Down) pressedItems.Add("CLEAR");
            if (boolKeys.Is_RETURN_Down) pressedItems.Add("ENTER");
            if (boolKeys.Is_SHIFT_Down) pressedItems.Add("SHIFT");
            if (boolKeys.Is_CONTROL_Down) pressedItems.Add("CONTROL");
            if (boolKeys.Is_MENU_Down) pressedItems.Add("ALT");
            if (boolKeys.Is_PAUSE_Down) pressedItems.Add("PAUSE");
            if (boolKeys.Is_CAPITAL_Down) pressedItems.Add("CAPS LOCK");
            if (boolKeys.Is_ESCAPE_Down) pressedItems.Add("ESCAPE");
            if (boolKeys.Is_SPACE_Down) pressedItems.Add("ПРОБЕЛ");
            if (boolKeys.Is_PRIOR_Down) pressedItems.Add("PAGE UP");
            if (boolKeys.Is_NEXT_Down) pressedItems.Add("PAGE DOWN");
            if (boolKeys.Is_END_Down) pressedItems.Add("END");
            if (boolKeys.Is_HOME_Down) pressedItems.Add("HOME");
            if (boolKeys.Is_LEFT_Down) pressedItems.Add("СТРЕЛКА ВЛЕВО");
            if (boolKeys.Is_UP_Down) pressedItems.Add("СТРЕЛКА ВВЕРХ");
            if (boolKeys.Is_RIGHT_Down) pressedItems.Add("СТРЕЛКА ВПРАВО");
            if (boolKeys.Is_DOWN_Down) pressedItems.Add("СТРЕЛКА ВНИЗ");
            if (boolKeys.Is_SELECT_Down) pressedItems.Add("SELECT");
            if (boolKeys.Is_PRINT_Down) pressedItems.Add("PRINT");
            if (boolKeys.Is_EXECUTE_Down) pressedItems.Add("EXECUTE");
            if (boolKeys.Is_SNAPSHOT_Down) pressedItems.Add("PRINT SCREEN");
            if (boolKeys.Is_INSERT_Down) pressedItems.Add("INSERT");
            if (boolKeys.Is_DELETE_Down) pressedItems.Add("DELETE");
            if (boolKeys.Is_HELP_Down) pressedItems.Add("HELP");

            // Цифры
            for (int i = 0; i <= 9; i++)
            {
                if ((bool)typeof(boolKeys).GetField($"Is_{i}_Down")?.GetValue(null))
                    pressedItems.Add(i.ToString());
            }

            // Буквы A–Z
            for (char c = 'A'; c <= 'Z'; c++)
            {
                if ((bool)typeof(boolKeys).GetField($"Is_{c}_Down")?.GetValue(null))
                    pressedItems.Add(c.ToString());
            }

            // Специальные клавиши
            if (boolKeys.Is_LWIN_Down) pressedItems.Add("ЛЕВЫЙ WIN");
            if (boolKeys.Is_RWIN_Down) pressedItems.Add("ПРАВЫЙ WIN");
            if (boolKeys.Is_APPS_Down) pressedItems.Add("КЛАВИША МЕНЮ");

            // Numpad
            if (boolKeys.Is_NUMPAD0_Down) pressedItems.Add("NUMPAD 0");
            if (boolKeys.Is_NUMPAD1_Down) pressedItems.Add("NUMPAD 1");
            if (boolKeys.Is_NUMPAD2_Down) pressedItems.Add("NUMPAD 2");
            if (boolKeys.Is_NUMPAD3_Down) pressedItems.Add("NUMPAD 3");
            if (boolKeys.Is_NUMPAD4_Down) pressedItems.Add("NUMPAD 4");
            if (boolKeys.Is_NUMPAD5_Down) pressedItems.Add("NUMPAD 5");
            if (boolKeys.Is_NUMPAD6_Down) pressedItems.Add("NUMPAD 6");
            if (boolKeys.Is_NUMPAD7_Down) pressedItems.Add("NUMPAD 7");
            if (boolKeys.Is_NUMPAD8_Down) pressedItems.Add("NUMPAD 8");
            if (boolKeys.Is_NUMPAD9_Down) pressedItems.Add("NUMPAD 9");
            if (boolKeys.Is_MULTIPLY_Down) pressedItems.Add("NUMPAD *");
            if (boolKeys.Is_ADD_Down) pressedItems.Add("NUMPAD +");
            if (boolKeys.Is_SEPARATOR_Down) pressedItems.Add("NUMPAD РАЗДЕЛИТЕЛЬ");
            if (boolKeys.Is_SUBTRACT_Down) pressedItems.Add("NUMPAD -");
            if (boolKeys.Is_DECIMAL_Down) pressedItems.Add("NUMPAD .");
            if (boolKeys.Is_DIVIDE_Down) pressedItems.Add("NUMPAD /");

            // F1–F24
            for (int i = 1; i <= 24; i++)
            {
                if ((bool)typeof(boolKeys).GetField($"Is_F{i}_Down")?.GetValue(null))
                    pressedItems.Add($"F{i}");
            }

            if (boolKeys.Is_NUMLOCK_Down) pressedItems.Add("NUM LOCK");
            if (boolKeys.Is_SCROLL_Down) pressedItems.Add("SCROLL LOCK");

            if (boolKeys.Is_LSHIFT_Down) pressedItems.Add("ЛЕВЫЙ SHIFT");
            if (boolKeys.Is_RSHIFT_Down) pressedItems.Add("ПРАВЫЙ SHIFT");
            if (boolKeys.Is_LCONTROL_Down) pressedItems.Add("ЛЕВЫЙ CONTROL");
            if (boolKeys.Is_RCONTROL_Down) pressedItems.Add("ПРАВЫЙ CONTROL");
            if (boolKeys.Is_LMENU_Down) pressedItems.Add("ЛЕВЫЙ ALT");
            if (boolKeys.Is_RMENU_Down) pressedItems.Add("ПРАВЫЙ ALT");

            if (boolKeys.Is_BROWSER_BACK_Down) pressedItems.Add("НАЗАД (браузер)");
            if (boolKeys.Is_BROWSER_FORWARD_Down) pressedItems.Add("ВПЕРЁД (браузер)");
            if (boolKeys.Is_BROWSER_REFRESH_Down) pressedItems.Add("ОБНОВИТЬ");
            if (boolKeys.Is_BROWSER_STOP_Down) pressedItems.Add("СТОП (браузер)");
            if (boolKeys.Is_BROWSER_SEARCH_Down) pressedItems.Add("ПОИСК");
            if (boolKeys.Is_BROWSER_FAVORITES_Down) pressedItems.Add("ИЗБРАННОЕ");
            if (boolKeys.Is_BROWSER_HOME_Down) pressedItems.Add("ДОМОЙ");
            if (boolKeys.Is_VOLUME_MUTE_Down) pressedItems.Add("ВЫКЛ. ЗВУКА");
            if (boolKeys.Is_VOLUME_DOWN_Down) pressedItems.Add("ГРОМКОСТЬ -");
            if (boolKeys.Is_VOLUME_UP_Down) pressedItems.Add("ГРОМКОСТЬ +");
            if (boolKeys.Is_MEDIA_NEXT_TRACK_Down) pressedItems.Add("СЛЕД. ТРЕК");
            if (boolKeys.Is_MEDIA_PREV_TRACK_Down) pressedItems.Add("ПРЕД. ТРЕК");
            if (boolKeys.Is_MEDIA_STOP_Down) pressedItems.Add("СТОП МЕДИА");
            if (boolKeys.Is_MEDIA_PLAY_PAUSE_Down) pressedItems.Add("PLAY/PAUSE");
            if (boolKeys.Is_LAUNCH_MAIL_Down) pressedItems.Add("ПОЧТА");
            if (boolKeys.Is_LAUNCH_MEDIA_SELECT_Down) pressedItems.Add("МЕДИА-ВЫБОР");
            if (boolKeys.Is_LAUNCH_APP1_Down) pressedItems.Add("ПРИЛОЖЕНИЕ 1");
            if (boolKeys.Is_LAUNCH_APP2_Down) pressedItems.Add("ПРИЛОЖЕНИЕ 2");

            if (boolKeys.Is_OEM_1_Down) pressedItems.Add("OEM 1 (; : на US)");
            if (boolKeys.Is_OEM_PLUS_Down) pressedItems.Add("OEM + (= + на US)");
            if (boolKeys.Is_OEM_COMMA_Down) pressedItems.Add("OEM , (< на US)");
            if (boolKeys.Is_OEM_MINUS_Down) pressedItems.Add("OEM - (_ на US)");
            if (boolKeys.Is_OEM_PERIOD_Down) pressedItems.Add("OEM . (> на US)");
            if (boolKeys.Is_OEM_2_Down) pressedItems.Add("OEM 2 (/ ? на US)");
            if (boolKeys.Is_OEM_3_Down) pressedItems.Add("OEM 3 (` ~ на US)");
            if (boolKeys.Is_OEM_4_Down) pressedItems.Add("OEM 4 ([ { на US)");
            if (boolKeys.Is_OEM_5_Down) pressedItems.Add("OEM 5 (\\ | на US)");
            if (boolKeys.Is_OEM_6_Down) pressedItems.Add("OEM 6 (] } на US)");
            if (boolKeys.Is_OEM_7_Down) pressedItems.Add("OEM 7 (' \" на US)");
            if (boolKeys.Is_OEM_8_Down) pressedItems.Add("OEM 8");
            if (boolKeys.Is_OEM_102_Down) pressedItems.Add("OEM 102");

            if (boolKeys.Is_PROCESSKEY_Down) pressedItems.Add("PROCESS KEY");
            if (boolKeys.Is_PACKET_Down) pressedItems.Add("PACKET");
            if (boolKeys.Is_ATTN_Down) pressedItems.Add("ATTN");
            if (boolKeys.Is_CRSEL_Down) pressedItems.Add("CRSEL");
            if (boolKeys.Is_EXSEL_Down) pressedItems.Add("EXSEL");
            if (boolKeys.Is_EREOF_Down) pressedItems.Add("EREOF");
            if (boolKeys.Is_PLAY_Down) pressedItems.Add("PLAY");
            if (boolKeys.Is_ZOOM_Down) pressedItems.Add("ZOOM");
            if (boolKeys.Is_NONAME_Down) pressedItems.Add("NONAME");
            if (boolKeys.Is_PA1_Down) pressedItems.Add("PA1");
            if (boolKeys.Is_OEM_CLEAR_Down) pressedItems.Add("OEM CLEAR");

            
            if (pressedItems.Count > 0)
            {
                string itemsList = string.Join("\n• ", pressedItems);
                string fullMessage =
                    "В вашем псевдокоде остались нажатыми следующие клавиши и/или кнопки мыши:\n\n" +
                    "• " + itemsList + "\n\n" +
                    "Это может привести к некорректному неконтролируемому поведению системы.\n\n" +
                    "Продолжить выполнение?";

                MessageBoxResult result = System.Windows.MessageBox.Show(
                    owner: this,
                    messageBoxText: fullMessage,
                    caption: "Предупреждение: не отпущены клавиши/кнопки",
                    button: MessageBoxButton.YesNo,
                    icon: MessageBoxImage.Warning
                );

                if (result == MessageBoxResult.No)
                {
                    return;
                }
            }



            if (!Is_Console_App)
            {
                if (Main_Proc == null)
                {
                    error_APP_L.Content = $"Ошибка! Процесс {Patch_APP} отсутствует";
                    error_APP_L.Visibility = Visibility.Visible;
                    status_L.Content = "NONE";
                    status_L.Foreground = System.Windows.Media.Brushes.Red;
                    
                    return;
                }

                if (Main_Root == null)
                {
                    error_APP_L.Content = $"Ошибка при получении корневого элемента  {Main_Proc.ProcessName}  \r\n Автоматическое закрытие процесса";
                    error_APP_L.Visibility = Visibility.Visible;
                    Main_Proc.Kill();
                    return;
                }
                try
                {
                    DATE = new List<ElementRecord>();
                    DATE = UI_Scanner.SnapshotControls(Main_Root);
                    Console.WriteLine();
                }
                catch
                {
                    error_APP_L.Content = $"Ошибка при обработке корневого элемента  {Main_Proc.ProcessName}  \r\n Автоматическое закрытие процесса";
                    error_APP_L.Visibility = Visibility.Visible;
                    Main_Proc.Kill();
                    return;
                }


                foreach (var item in command_s)
                {
                    if (Main_Proc == null)
                    {
                        error_APP_L.Content = $"Ошибка! Процесс {Patch_APP} отсутствует";
                        error_APP_L.Visibility = Visibility.Visible;
                        status_L.Content = "NONE";
                        status_L.Foreground = System.Windows.Media.Brushes.Red;
                        return;
                    }
                    if (Main_Root == null)
                    {
                        error_APP_L.Content = $"Ошибка при получении корневого элемента  {Main_Proc.ProcessName}  \r\n Автоматическое закрытие процесса";
                        error_APP_L.Visibility = Visibility.Visible;
                        Main_Proc.Kill();
                        return;
                    }

                    // Проверка существования команды в словаре
                    if (!Commands.ContainsKey(item.Command_Name))
                    {
                        int lineNumber = 0;
                        int offset = code_TB.Document.Text.IndexOf(item.Full_Command);
                        if (offset >= 0)
                        {
                            DocumentLine line = code_TB.Document.GetLineByOffset(offset);
                            lineNumber = line.LineNumber;
                        }

                        System.Windows.MessageBox.Show(
                            owner: this,
                            messageBoxText: $"\tОшибка синтаксиса:\n неизвестная команда '{item.Command_Name}' в строке с номером {lineNumber}",
                            caption: "Парсер псевдокода",
                            button: MessageBoxButton.OK,
                            icon: MessageBoxImage.Error
                        );
                        return;
                    }

                    bool rez = true;
                    try
                    {
                        rez = Commands[item.Command_Name](item.args, item.Full_Command);
                    }
                    catch (Exception ex)
                    {
                        rez = false;
                        int lineNumber = 0;
                        int offset = code_TB.Document.Text.IndexOf(item.Full_Command);
                        if (offset >= 0)
                        {
                            DocumentLine line = code_TB.Document.GetLineByOffset(offset);
                            lineNumber = line.LineNumber;
                        }

                        System.Windows.MessageBox.Show(
                            owner: this,
                            messageBoxText: $"\tОшибка выполнения команды:\n в строке с номером {lineNumber}\n\nДетали: {ex.Message}",
                            caption: "Парсер псевдокода",
                            button: MessageBoxButton.OK,
                            icon: MessageBoxImage.Error
                        );
                    }

                    if (!rez)
                    {
                        return;
                    }
                }
            }

        }


        //                                                                                  обработчик времени на загрузку и закрытие
        private void TimeTextBox_TextChanged(object sender, TextChangedEventArgs e)
        {
            var tb = sender as System.Windows.Controls.TextBox;
            if (tb == null) return;

            int oldCaret = tb.CaretIndex;
            string oldText = tb.Text;

            
            string newText = new string(oldText.Where(char.IsDigit).ToArray());

            if (newText != oldText)
            {
                int removedLeft = oldText.Substring(0, oldCaret).Count(ch => !char.IsDigit(ch));

                tb.Text = newText;

                
                int newCaret = Math.Max(0, oldCaret - removedLeft);
                if (newCaret > newText.Length) newCaret = newText.Length;

                tb.CaretIndex = newCaret;
            }
        }

      
    }
}