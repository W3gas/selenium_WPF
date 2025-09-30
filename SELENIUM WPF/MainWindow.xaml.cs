
using System.IO;
using System.Drawing;
using Microsoft.Win32;
using System.Diagnostics;
using System.Windows;
using System.Windows.Automation;
using System.Windows.Controls;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Threading;
using System;
using System.Collections.Generic;
using System.Linq;


//     проверкаа гита шрывшгыкршпру
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
                        crash_L.Visibility = Visibility.Visible;
                        crash_Code_L.Content = Convert.ToString(code);
                        crash_Code_L.Visibility = Visibility.Visible;
                        start_B.IsEnabled = true;

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


        public int Time_Out_APP = 5000;  //  время на закрытие приложения (мc)
        public int Time_Load_APP = 10000;  //  время на открытие приложения (мc)

        public string Patch_APP = "";  //  путь к исполнительному файлу

        public bool Is_Closed = false;  //  закрыт ли процесс корректно
        public bool Is_Working = false;  //  работает ли процесс
        public bool Is_Run_Update = false;  //  работает ли апдейт


        public string This_Window_Name = "second_Form";  //  окно для обработки и поиска
        public string This_Control_Name = "open_Form_B";  // контрол для поиска и работы














        public MainWindow()
        {
            InitializeComponent();

            //  ссылка на окно для доступа 
            //******************************************
            GUI_Access.Main_W = this;
            //******************************************
<<<<<<< HEAD
            
            
=======

>>>>>>> refs/remotes/origin/master

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
            //******************************************

        }

        //AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA
        private void Button1_Click(object sender, RoutedEventArgs e)
        {                                                                //  УДАЛИТЬ НА*** В КОНЦЕ
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
        }







        //                                                                          ЗАПУСКАТОР
        private void start_B_Click(object sender, EventArgs e)
        {
            string ext = System.IO.Path.GetExtension(Patch_APP);

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
                        //   КАКИМ ТО БОКОМ ЗАКРЫВАТЬ ПРОЫОДНИК ЕСЛИ  ЯРЛЫК УКАЗЫВАЕТ НА ПАПКУ
                        return;
                    }

                    Macroses();
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
                        Console.WriteLine("ЕБАТЬ Я ЛОООХ"); 
                    }
                }

            }

        }


        //                                                                              ЖОПАБОЛЬ
        public void Macroses()
        {
            //  начальное получение доступа к GUI
            //$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$
            if (!GUI_Access.Is_Console_APP(Patch_APP))   //  проверка на консоль
            {
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
            //$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$


        }


        //                                                             завершить внешний процесс
        public async void STOP_APP()
        {

            ProcessCloser.CloseProcessGracefully(Main_Proc, Time_Out_APP);
            start_B.IsEnabled = true;

        }




        //                                                                       тесты
        private void temp_B_Click(object sender, RoutedEventArgs e)
        {
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



            Console.WriteLine();

            
            ElementRecord Clic_El = Find_By_Name(This_Control_Name);
            if (Clic_El == null)
            {
                error_APP_L.Content = $"Ошибка при зполучении элемента  {This_Control_Name}  \r\n Автоматическое закрытие процесса";
                error_APP_L.Visibility = Visibility.Visible;
                Main_Proc.Kill();
                return;
            }
            Activate_Main_Process_W(Main_Proc);
            //Thread.Sleep(100);
            DATE = new List<ElementRecord>();
            DATE = UI_Scanner.SnapshotControls(Main_Root);
            Clic_El = Find_By_Name(This_Control_Name);


            Mouse_Emulator.MoveAndClick(Clic_El.CenterPoint.X, Clic_El.CenterPoint.Y);
            Console.WriteLine();

            try
            {
                DATE = new List<ElementRecord>();
                Thread.Sleep(200);  //  возможно, стоит перекинуть внутрь функции активации главного окна
                DATE = WindowUITool.CaptureWindowUI(Main_Proc, This_Window_Name);
                Console.WriteLine();
            }
            catch
            {
                error_APP_L.Content = $"Ошибка при обработке корневого элемента окна  {This_Window_Name}  \r\n Автоматическое закрытие процесса";
                error_APP_L.Visibility = Visibility.Visible;
                Main_Proc.Kill();
                return;
            }



        }
    }
}