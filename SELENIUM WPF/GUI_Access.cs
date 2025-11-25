using System;
using System;
using System.Collections.Generic;
using System.Collections.Generic;
using System.Diagnostics;
using System.Diagnostics;
using System.Drawing;
using System.Drawing.Imaging;
using System.IO;
using System.Linq;
using System.Linq;
using System.Runtime.InteropServices;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading.Tasks;
using System.Threading.Tasks;
using System.Windows.Automation;
using System.Windows.Automation;
using static System.Windows.Forms.VisualStyles.VisualStyleElement;


namespace SELENIUM_WPF
{
    public static class GUI_Access
    {
        public static MainWindow Main_W { get; set; }

        //                                                                  Доступ к корню
        public static AutomationElement Get_Root(Process Main_Proc, int time)
        {
            try
            {
                return GetRootAsync(Main_Proc, time).Result;
            }
            catch
            {
                return null;
            }
        }


        //                                                               Доступ к корню (приват)
        private static async Task<AutomationElement> GetRootAsync(Process proc, int time)
        {
            await Task.Delay(time).ConfigureAwait(false);
            IntPtr h = proc.MainWindowHandle;
            return AutomationElement.FromHandle(h);
        }


        //                                                             Проверка на нраличие GUI
        public static bool Is_Console_APP(string proc)
        {
            try
            {
                string exePath = proc;

                using var fs = new FileStream(exePath, FileMode.Open, FileAccess.Read);
                using var reader = new BinaryReader(fs);

                fs.Seek(0x3C, SeekOrigin.Begin);
                int peHeaderOffset = reader.ReadInt32();

                fs.Seek(peHeaderOffset + 0x5C, SeekOrigin.Begin);
                ushort subsystem = reader.ReadUInt16();

                return subsystem switch
                {
                    2 => false,
                    3 => true,
                    _ => false,
                };
            }
            catch
            {
                return false;
            }
        }





    }


}



