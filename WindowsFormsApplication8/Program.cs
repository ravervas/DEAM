using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace WindowsFormsApplication8
{
    static class Program
    {
        public static Form3 f3;   // переменная, которая будет содержать ссылку на форму Form3
        public static Form1 f1;
        public static Form2 f2;
        public static Form4 f4;
        public static Form5 f5;
        public static Form7 f7;
        public static Form8 f8;
        public static Form9 f9;
        public static Form10 f10;
        public static Form11 f11;
        public static Form12 f12;
        public static Form13 f13;
        public static Form14 f14;
        public static Form15 f15;
        /// <summary>
        /// Главная точка входа для приложения.
        /// </summary>
        [STAThread]
        static void Main()
        {           
            Application.EnableVisualStyles();
            Application.SetCompatibleTextRenderingDefault(false);
            Form6 first = new Form6();
            DateTime end = DateTime.Now + TimeSpan.FromSeconds(3);
            first.Show();
            while (end > DateTime.Now)
                Application.DoEvents();
            first.Close();
            first.Dispose();
            Application.Run(new Form1());
        }
    }
}
