using System;
using System.Windows.Forms;

namespace debet_kredit_xls
{
    static class Program
    {
        /// <summary>
        /// Главная точка входа для приложения
        /// </summary>
        [STAThread]
        static void Main()
        {
            Application.EnableVisualStyles(); // включаем визуальные стили для приложения
            Application.SetCompatibleTextRenderingDefault(false); //сохраняем старый стиль для старых версий .Net Framework
            Application.Run(new Form1());
        }
    }
}
