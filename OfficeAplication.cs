using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace WindowsFormsApp1
{
    internal class  AdvancedCursor
    {
        [System.Runtime.InteropServices.DllImport("User32.dll")]
        private static extern IntPtr LoadCursorFromFile(String str);
        public static Cursor Create(string filename)
        {
            IntPtr hCurosr = LoadCursorFromFile(filename);
                if(!IntPtr.Zero.Equals(hCurosr))
            {
                return new Cursor(hCurosr);

            }
            else
            {
                throw new ApplicationException("Ошибка загрузки курсора" + filename);
            }
        }
    }
}
