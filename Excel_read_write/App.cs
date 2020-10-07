using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Excel_read_write
{
    class App
    {
        static void Main()
        {

            Excel excel = new Excel("C:\\Users\\Kavinda Ravishan\\source\\repos\\Excel_read_write\\Excel_read_write\\ex.xlsx", 1);

            for (int i = 0; i < 10; i++)
            {
                excel.WriteToCell(i, 0, "hello kavinda");
            }


            Console.WriteLine(excel.ReadExcel(0, 0));


            excel.Save();
            excel.Close();

            Console.Read();
        }
    }
}
