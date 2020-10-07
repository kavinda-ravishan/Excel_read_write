using System;

namespace Excel_read_write
{
    class App
    {
        static void creat(string path)
        {
            Excel excel = new Excel();
            excel.CreatNewFile();
            excel.SaveAs(path);
            excel.Close();
        }

        static void Main()
        {
            string path = "C:\\Users\\Kavinda Ravishan\\source\\repos\\kavinda-ravishan\\Excel_read_write\\Excel_read_write\\ex.xlsx";

            creat(path);

            Excel excel = new Excel(path, 1);

            //excel.CreatNewSheet();

            excel.SelectWorkSheet(1);

            excel.WriteToCell(0, 0, "Wave Lenght");
            excel.WriteToCell(0, 1, "Power");

            Random random = new Random();

            int waveLenght = 1540;
            for (int i = 1; i < 100; i++)
            {
                excel.WriteToCell(i, 0, waveLenght.ToString());
                excel.WriteToCell(i, 1, random.Next(0, 10).ToString());

                waveLenght++;
            }


            //Console.WriteLine(excel.ReadExcel(0, 0));

            //excel.DeleteWorkSheet(1);

            excel.Save();
            excel.Close();

            Console.WriteLine("DONE");
            Console.Read();
        }
    }
}
