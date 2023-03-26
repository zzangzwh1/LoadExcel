using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;


namespace LoadExcel
{
    class Program
    {
        static void Main(string[] args)
        {
            Excel excel = new Excel();
            excel.WriteExcel();
            Console.ReadLine();
        }
    }
}
