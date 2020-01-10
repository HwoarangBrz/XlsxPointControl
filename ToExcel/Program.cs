using System;

namespace ToExcel
{
    class Program
    {
        static void Main(string[] args)
        {
            Console.WriteLine("Informe o mês/ano (mm/yyyy) de referência:");
            string val = Console.ReadLine();

            DateTime temp;
            if (DateTime.TryParse(val, out temp))
            {
                CreateExcelFile.GenerateWorksheet(temp);
            }
            else
            {
                Console.WriteLine("Data inválida!");
            }
            Console.WriteLine("Press any key!");
            Console.ReadKey();
        }
    }
}
