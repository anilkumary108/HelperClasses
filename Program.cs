using System;

namespace ConsoleTestApp
{
    public class Program
    {

        bool[] b = new bool[3];
        int count = 0;
        static void Main(string[] args)
        {


            ExcelHelper excelHelper = new ExcelHelper();
            //excelHelper.GenerateExcel();
            excelHelper.CreateExcelFile(@"F:\FileName.xls");

            var timeWindow = "16:00 - 20:00".Split('-');
            var date = "06/12/2017";

            var startDate = Convert.ToDateTime(date);
            Console.WriteLine(startDate.AddHours(Int32.Parse(timeWindow[0].Split(':')[0])));
            Console.WriteLine(startDate.AddHours(Int32.Parse(timeWindow[1].Split(':')[0])));

            //SyncAwaitExample sae = new SyncAwaitExample();
            //sae.callAsync();

            //int a = 5;
            //int b = 0, c = 0;

            //a = method(a, b, ref c);

            //Console.WriteLine(a + " " + b + " " + c);
            Console.WriteLine("Main method");
            Console.ReadLine();
            //DerivedClass objDC = new DerivedClass();
            //objDC.Method1();

            //BaseClass baseClass = new BaseClass();
            //baseClass.Method1();

            //BaseClass b = new DerivedClass();
            //b.Method1();
            //Console.ReadLine();
            Console.WriteLine(" END Main method");
        }

        private static int method(int x, int p, ref int k)
        {

            p = ++x + x*x++;
            k = x*x + p;
            return 0;
        }

    }

    // Base class
    public class BaseClass
    {
        public virtual void Method1()
    {
        Console.Write("Base Class Method");
    }
    }
    // Derived class
    public class DerivedClass : BaseClass
    {
        public override void Method1()
        {
            base.Method1();
            Console.Write("Derived Class Method");
        }
    }
}
