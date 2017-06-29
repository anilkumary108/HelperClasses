using System;
using System.Threading.Tasks;
using SyncTest;

namespace ConsoleTestApp
{
    class SyncAwaitExample
    {
        public void callAsync()
        {
            AsyncClass.ProcessDataAsync();
         }
    }
}

namespace SyncTest
{
    internal class AsyncClass
    {
        internal static async void ProcessDataAsync()
        {
            // This method runs asynchronously.
            Console.WriteLine("  START  ");

            int waitresults = await Task.Run(() => GetNumber());
            int t = waitresults;
            Console.WriteLine("Get Number: " + t);
            Console.WriteLine(" END Async");
        }

        internal static int GetNumber()
        {
            // Compute total count of digits in strings.
            int size = 0;
            for (int z = 0; z < 100; z++)
            {
                for (int i = 0; i < 1000000; i++)
                {
                    string value = i.ToString();
                    if (value == null)
                    {
                        return 0;
                    }
                    size += value.Length;
                }
            }
            return size;
        }
    }
}
