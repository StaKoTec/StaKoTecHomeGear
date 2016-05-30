using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace StaKoTecHomeGear
{
    class Program
    {
        static void Main(string[] args)
        {
            if (args.Length < 1)
            {
                Console.WriteLine("Too less arguments! Example: StaKoTecHomeGear.exe Instancename");
                Environment.Exit(-1);
            }

            string instance = args[0];

            String version = "0.6";
            App app = new App();
            app.Run(instance, version);
        }
    }
}
