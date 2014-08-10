using AutomationX;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace StaKoTecHomeGear
{
    static class Logging
    {
        private static AX _aX = null;
        private static AXInstance _mainInstance = null;

        public static void Init(AX ax, AXInstance mainInstance)
        {
            _aX = ax;
            _mainInstance = mainInstance;
        }

        public static void WriteLog(String message, String stackTrace = "")
        {
            String logPath = "d:\\StaKoTecHomeGear.txt";
            using (System.IO.StreamWriter file = new System.IO.StreamWriter(logPath, true))
            {
                file.WriteLine(DateTime.Now.ToString() + ": " + message);
            }
            _aX.WriteJournal(0, _mainInstance.Name, message, "ON", "HomeGear");
            Console.WriteLine(message);
            if (stackTrace.Length > 0)
            {
                _aX.WriteJournal(0, _mainInstance.Name, stackTrace, "ON", "HomeGear");
                Console.WriteLine(stackTrace);
                using (System.IO.StreamWriter file = new System.IO.StreamWriter(@logPath, true))
                {
                    file.WriteLine(DateTime.Now.ToString() + ": " + stackTrace);
                }
            }
            _mainInstance.Status = message;
        }
    }
}
