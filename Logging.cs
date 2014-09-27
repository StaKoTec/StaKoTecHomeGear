using AutomationX;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading;
using System.Threading.Tasks;

namespace StaKoTecHomeGear
{
    static class Logging
    {
        private static AX _aX = null;
        private static AXInstance _mainInstance = null;
        private static System.IO.StreamWriter _logWriter = null;

        public static void Init(AX ax, AXInstance mainInstance)
        {
            _aX = ax;
            _mainInstance = mainInstance;
            String logPath = _mainInstance.Get("$$_working_part").GetString() + "\\StaKoTecHomeGear.txt";
            _logWriter = new System.IO.StreamWriter(logPath, true, Encoding.UTF8, 1024);
            _logWriter.AutoFlush = true;
        }

        public static void WriteLog(String message, String stackTrace = "", Boolean setError = false)
        {
            try
            {
                _logWriter.WriteLine(DateTime.Now.ToString() + ": " + message);

                _aX.WriteJournal(0, _mainInstance.Name, message, "ON", "HomeGear");
                Console.WriteLine(message);
                if (stackTrace.Length > 0)
                {
                    _aX.WriteJournal(0, _mainInstance.Name, stackTrace, "ON", "HomeGear");
                    Console.WriteLine(stackTrace);
                    _logWriter.WriteLine(DateTime.Now.ToString() + ": " + stackTrace);
                }
                if (setError)
                    _mainInstance.Error = message;

                _mainInstance.Status = message;
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.Message + "\r\n" + ex.StackTrace);
            }
        }
    }
}
