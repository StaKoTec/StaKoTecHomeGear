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
        private static Ax _aX = null;
        private static AxInstance _mainInstance = null;
        private static System.IO.StreamWriter _logWriter = null;

        public static void Init(Ax ax, AxInstance mainInstance)
        {
            _aX = ax;
            _mainInstance = mainInstance;
            String DatumString = DateTime.Now.Year.ToString() + DateTime.Now.Month.ToString("D2") + DateTime.Now.Day.ToString("D2") + DateTime.Now.Hour.ToString("D2") + DateTime.Now.Minute.ToString("D2") + DateTime.Now.Second.ToString("D2");
            String logPath = _mainInstance.Get("$$_working_part").GetString() + "\\Log_" + mainInstance.Name + "_" + DatumString + ".txt";
            _logWriter = new System.IO.StreamWriter(logPath, true, Encoding.UTF8, 1024);
            _logWriter.AutoFlush = true;
        }

        public static void WriteLog(LogLevel logLevel, AxInstance instance, String message, String stackTrace = "")
        {
            try
            {
                if ((logLevel > (LogLevel)_mainInstance.Get("LogLevel").GetLongInteger()) && !(logLevel == LogLevel.Always))
                    return;

                String prefix = "";
                Int32 position = 0;
                switch (logLevel)
                {
                    case LogLevel.Always:
                        prefix = "";
                        position = 10;
                        break;
                    case LogLevel.Debug:
                        prefix = "Debug: ";
                        position = 1;
                        break;
                    case LogLevel.Error:
                        prefix = "Error: ";
                        position = 0;
                        break;
                    case LogLevel.Info:
                        prefix = "Info: ";
                        position = 20;
                        break;
                    case LogLevel.Warning:
                        prefix = "Warning: ";
                        position = 2;
                        break;
                    default:
                        prefix = "";
                        position = 0;
                        break;
                }
                _logWriter.WriteLine(DateTime.Now.ToString() + "." + DateTime.Now.Millisecond.ToString("D3") + ": (" + instance.Name + ") " + prefix + message);
                _aX.WriteJournal(position, instance.Name, prefix + message, "ON", _mainInstance.Name);
                Console.WriteLine(prefix + message);
                if (stackTrace.Length > 0)
                {
                    _aX.WriteJournal(position, instance.Name, stackTrace, "ON", _mainInstance.Name);
                    Console.WriteLine(stackTrace);
                    _logWriter.WriteLine(DateTime.Now.ToString() + "." + DateTime.Now.Millisecond.ToString("D3") + ": " + stackTrace);
                }
                if (logLevel == LogLevel.Error)
                {
                    if (instance.VariableExists("err"))
                    {
                        _mainInstance["err.TEXT"].Set(message);
                        _mainInstance["err"].Set(true);
                    }
                }
                else
                {
                    if (instance.VariableExists("Status"))
                        _mainInstance["Status"].Set(message);
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.Message + "\r\n" + ex.StackTrace);
            }
        }
    }
}
