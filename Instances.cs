using AutomationX;
using HomegearLib;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading;

namespace StaKoTecHomeGear
{
    public class Instances : Dictionary<Int32, AXInstance>
    {
        public delegate void VariableValueChangedEventHandler(AXVariable variable);

        public event VariableValueChangedEventHandler VariableValueChanged;
        public event VariableValueChangedEventHandler SubinstanceVariableValueChanged;

        protected AX _aX = null;
        protected AXInstance _mainInstance = null;

        protected Mutex _mutex = new Mutex();
        public Boolean MutexLocked { set { if (value) _mutex.WaitOne(); else _mutex.ReleaseMutex(); } }

        protected Int32 _polledVariablesCount = 0;
        public Int32 PolledVariablesCount { get { return _polledVariablesCount; } }

        public Instances(AX ax, AXInstance mainInstance) 
        {
            _aX = ax;
            _mainInstance = mainInstance;
        }

        public void Lifetick()
        {
            _mutex.WaitOne();
            if (Count > 0)
            {
                foreach (KeyValuePair<Int32, AXInstance> instance in this)
                {
                    try
                    {
                        instance.Value.Get("Lifetick").Set(true);
                    }
                    catch (Exception ex)
                    {
                        Logging.WriteLog(LogLevel.Error, "Couldn't set variable " + instance.Value.Name + " to true", ex.StackTrace);
                    }
                }
            }
            _mutex.ReleaseMutex();
        }

        public void Reload(Devices homegearDevices)
        {
            _mutex.WaitOne();
            try
            {
                Logging.WriteLog(LogLevel.Debug, "Lösche instanzen-handles");
                Clear(false);

                //////////////////////////////////////////////////////////////
                // Alle Instanzen im aX nach ihrer ID abfragen und in Homegear.Devices suchen
                Logging.WriteLog(LogLevel.Debug, "Hole alle HomeGear-Klassen");
                List<String> homegearClasses = getHomeGearClasses();
                Dictionary<String, List<String>> homegearInstances = getHomeGearInstances(homegearClasses);
                List<String> instancesToDispose = new List<String>();
                Logging.WriteLog(LogLevel.Debug, "Ab geht die Party");
                _polledVariablesCount = 0;

                foreach (KeyValuePair<String, List<String>> aktInstancePair in homegearInstances)
                {
                    foreach (String aktaXInstanceName in aktInstancePair.Value)
                    {
                        AXInstance testInstance = new AXInstance(_aX, aktaXInstanceName, "Status", "err");
                        
                        //Vorbereitung auf prüfung ob Instanz schon vorhanden ist. Dann kann auch das Clear(false) von oben raus
                        Boolean instanceVorhanden = false;
                        foreach(KeyValuePair<Int32, AXInstance> testPair in this)
                        {
                            //Logging.WriteLog("if  " + testPair.Value.Name + " == " + aktaXInstanceName);
                            if (testPair.Value.Name == aktaXInstanceName)
                            {
                                instanceVorhanden = true;

                                //Logging.WriteLog(aktaXInstanceName + " gibt's schon!...");
                                _polledVariablesCount += testPair.Value.PolledVariablesCount;
                                foreach (AXInstance testSubInstance in testPair.Value.Subinstances)
                                    _polledVariablesCount += testSubInstance.PolledVariablesCount;

                                break;
                            }
                        }
                        if (instanceVorhanden)
                            continue;

                        //Logging.WriteLog(aktaXInstanceName + " ist neu. Hole handles...");

                        Int32 aktID = testInstance.Get("ID").GetLongInteger();
                        if ((aktID <= 0) || (!homegearDevices.ContainsKey(aktID)))  //Wenn eine Instanz frisch im aX instanziert wurde und keine ID vergeben wurde, ist die ID -1 
                        {
                            testInstance.Dispose();
                            continue;
                        }
                        Add(aktID, testInstance);
                        if (aktInstancePair.Key != homegearDevices[aktID].TypeString)
                            continue;
                        testInstance.SetVariableEvents(true);
                        testInstance.PollingInterval = 20;
                        testInstance.VariableValueChanged += OnVariableValueChanged;
                        _polledVariablesCount += testInstance.PolledVariablesCount;
                        ////////////////////////////////////////////////////////////////////
                        // Subinstanzen suchen und Ereignishandler hinzufügen
                        foreach (AXInstance aktSubinstance in testInstance.Subinstances)
                        {
                            aktSubinstance.SetVariableEvents(true);
                            aktSubinstance.PollingInterval = 20;
                            aktSubinstance.VariableValueChanged += OnSubinstanceVariableValueChanged;
                            _polledVariablesCount += aktSubinstance.PolledVariablesCount;
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                Logging.WriteLog(LogLevel.Error, ex.Message, ex.StackTrace);
            }
            _mutex.ReleaseMutex();
        }

        public new void Clear()
        {
            throw new NotImplementedException("You need to pass the parameter \"lockMutex\".");
        }

        public void Clear(bool lockMutex)
        {
            if (lockMutex)
                _mutex.WaitOne();
            try
            {
                foreach (KeyValuePair<Int32, AXInstance> instancePair in this)
                {
                    try
                    {
                        instancePair.Value.VariableValueChanged -= OnVariableValueChanged;
                        foreach (AXInstance aktSubinstance in instancePair.Value.Subinstances)
                        {
                            aktSubinstance.VariableValueChanged -= OnSubinstanceVariableValueChanged;
                        }
                        instancePair.Value.Dispose();
                    }
                    catch (Exception ex)
                    {
                        Logging.WriteLog(LogLevel.Error, ex.Message, ex.StackTrace);
                    }
                    _mainInstance.Get("Lifetick").Set(true);
                }
                _polledVariablesCount = 0;
                base.Clear();
            }
            catch (Exception ex)
            {
                Logging.WriteLog(LogLevel.Error, ex.Message, ex.StackTrace);
            }
            if (lockMutex)
                _mutex.ReleaseMutex();
        }

        public new bool Remove(Int32 key)
        {
            throw new NotImplementedException("You need to pass the parameter \"lockMutex\".");
        }

        public void Remove(Int32 key, bool lockMutex)
        {
            if (lockMutex) _mutex.WaitOne();
            try
            {
                if (ContainsKey(key))
                {
                    AXInstance removeInstance = this[key];
                    base.Remove(key);
                    removeInstance.VariableValueChanged -= OnVariableValueChanged;
                    foreach (AXInstance aktSubinstance in removeInstance.Subinstances)
                    {
                        aktSubinstance.VariableValueChanged -= OnSubinstanceVariableValueChanged;
                    }
                    removeInstance.Dispose();
                }
            }
            catch (Exception ex)
            {
                Logging.WriteLog(LogLevel.Error, ex.Message, ex.StackTrace);
            }
            if (lockMutex) _mutex.ReleaseMutex();
        }

        private void OnVariableValueChanged(AXVariable sender)
        {
            if (VariableValueChanged != null)
                VariableValueChanged(sender);
        }

        private void OnSubinstanceVariableValueChanged(AXVariable sender)
        {
            if (SubinstanceVariableValueChanged != null)
                SubinstanceVariableValueChanged(sender);
        }


        protected List<String> getHomeGearClasses()
        {
            List<String> homeGearClassNames = new List<String>();
            try
            {
                AXVariable homeGearKlassen = _mainInstance.Get("HomeGearKlassen");
                List<String> classNames = _aX.GetClassNames();

                foreach (String name in classNames)
                {
                    String[] tempName = name.Split('/');
                    if (tempName.Length < 3) continue;
                    String system = tempName[tempName.Length - 3];
                    String aktName = "";
                    if (system == "HomeGear")  //Klassen müssen im Verzeichnis HomeGear sein
                    {
                        aktName = tempName[tempName.Length - 1].Substring(0, tempName[tempName.Length - 1].Length - 7);
                        if (aktName == "HomeGear")  //Die HomeGear-Verwaltungs heisst HomeGear - sie sol nicht in die Klassen-übersicht
                            continue;

                        if (!homeGearClassNames.Contains(aktName))
                            homeGearClassNames.Add(aktName);
                    }
                }
                homeGearClassNames.Sort();

                UInt16 i = 0;
                foreach (String name in homeGearClassNames)
                {
                    if (i > homeGearKlassen.Length)
                        break;
                    homeGearKlassen.Set(i, name);
                    i++;
                }

                for (; i < homeGearKlassen.Length; i++)
                    homeGearKlassen.Set(i, "");
            }
            catch (Exception ex)
            {
                Logging.WriteLog(LogLevel.Error, ex.Message, ex.StackTrace);
            }

            return homeGearClassNames;
        }


        protected Dictionary<String, List<String>> getHomeGearInstances(List<String> homegearClasses)
        {
            Dictionary<String, List<String>> homegearInstances = new Dictionary<String, List<String>>();
            try
            {
                foreach (String aktHomegearClass in homegearClasses)
                {
                    List<String> aXInstanceNames = _aX.GetInstanceNames(aktHomegearClass);
                    if (!homegearInstances.ContainsKey(aktHomegearClass))
                        homegearInstances.Add(aktHomegearClass, aXInstanceNames);
                }
            }
            catch (Exception ex)
            {
                Logging.WriteLog(LogLevel.Error, ex.Message, ex.StackTrace);
            }
            return (homegearInstances);
        }
    }
}
