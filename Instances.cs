using AutomationX;
using HomegearLib;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading;

namespace StaKoTecHomeGear
{
    public class Instances : Dictionary<Int32, AxInstance>
    {
        public delegate void VariableValueChangedEventHandler(AxVariable variable);

        public event VariableValueChangedEventHandler VariableValueChanged;
        public event VariableValueChangedEventHandler SubinstanceVariableValueChanged;

        protected Ax _aX = null;
        protected AxInstance _mainInstance = null;

        public Mutex _mutex = new Mutex();
        public Boolean mutexIsLocked = false;

        protected Int32 _polledVariablesCount = 0;
        public Int32 PolledVariablesCount { get { return _polledVariablesCount; } }

        public Instances(Ax ax, AxInstance mainInstance) 
        {
            _aX = ax;
            _mainInstance = mainInstance;
        }

        public void Lifetick()
        {
            lock (_mutex)
            {
                mutexIsLocked = true;
                if (Count > 0)
                {
                    foreach (KeyValuePair<Int32, AxInstance> instance in this)
                    {
                        try
                        {
                            instance.Value["Lifetick"].Set(true);
                        }
                        catch (Exception ex)
                        {
                            Logging.WriteLog(LogLevel.Error, instance.Value, "Couldn't set variable " + instance.Value.Name + " to true", ex.StackTrace);
                        }
                    }
                }
                mutexIsLocked = false;
            }
        }

        public void Reload(Devices homegearDevices)
        {
            lock (_mutex)
            {
                mutexIsLocked = true;
                try
                {
                    Logging.WriteLog(LogLevel.Debug, _mainInstance, "Hole alle HomeGear-Klassen");
                    List<String> homegearClasses = getHomeGearClasses();
                    Dictionary<String, List<String>> homegearInstances = getHomeGearInstances(homegearClasses);
                    Logging.WriteLog(LogLevel.Debug, _mainInstance, "Ab geht die Party");
                    _polledVariablesCount = 0;

                    ////////////////////////////////////////////////////
                    // NEU:

                    List<KeyValuePair<Int32, AxInstance>> instancesToRemove = new List<KeyValuePair<Int32, AxInstance>>();
                    List<KeyValuePair<Int32, AxInstance>> instancesToReload = new List<KeyValuePair<Int32, AxInstance>>();
                    foreach (KeyValuePair<Int32, AxInstance> pair in this)
                    {
                        if (pair.Value.CleanUp) instancesToRemove.Add(pair);
                        else if (pair.Value.ReloadRequired) instancesToReload.Add(pair);
                    }
                    foreach (KeyValuePair<Int32, AxInstance> instance in instancesToRemove)
                    {
                        instance.Value.Dispose();
                        this.Remove(instance.Key);
                    }
                    foreach (KeyValuePair<Int32, AxInstance> instance in instancesToReload)
                    {
                        instance.Value.Dispose();
                        this.Remove(instance.Key);
                        AxInstance newInstance = new AxInstance(_aX, instance.Value.Name);
                        this.Add(instance.Key, newInstance);
                        newInstance.VariableValueChanged += OnVariableValueChanged;
                        //newInstance.ArrayValueChanged += variable_OnArrayValueChanged;

                        AxVariable[] variables = newInstance.Variables;
                        foreach (AxVariable variable in variables)
                            variable.Events = true;

                        AxInstance[] subinstances = newInstance.Subinstances;
                        foreach (AxInstance subinstance in subinstances)
                        {
                            subinstance.VariableValueChanged += OnSubinstanceVariableValueChanged;
                            //subinstance.ArrayValueChanged += OnArrayValueChanged;
                            //this.Add(subinstance.Path, subinstance);
                            variables = subinstance.Variables;
                            foreach (AxVariable variable in variables)
                                variable.Events = true;
                        }
                    }




                    ///////////////////////////////////////////////////////////////
                    //////////////////////////////////////////////////////////////
                    // Alle Instanzen im aX nach ihrer ID abfragen und in Homegear.Devices suchen
                    //Alt:
                    foreach (KeyValuePair<String, List<String>> aktInstancePair in homegearInstances)
                    {
                        foreach (String aktaXInstanceName in aktInstancePair.Value)
                        {
                            AxInstance testInstance = new AxInstance(_aX, aktaXInstanceName);
                            Int32 aktID = testInstance.Get("ID").GetLongInteger();
                            if ((aktID <= 0) || (!homegearDevices.ContainsKey(aktID)))  //Wenn eine Instanz frisch im aX instanziert wurde und keine ID vergeben wurde, ist die ID -1 
                            {
                                Logging.WriteLog(LogLevel.Debug, _mainInstance, "Ignoriere Instanz " + testInstance.Name + " mit ID " + aktID.ToString() + ", da nicht in homegearDevices.");
                                testInstance.Dispose();
                                continue;
                            }
                            Logging.WriteLog(LogLevel.Debug, _mainInstance, "Füge Instanz (" + aktaXInstanceName + ") " + testInstance.Name + " mit ID " + aktID.ToString() + " hinzu.");
                            Add(aktID, testInstance);
                        }
                    }
                }
                catch (Exception ex)
                {
                    Logging.WriteLog(LogLevel.Error, _mainInstance, ex.Message, ex.StackTrace);
                }
                mutexIsLocked = false;
            }
        }


        private void OnVariableValueChanged(AxVariable sender, AxVariableValue value, DateTime timestamp)
        {
            if (VariableValueChanged != null)
                VariableValueChanged(sender);
        }

        private void OnSubinstanceVariableValueChanged(AxVariable sender, AxVariableValue value, DateTime timestamp)
        {
            if (SubinstanceVariableValueChanged != null)
                SubinstanceVariableValueChanged(sender);
        }



        protected List<String> getHomeGearClasses()
        {
            List<String> homeGearClassNames = new List<String>();
            try
            {
                AxVariable homeGearKlassen = _mainInstance.Get("HomeGearKlassen");
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
                Logging.WriteLog(LogLevel.Error, _mainInstance, ex.Message, ex.StackTrace);
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
                    Logging.WriteLog(LogLevel.Debug, _mainInstance, "Istanzen von Klasse " + aktHomegearClass + ":");
                    foreach(String name in aXInstanceNames)
                        Logging.WriteLog(LogLevel.Debug, _mainInstance, name);
                    if (!homegearInstances.ContainsKey(aktHomegearClass))
                        homegearInstances.Add(aktHomegearClass, aXInstanceNames);
                }
            }
            catch (Exception ex)
            {
                Logging.WriteLog(LogLevel.Error, _mainInstance, ex.Message, ex.StackTrace);
            }
            return (homegearInstances);
        }
    }
}
