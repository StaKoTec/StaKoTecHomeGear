using HomegearLib;
using HomegearLib.RPC;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Threading;
using AutomationX;
using System.Security.Authentication;

namespace StaKoTecHomeGear
{
    enum DeviceStatus
    {
        Nichts = 0,
        OK = 1,
        Fehler = 2,
        KeineInstanzVorhanden = 3
    }

    class App
    {
        //Globale Variablen
        bool _disposing = false;
        bool _initCompleted = false;
        AX _aX = null;
        AXInstance _mainInstance = null;
        Int32 _polledVariablesCount = 0;

        VariableConverter _varConverter = null;

        AXVariable _deviceID = null;
        AXVariable _deviceInstance = null;
        AXVariable _deviceRemark = null;
        AXVariable _deviceTypeString = null;
        AXVariable _deviceState = null;
        AXVariable _deviceStateColor = null;
        HomegearLib.RPC.RPCController _rpc = null;
        HomegearLib.Homegear _homegear = null;

        Dictionary<Int32, AXInstance> _instanzen = null;
        Mutex _homegearDevicesMutex = new Mutex();
        Mutex _instanzenMutex = new Mutex();


        public void Run(String instanceName)
        {
            try
            {
                try
                {
                    _aX = new AX();
                }
                catch (AXException ex)
                {
                    Console.WriteLine(ex.Message + "\r\n" + ex.StackTrace);
                    Dispose();
                }


                _aX.ShuttingDown += aX_ShuttingDown;
                _aX.SpsIdChanged += _aX_SpsIdChanged;

                _mainInstance = new AXInstance(_aX, instanceName, "Status", "err"); //Instanz-Objekt erstellen
                _mainInstance.PollingInterval = 1000;
                _mainInstance.SetVariableEvents(true);
                Logging.Init(_aX, _mainInstance);
                _varConverter = new VariableConverter(_mainInstance);

                AXVariable init = _mainInstance.Get("Init");
                init.Set(false);
                init.ValueChanged += init_ValueChanged;
                AXVariable pairingMode = _mainInstance.Get("PairingMode");
                pairingMode.Set(false);
                pairingMode.ValueChanged += pairingMode_ValueChanged;
                AXVariable deviceUnpair = _mainInstance.Get("DeviceUnpair");
                deviceUnpair.Set(false);
                deviceUnpair.ValueChanged += deviceUnpair_ValueChanged;
                AXVariable deviceReset = _mainInstance.Get("DeviceReset");
                deviceReset.Set(false);
                deviceReset.ValueChanged += deviceReset_ValueChanged;
                AXVariable deviceRemove = _mainInstance.Get("DeviceRemove");
                deviceRemove.Set(false);
                deviceRemove.ValueChanged += deviceRemove_ValueChanged;
                AXVariable getDeviceVars = _mainInstance.Get("GetDeviceVars");
                getDeviceVars.Set(false);
                getDeviceVars.ValueChanged += getDeviceVars_ValueChanged;
                AXVariable getDeviceConfigVars = _mainInstance.Get("GetDeviceConfigVars");
                getDeviceConfigVars.Set(false);
                getDeviceConfigVars.ValueChanged += getDeviceVars_ValueChanged;
                AXVariable axStartID = _mainInstance.Get("StartID");
                axStartID.ValueChanged += axStartID_ValueChanged;

                AXVariable lifetick = _mainInstance.Get("Lifetick");
                AXVariable aXcycleCounter = _mainInstance.Get("CycleCounter");
                Int32 cycleCounter = 0;
                Int32 i = 0;

                

                AXVariable start_CAPI_Release = _mainInstance.Get("Start_CAPI_Release");
                start_CAPI_Release.ValueChanged += start_CAPI_Release_ValueChanged;
                if (!start_CAPI_Release.GetBool())
                    Dispose();

                _mainInstance.Get("RPC_InitComplete").Set(false);
                _mainInstance.Get("CAPI_Running").Set(true);

                Logging.WriteLog("HomeGear started");

                HomeGearConnect();

                UInt32 j = 0;
                while (!_disposing)
                {
                    try
                    {
                        lifetick.Set(true);
                        aXcycleCounter.Set(cycleCounter);
                        cycleCounter++;

                        if (!_rpc.IsConnected)
                        {
                            Console.WriteLine(i.ToString() + ": Waiting for RPC-Server connection");
                            _mainInstance.Status = "Waiting for RPC-Server connection...";
                            Thread.Sleep(5000);
                            i++;
                            continue;
                        }

                        //Allen Devices einen Lifetick senden um DataValid zu generieren
                        _instanzenMutex.WaitOne();
                        if (_instanzen != null && _instanzen.Count() > 0)
                        {
                            foreach (KeyValuePair<Int32, AXInstance> instance in _instanzen)
                            {
                                instance.Value.Get("Lifetick").Set(true);
                            }
                        }
                        _instanzenMutex.ReleaseMutex();

                        if (_homegear != null && _initCompleted && j % 10 == 0)
                        {
                            Boolean serviceMessageVorhanden = false;
                            List<ServiceMessage> serviceMessages = _homegear.ServiceMessages;
                            UInt16 x = 0;
                            AXVariable aX_serviceMessages = _mainInstance.Get("ServiceMessages");
                            foreach (ServiceMessage message in serviceMessages)
                            {
                                String aktMessage = "Device ID: " + message.PeerID.ToString() + " " + "Channel: " + message.Channel.ToString() + " " + "Type: " + message.Type + " " + "Value: " + message.Value.ToString();
                                aX_serviceMessages.Set(x, aktMessage);
                                x++;
                                serviceMessageVorhanden = true;
                            }
                            for (; x < aX_serviceMessages.Length; x++)
                                aX_serviceMessages.Set(x, "");

                            _mainInstance.Get("ServiceMessageVorhanden").Set(serviceMessageVorhanden);
                        }

                        //Alle 60 Sekunden checken wie lange es her ist, dass die letzten Pakete Empfangen / gesendet wurden
                        Int32 currentTime = (Int32)(DateTime.UtcNow.Subtract(new DateTime(1970, 1, 1))).TotalSeconds;
                        if (_homegear != null && _initCompleted && j % 60 == 0)
                        {
                            foreach(KeyValuePair<String, Interface> aktInterface in _rpc.ListInterfaces())
                            {
                                Console.WriteLine("Interface " + aktInterface.Value.ID + ": LastPacketReceived: " + (currentTime - aktInterface.Value.LastPacketReceived).ToString() + "s ago");
                                Console.WriteLine("Interface " + aktInterface.Value.ID + ": LastPacketSent: " + (currentTime - aktInterface.Value.LastPacketSent).ToString() + "s ago");
                            }
                        }

                        j++;
                    }
                    catch (Exception ex)
                    {
                        Logging.WriteLog(ex.Message, ex.StackTrace);
                    }
                    Thread.Sleep(1000);
                }
            }
            catch(Exception ex)
            {
                Logging.WriteLog(ex.Message, ex.StackTrace);
            }
        }

        void axStartID_ValueChanged(AXVariable sender)
        {
            try
            {
                Logging.WriteLog("StartID has changed! Exiting!!!");
                Dispose();
            }
            catch (Exception ex)
            {
                Logging.WriteLog(ex.Message, ex.StackTrace);
            }
        }

        
        void getDeviceVars_ValueChanged(AXVariable sender)
        {
            try
            {
                AXVariable deviceVars_Name = _mainInstance.Get("DeviceVars_Name");
                AXVariable deviceVars_Type = _mainInstance.Get("DeviceVars_Type");
                AXVariable deviceVars_Min = _mainInstance.Get("DeviceVars_Min");
                AXVariable deviceVars_Max = _mainInstance.Get("DeviceVars_Max");
                AXVariable deviceVars_Default = _mainInstance.Get("DeviceVars_Default");
                AXVariable deviceVars_RW = _mainInstance.Get("DeviceVars_RW");
                UInt16 x = 0;
                Int32 deviceID = _mainInstance.Get("ActionID").GetLongInteger();
       
                _homegearDevicesMutex.WaitOne();
                if (_homegear.Devices.ContainsKey(deviceID))
                {
                    Device aktDevice = _homegear.Devices[deviceID];
                    foreach (KeyValuePair<Int32, Channel> aktChannel in aktDevice.Channels)
                    {
                        if (sender.Name == "GetDeviceConfigVars")
                        {
                            sender.Instance.Status = "Get VariableConfigNames for DeviceID: " + deviceID.ToString();
                            foreach (KeyValuePair<String, ConfigParameter> configName in aktDevice.Channels[aktChannel.Key].Config)
                            {
                                if (x >= deviceVars_Name.Length)
                                {
                                    _mainInstance.Error = "Array-Index zu klein bei 'DeviceVars_Name'";
                                    Logging.WriteLog("Array-Index zu klein bei 'DeviceVars_Name'");
                                    _homegearDevicesMutex.ReleaseMutex();
                                    return;
                                }

                                var aktVar = configName.Value;
                                String minVar = "";
                                String maxVar = "";
                                String typ = "";
                                String defaultVar = "";
                                String rwVar = "";
                                _varConverter.ParseDeviceConfigVars(aktVar, out minVar, out maxVar, out typ, out defaultVar, out rwVar);

                                deviceVars_Name.Set(x, aktVar.Name + "_C" + aktChannel.Key.ToString("D2"));
                                deviceVars_Type.Set(x, typ);
                                deviceVars_Min.Set(x, minVar);
                                deviceVars_Max.Set(x, maxVar);
                                deviceVars_Default.Set(x, defaultVar);
                                deviceVars_RW.Set(x, rwVar);

                                x++;
                            }
                        }

                        if (sender.Name == "GetDeviceVars")
                        {
                            sender.Instance.Status = "Get VariableNames for DeviceID: " + deviceID.ToString();
                            foreach (KeyValuePair<String, Variable> varName in aktDevice.Channels[aktChannel.Key].Variables)
                            {
                                if (x >= deviceVars_Name.Length)
                                {
                                    _mainInstance.Error = "Array-Index zu klein bei 'DeviceVars_Name'";
                                    Logging.WriteLog("Array-Index zu klein bei 'DeviceVars_Name'");
                                    _homegearDevicesMutex.ReleaseMutex();
                                    return;
                                }

                                var aktVar = varName.Value;
                                String minVar = "";
                                String maxVar = "";
                                String typ = "";
                                String defaultVar = "";
                                String rwVar = "";
                                _varConverter.ParseDeviceVars(aktVar, out minVar, out maxVar, out typ, out defaultVar, out rwVar);

                                deviceVars_Name.Set(x, aktVar.Name + "_V" + aktChannel.Key.ToString("D2"));
                                deviceVars_Type.Set(x, typ);
                                deviceVars_Min.Set(x, minVar);
                                deviceVars_Max.Set(x, maxVar);
                                deviceVars_Default.Set(x, defaultVar);
                                deviceVars_RW.Set(x, rwVar);

                                x++;
                            }
                        }
                    }

                    for (; x < deviceVars_Name.Length; x++)
                    {
                        deviceVars_Name.Set(x, "");
                        deviceVars_Type.Set(x, "");
                        deviceVars_Min.Set(x, "");
                        deviceVars_Max.Set(x, "");
                        deviceVars_Default.Set(x, "");
                        deviceVars_RW.Set(x, "");
                    }
                }
                _homegearDevicesMutex.ReleaseMutex();
            }
            catch (Exception ex)
            {
                _homegearDevicesMutex.ReleaseMutex();
                sender.Instance.Error = ex.Message;
                sender.Instance.Status = ex.Message;
                Logging.WriteLog(ex.Message, ex.StackTrace);
            }
            finally
            {
                try
                {
                    sender.Set(false);
                }
                catch (Exception ex)
                {
                    Logging.WriteLog(ex.Message, ex.StackTrace);
                }
            }
        }

        void deviceRemove_ValueChanged(AXVariable sender)
        {
            try
            {
                Int32 deviceID = _mainInstance.Get("ActionID").GetLongInteger();
                sender.Instance.Status = "Removing Device ID " + deviceID.ToString();
                _homegearDevicesMutex.WaitOne();

                _rpc.DeleteDevice(deviceID, RPCDeleteDeviceFlags.Force);

                _homegearDevicesMutex.ReleaseMutex();
                sender.Set(false);
                sender.Instance.Status = "Removing Device ID " + deviceID.ToString() + " complete";
            }
            catch (Exception ex)
            {
                _homegearDevicesMutex.ReleaseMutex();
                sender.Instance.Error = ex.Message;
                sender.Instance.Status = ex.Message;
                Logging.WriteLog(ex.Message, ex.StackTrace);
            }
        }

        void deviceReset_ValueChanged(AXVariable sender)
        {
            try
            {
                Int32 deviceID = _mainInstance.Get("ActionID").GetLongInteger();
                sender.Instance.Status = "Resetting Device ID " + deviceID.ToString();
                _homegearDevicesMutex.WaitOne();

                _rpc.DeleteDevice(deviceID, RPCDeleteDeviceFlags.Reset | RPCDeleteDeviceFlags.Defer);

                _homegearDevicesMutex.ReleaseMutex();
                sender.Set(false);
                sender.Instance.Status = "Resetting Device ID " + deviceID.ToString() + " complete";
            }
            catch (Exception ex)
            {
                _homegearDevicesMutex.ReleaseMutex();
                sender.Instance.Error = ex.Message;
                sender.Instance.Status = ex.Message;
                Logging.WriteLog(ex.Message, ex.StackTrace);
            }
        }

        void deviceUnpair_ValueChanged(AXVariable sender)
        {
            try
            {
                Int32 deviceID = _mainInstance.Get("ActionID").GetLongInteger();
                sender.Instance.Status = "Unpairing Device ID " + deviceID.ToString();
                _homegearDevicesMutex.WaitOne();

                _rpc.DeleteDevice(deviceID, RPCDeleteDeviceFlags.Defer);

                _homegearDevicesMutex.ReleaseMutex();
                sender.Set(false);
                sender.Instance.Status = "Unpairing Device ID " + deviceID.ToString() + " complete";
            }
            catch (Exception ex)
            {
                _homegearDevicesMutex.ReleaseMutex();
                sender.Instance.Error = ex.Message;
                sender.Instance.Status = ex.Message;
                Logging.WriteLog(ex.Message, ex.StackTrace);
            }
        }

        void pairingMode_ValueChanged(AXVariable sender)
        {
            try
            {
                _homegear.EnablePairingMode(sender.GetBool());
            }
            catch (Exception ex)
            {
                sender.Instance.Error = ex.Message;
                sender.Instance.Status = ex.Message;
                Logging.WriteLog(ex.Message, ex.StackTrace);
            }
        }

        void start_CAPI_Release_ValueChanged(AXVariable sender)
        {
            try
            {
                if (!sender.GetBool())
                {
                    _mainInstance.Status = "Beende StaKoTecHomeGear.exe";
                    _mainInstance.Get("err").Set(false);
                    Logging.WriteLog("Beende StaKoTecHomeGear.exe");

                    Dispose();
                }
            }
            catch (Exception ex)
            {
                sender.Instance.Error = ex.Message;
                sender.Instance.Status = ex.Message;
                Logging.WriteLog(ex.Message, ex.StackTrace);
            }
        }


        String GetClassname(String instanceName)
        {
            String className = "";
            try
            {
                String temp = _aX.GetClassPath(instanceName);
                if (temp.Length > 8)
                {
                    temp = temp.Replace(".symbol", "");
                    char[] delimiterChars = { '\\', '/' };
                    string[] teile = temp.Split(delimiterChars);
                    className = teile.Last();
                }
            }
            catch (Exception ex)
            {
                _mainInstance.Error = ex.Message;
                _mainInstance.Status = ex.Message;
                Logging.WriteLog(ex.Message, ex.StackTrace);
            }
            return className;
        }


        List<String> getHomeGearClasses()
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
                UInt16 x = 0;
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
                Logging.WriteLog(ex.Message, ex.StackTrace);
            }

            return homeGearClassNames;
        }

        void init_ValueChanged(AXVariable sender)
        {
            try
            {
                if (!sender.GetBool())
                    return;

                _initCompleted = false;
                _polledVariablesCount = _mainInstance.PolledVariablesCount;

                if(!_instanzenMutex.WaitOne(10000))
                {
                    Logging.WriteLog("Instanzenmutex-Deadlock in init_ValueChanged.");
                }
                if(!_homegearDevicesMutex.WaitOne(10000))
                {
                    Logging.WriteLog("Devicesmutex-Deadlock in init_ValueChanged.");
                }

                UInt16 x = 0;
                _mainInstance.Status = "Init";
                Logging.WriteLog("Init Devices");

                _deviceID = sender.Instance.Get("DeviceID");
                _deviceInstance = sender.Instance.Get("DeviceInstance");
                _deviceRemark = sender.Instance.Get("DeviceRemark");
                _deviceTypeString = sender.Instance.Get("DeviceTypeString");
                _deviceState = sender.Instance.Get("DeviceState");
                _deviceStateColor = sender.Instance.Get("DeviceStateColor");

                _mainInstance.Get("HomeGearVersion").Set(_homegear.Version);


                // Alle HomeGear Instanzen auslesen
                List<String> homegearClasses = getHomeGearClasses();
                
                //Alle instanz-handles wieder aufheben 
                /*if (_instanzen != null)
                {
                    foreach (KeyValuePair<Int32, AXInstance> disposeInstance in _instanzen)
                    {
                        Logging.WriteLog("Moin " + disposeInstance.Value.Name);
                        disposeInstance.Value.Dispose();
                    }
                }*/
                //Nicht disposen -> HomegearLib  merkt, wenn sich spsid gechanged hat.
                // lieber gucken, ob neue instanzen hinzugekommen sind oder alte gelöscht wurden und in _instanzen hinzufügen oder löschen#!!!!!!!!
                Dictionary<Int32, AXInstance> tempinstanzen = new Dictionary<Int32, AXInstance>();
                Dictionary<Int32, String> classnames = new Dictionary<Int32, String>();
                foreach (String name in homegearClasses)
                {
                    List<String> instanceNames = _aX.GetInstanceNames(name);
                    foreach (String name2 in instanceNames)
                    {
                        AXInstance instanz = new AXInstance(_aX, name2, "Status", "err");
                        if (instanz.Get("ID").GetLongInteger() >= 0)  //Wenn eine Instanz frisch im aX instanziert wurde und keine ID vergeben wurde, ist die ID -1
                        {
                            //Logging.WriteLog("Adding Instance " + instanz.Name + " (ID: " + instanz.Get("ID").GetLongInteger().ToString() + ")");
                            tempinstanzen.Add(instanz.Get("ID").GetLongInteger(), instanz);
                            classnames.Add(instanz.Get("ID").GetLongInteger(), GetClassname(instanz.Name));
                        }
                    }
                }

                //Gucken ob noch in _instanzen irgendwelche Instanzen stehen die es in tempinstanzen schon nicht mehr gibt.
                /*if (_instanzen != null)
                {
                    foreach (KeyValuePair<Int32, AXInstance> disposeInstance in _instanzen)
                    {
                        Logging.WriteLog("Checke " + disposeInstance.Value.Name);
                        if (!tempinstanzen.ContainsValue(disposeInstance.Value))
                        {
                            Logging.WriteLog("Dispose " + disposeInstance.Value.Name);
                            disposeInstance.Value.Dispose();
                        }
                        else
                            Logging.WriteLog(disposeInstance.Value.Name + " ist noch vorhanden");
                    }
                }*/
                _instanzen = new Dictionary<Int32, AXInstance>();
                x = 0;
                foreach (KeyValuePair<Int32, Device> devicePair in _homegear.Devices)
                {
                    _deviceID.Set(x, devicePair.Key);
                    if (tempinstanzen.ContainsKey(devicePair.Key))
                    {
                        if (classnames[devicePair.Key] == devicePair.Value.TypeString)
                        {
                            AXInstance aktInstanz = tempinstanzen[devicePair.Key];
                            _instanzen[devicePair.Key] = tempinstanzen[devicePair.Key];
                            _deviceTypeString.Set(x, devicePair.Value.TypeString);
                            _deviceInstance.Set(x, aktInstanz.Name);
                            if (aktInstanz.Remark.Trim() != "")
                                _deviceRemark.Set(x, aktInstanz.Remark);
                            else if (devicePair.Value.Name.Trim() != "")
                                _deviceRemark.Set(x, "HomeGear-Name: " + devicePair.Value.Name);
                            else
                                _deviceRemark.Set(x, "");
                            _deviceState.Set(x, "OK");
                            _deviceStateColor.Set(x, (Int16)DeviceStatus.OK);

                            aktInstanz.Get("SerialNo").Set(devicePair.Value.SerialNumber);
                            if (aktInstanz.VariableExists("Name"))
                                aktInstanz.Get("Name").Set(devicePair.Value.Name);

                            aktInstanz.Get("Lifetick").Set(true);

                            aktInstanz.SetVariableEvents(true);
                            aktInstanz.PollingInterval = 20;
                            aktInstanz.VariableValueChanged += aktInstanz_VariableValueChanged;
                            _polledVariablesCount += aktInstanz.PolledVariablesCount;
                            //Alle Sub-Instanzen auslesen und ebenfalls VariableChanged-Events drauf los lassen
                            //Logging.WriteLog("Adding Instance-Event for " + aktInstanz.Path + " (" + aktInstanz.Subinstances.Length + " Subinstances)");
                            foreach (AXInstance aktSubinstance in aktInstanz.Subinstances)
                            {
                                aktSubinstance.SetVariableEvents(true);
                                aktSubinstance.PollingInterval = 20;
                                aktSubinstance.VariableValueChanged += Subinstance_VariableValueChanged;
                                //Logging.WriteLog("Adding Subinstance-Event for " + aktSubinstance.Path);
                                _polledVariablesCount += aktSubinstance.PolledVariablesCount;
                            }

                            //Aktuelle Config- und Statuswerte Werte auslesen
                            Device aktDevice = devicePair.Value;
                            foreach (KeyValuePair<Int32, Channel> aktChannel in aktDevice.Channels)
                            {
                                foreach (KeyValuePair<String, Variable> Wert in aktDevice.Channels[aktChannel.Key].Variables)
                                {
                                    var aktVar = Wert.Value;
                                    String aktVarName = aktVar.Name + "_V" + aktChannel.Key.ToString("D2");
                                    if (aktInstanz.VariableExists(aktVarName))
                                    {
                                        AXVariable aktAXVar = aktInstanz.Get(aktVarName);
                                        if (aktAXVar != null) _varConverter.SetAXVariable(aktAXVar, aktVar);
                                    }
                                    String subinstance = "V" + aktChannel.Key.ToString("D2");
                                    if (aktInstanz.SubinstanceExists(subinstance))
                                    {
                                        AXVariable aktAXVar2 = aktInstanz.GetSubinstance(subinstance).Get(aktVar.Name);
                                        if (aktAXVar2 != null) _varConverter.SetAXVariable(aktAXVar2, aktVar);
                                    }
                                }

                                foreach (KeyValuePair<String, ConfigParameter> configName in aktDevice.Channels[aktChannel.Key].Config)
                                {
                                    var aktVar = configName.Value;
                                    String aktVarName = aktVar.Name + "_C" + aktChannel.Key.ToString("D2");
                                    if (aktInstanz.VariableExists(aktVarName))
                                    {
                                        //Console.WriteLine("Setze " + aktInstanz.Name + "." + aktVarName);
                                        AXVariable aktAXVar = aktInstanz.Get(aktVarName);
                                        if (aktAXVar != null) _varConverter.SetAXVariable(aktAXVar, aktVar);
                                    }
                                    String subinstance = "C" + aktChannel.Key.ToString("D2");
                                    if (aktInstanz.SubinstanceExists(subinstance))
                                    {
                                        AXVariable aktAXVar2 = aktInstanz.GetSubinstance(subinstance).Get(aktVar.Name);
                                        if (aktAXVar2 != null) _varConverter.SetAXVariable(aktAXVar2, aktVar);
                                    }
                                }
                            }
                            aktInstanz.Get("ConfigValuesChanged").Set(false);
                        }
                        else  //if (classnames[devicePair.Key] == devicePair.Value.TypeString)
                        {
                            _deviceTypeString.Set(x, devicePair.Value.TypeString);
                            _deviceInstance.Set(x, "");
                            _deviceRemark.Set(x, "");
                            _deviceState.Set(x, "Falsche Klasse! (" + classnames[devicePair.Key] + ")");
                            _deviceStateColor.Set(x, (Int16)DeviceStatus.Fehler);
                        }
                    }
                    else
                    {
                        _deviceTypeString.Set(x, devicePair.Value.TypeString);
                        _deviceInstance.Set(x, "");
                        _deviceRemark.Set(x, "");
                        _deviceState.Set(x, "Keine Instanz gefunden");
                        _deviceStateColor.Set(x, (Int16)DeviceStatus.KeineInstanzVorhanden);
                    }
                    x++;
                }

                for (; x < _deviceID.Length; x++)
                {
                    _deviceID.Set(x, 0);
                    _deviceTypeString.Set(x, "");
                    _deviceInstance.Set(x, "");
                    _deviceRemark.Set(x, "");
                    _deviceState.Set(x, "");
                    _deviceStateColor.Set(x, (Int16)DeviceStatus.Nichts);

                }
                _initCompleted = true;
                Logging.WriteLog("Init Devices completed");
                _homegearDevicesMutex.ReleaseMutex();
                _instanzenMutex.ReleaseMutex();
            }
            catch (Exception ex)
            {
                _homegearDevicesMutex.ReleaseMutex();
                _instanzenMutex.ReleaseMutex();
                sender.Instance.Error = ex.Message;
                Logging.WriteLog(ex.Message, ex.StackTrace);
            }
            finally
            {
                try
                {
                    sender.Set(false);
                    _mainInstance.Get("PolledVariables").Set(_polledVariablesCount);
                }
                catch (Exception ex)
                {
                    Logging.WriteLog(ex.Message, ex.StackTrace);
                }
            }
        }

        void Subinstance_VariableValueChanged(AXVariable sender)
        {
            Logging.WriteLog("Variable " + sender.Path + " has changed to: " + _varConverter.AutomationXVarToString(sender));
            string[] teile = sender.Path.Split('.');
            if (teile.Length == 0)
                return;

            String parentInstanceName = teile.First();
            
            _homegearDevicesMutex.WaitOne();
            try
            {
                AXInstance parentInstance = new AXInstance(_aX, parentInstanceName, "Status", "err");
                if (_homegear.Devices.ContainsKey(parentInstance.Get("ID").GetLongInteger()))
                {
                    Device aktDevice = _homegear.Devices[parentInstance.Get("ID").GetLongInteger()];
                    String name;
                    String type;
                    Int32 channelIndex;
                    name = sender.Name;
                    type = sender.Instance.Name.Substring(0, 1);
                    Int32.TryParse(sender.Instance.Name.Substring((sender.Instance.Name.Length - 2), 2), out channelIndex);


                    /////////////////////////////////////////////////////
                    // Variablen, die nicht beschrieben werden dürfen, aber in der XML-Datei als writeable gekennzeichnet sind.
                    // Beschreibt man so eine Variable, passieren komische Dinge in HomeGear
                    List<String> notWritableVars = new List<String>();
                    if (parentInstance.VariableExists("NotWritableVars"))
                    {
                        AXVariable notWritableVarsAX = parentInstance.Get("NotWritableVars");
                        for (UInt16 x = 0; x < notWritableVarsAX.Length; x++)
                            notWritableVars.Add(notWritableVarsAX.GetString(x));

                        if (notWritableVars.Contains(name))
                        {
                            //sender.Instance.Status = "VariableName '" + name + "' is in NotWritableVars";
                            //Logging.WriteLog("VariableName '" + name + "' is in NotWritableVars");
                            _homegearDevicesMutex.ReleaseMutex();
                            return;
                        }
                    }

                    if (aktDevice.Channels.ContainsKey(channelIndex))
                    {
                        Channel channel = aktDevice.Channels[channelIndex];
                        if (type == "V")
                        {
                            if (channel.Variables.ContainsKey(name))
                            {
                                Logging.WriteLog("Set Homegear Variable " + parentInstance.Name + "." + name + ", Channel:" + channelIndex.ToString() + " = " + _varConverter.AutomationXVarToString(sender));
                                _varConverter.SetHomeGearVariable(channel.Variables[name], sender);
                            }
                        }
                        else if (type == "C")
                        {
                            if (channel.Config.ContainsKey(name))
                            {
                                Logging.WriteLog("Set Homegear Config " + parentInstance.Name + "." + name + ", Channel: " + channelIndex.ToString() + " = " + _varConverter.AutomationXVarToString(sender));
                                _varConverter.SetHomeGearVariable(channel.Config[name], sender);
                            }
                        }
                    }
                }
                _homegearDevicesMutex.ReleaseMutex();
            }
            catch (Exception ex)
            {
                _homegearDevicesMutex.ReleaseMutex();
                Logging.WriteLog(ex.Message, ex.StackTrace);
            }
        }

        void aktInstanz_VariableValueChanged(AXVariable sender)
        {
            try
            {
                _homegearDevicesMutex.WaitOne();
                if(_homegear.Devices.ContainsKey(sender.Instance.Get("ID").GetLongInteger()))
                {
                    Device aktDevice = _homegear.Devices[sender.Instance.Get("ID").GetLongInteger()];
                    String name;
                    String type;
                    Int32 channelIndex;

                    if (sender.Name == "SetConfigValues")
                    {
                        if (!sender.GetBool())
                        {
                            _homegearDevicesMutex.ReleaseMutex();
                            return;
                        }

                        //Alle ConfigValues raussuchen
                        List<Int32> ChannelsPut = new List<int>();
                        ChannelsPut.Add(-1);
                        foreach(AXVariable aktVar in sender.Instance.Variables)
                        {
                            _varConverter.ParseAXVariable(aktVar, out name, out type, out channelIndex);
                            if (type == "C")
                            {
                                if (ChannelsPut.IndexOf(channelIndex) == -1)
                                {
                                    aktDevice.Channels[channelIndex].Config.Put();
                                    ChannelsPut.Add(channelIndex);
                                    Console.WriteLine("Pushe Config für Kanal " + channelIndex.ToString());
                                    sender.Instance.Status = "Pushe Config für Kanal " + channelIndex.ToString();
                                }
                            }
                        }
                    }
                    else if (sender.Name == "Name")
                    {
                        aktDevice.Name = sender.Instance.Get("Name").GetString();
                    }
                }
                _homegearDevicesMutex.ReleaseMutex();
            }
            catch (Exception ex)
            {
                _homegearDevicesMutex.ReleaseMutex();
                _mainInstance.Error = ex.Message;
                Logging.WriteLog(ex.Message, ex.StackTrace);
            }
            finally
            {
                try
                {
                    if (sender.Name == "SetConfigValues")
                        sender.Set(false);
                }
                catch (Exception ex)
                {
                    Logging.WriteLog(ex.Message, ex.StackTrace);
                }
            }
        }

        void aX_ShuttingDown(AX sender)
        {
            Dispose();
        }


        void _aX_SpsIdChanged(AX sender)
        {
            try
            {
                if (!_initCompleted) return;
                _initCompleted = false;

                Logging.WriteLog("SPS-ID Changed! Triggering Init!");
                AXVariable init = _mainInstance.Get("Init");
                init.Set(true);
                init_ValueChanged(init);
            }
            catch (Exception ex)
            {
                Logging.WriteLog(ex.Message, ex.StackTrace);
            }
        }


        void Dispose()
        {
            try
            {
                if (_disposing) return;
                _disposing = true;

                Console.WriteLine("Aus, Ende!");

                Console.WriteLine("Beende RPC-Server...");
                _mainInstance.Status = "Beende RPC-Server...";
                _homegear.Dispose();

                _mainInstance.Get("err").Set(false);
                _mainInstance.Get("Init").Set(false);
                _mainInstance.Get("RPC_InitComplete").Set(false);
                _mainInstance.Get("ServiceMessageVorhanden").Set(false);
                _mainInstance.Get("PairingMode").Set(false);
                _mainInstance.Get("CAPI_Running").Set(false);
                _mainInstance.Status = "StaKoTecHomeGear.exe beendet";
                Console.WriteLine("Und aus!!");
            }
            catch (Exception ex)
            {
                _mainInstance.Error = ex.Message;
                Logging.WriteLog(ex.Message, ex.StackTrace);
            }
            finally
            {
                Environment.Exit(0);
            }
        }


        private void HomeGearConnect()
        {
            try
            {
                if (_homegear != null) return;
                Int32 homegearPort = _mainInstance.Get("HomeGearPort").GetInteger();
                Int32 listenPort = _mainInstance.Get("ListenPort").GetInteger();
                String homegearHostName = _mainInstance.Get("HomeGearHostName").GetString();
                String aXHostName = _mainInstance.Get("aXHostName").GetString();
                String aXHostListenIP = _mainInstance.Get("aXHostListenIP").GetString();
                Boolean sslEnable = _mainInstance.Get("SSL_Enable").GetBool();
                String sslHomeGearUsername = _mainInstance.Get("SSL_HomeGear_Username").GetString();
                String sslHomeGearPassword = _mainInstance.Get("SSL_HomeGear_Password").GetString();
                String sslaXUsername = _mainInstance.Get("SSL_aX_Username").GetString();
                String sslaXPassword = _mainInstance.Get("SSL_aX_Password").GetString();
                String sslCertificatePath = _mainInstance.Get("SSL_CertificatePath").GetString();
                String sslCertificatePassword = _mainInstance.Get("SSL_CertificatePassword").GetString();
                Boolean sslVerifyCertificate = _mainInstance.Get("SSL_VerifyCertificate").GetBool();

                SSLClientInfo sslClientInfo = null;
                SSLServerInfo sslServerInfo = null;
                if (sslEnable)
                {
                    sslClientInfo = new SSLClientInfo(aXHostName, sslHomeGearUsername, sslHomeGearPassword, sslVerifyCertificate);
                    sslServerInfo = new SSLServerInfo(sslCertificatePath, sslCertificatePassword, sslaXUsername, sslaXPassword);
                }
                _rpc = new RPCController(homegearHostName, homegearPort, aXHostName, aXHostListenIP, listenPort, sslClientInfo, sslServerInfo);
                _rpc.ClientConnected += _rpc_ClientConnected;
                _rpc.ClientDisconnected += _rpc_ClientDisconnected;
                _rpc.ServerConnected += _rpc_ServerConnected;
                _rpc.ServerDisconnected += _rpc_ServerDisconnected;

                _homegear = new Homegear(_rpc);
                _homegear.ConnectError += _homegear_ConnectError;
                _homegear.SystemVariableUpdated += _homegear_SystemVariableUpdated;
                _homegear.MetadataUpdated += _homegear_MetadataUpdated;
                _homegear.DeviceVariableUpdated += _homegear_DeviceVariableUpdated;
                _homegear.DeviceConfigParameterUpdated += _homegear_DeviceConfigParameterUpdated;
                _homegear.DeviceLinkConfigParameterUpdated += _homegear_DeviceLinkConfigParameterUpdated;
                _homegear.ReloadRequired += _homegear_ReloadRequired;
                _homegear.DeviceReloadRequired += _homegear_DeviceReloadRequired;
                _homegear.Reloaded += _homegear_Reloaded;
                _homegear.EventUpdated += _homegear_EventUpdated;
            }
            catch (Exception ex)
            {
                _mainInstance.Error = ex.Message;
                Logging.WriteLog(ex.Message, ex.StackTrace);
            }
        }

        void _homegear_EventUpdated(Homegear sender, Event homegearEvent)
        {
            Logging.WriteLog("HomeGear Event " + homegearEvent.ToString() + " is updated");
        }

        void _homegear_Reloaded(Homegear sender)
        {
            try
            {
                _mainInstance.Status = "RPC: Reload feddich";
                _mainInstance.Get("RPC_InitComplete").Set(true);
                Logging.WriteLog("RPC Init Complete");
            }
            catch (Exception ex)
            {
                _mainInstance.Error = ex.Message;
                Logging.WriteLog(ex.Message, ex.StackTrace);
            }
        }

        void _homegear_DeviceReloadRequired(Homegear sender, Device device, Channel channel, DeviceReloadType reloadType)
        {
            try
            {
                _mainInstance.Status = "RPC: Reload erforderlich (" + reloadType.ToString() + ")";
                _homegearDevicesMutex.WaitOne();
                if (reloadType == DeviceReloadType.Full)
                {
                    Logging.WriteLog("Reloading device " + device.ID.ToString() + ".");
                    //Finish all operations on the device and then call:
                    device.Reload();
                }
                else if (reloadType == DeviceReloadType.Metadata)
                {
                    Logging.WriteLog("Reloading metadata of device " + device.ID.ToString() + ".");
                    //Finish all operations on the device's metadata and then call:
                    device.Metadata.Reload();
                }
                else if (reloadType == DeviceReloadType.Channel)
                {
                    Logging.WriteLog("Reloading channel " + channel.Index + " of device " + device.ID.ToString() + ".");
                    //Finish all operations on the device's channel and then call:
                    channel.Reload();
                }
                else if (reloadType == DeviceReloadType.Links)
                {
                    Logging.WriteLog("Device links were updated: Device type: \"" + device.TypeString + "\", ID: " + device.ID.ToString() + ", Channel: " + channel.Index.ToString());
                    Logging.WriteLog("Reloading links of channel " + channel.Index + " and device " + device.ID.ToString() + ".");
                    //Finish all operations on the channels's links and then call:
                    channel.Links.Reload();
                }
                else if (reloadType == DeviceReloadType.Team)
                {
                    Logging.WriteLog("Device team was updated: Device type: \"" + device.TypeString + "\", ID: " + device.ID.ToString() + ", Channel: " + channel.Index.ToString());
                    Logging.WriteLog("Reloading channel " + channel.Index + " of device " + device.ID.ToString() + ".");
                    //Finish all operations on the device's channel and then call:
                    channel.Reload();
                }
                else if (reloadType == DeviceReloadType.Events)
                {
                    Logging.WriteLog("Device events were updated: Device type: \"" + device.TypeString + "\", ID: " + device.ID.ToString() + ", Channel: " + channel.Index.ToString());
                    Logging.WriteLog("Reloading events of device " + device.ID.ToString() + ".");
                    //Finish all operations on the device's events and then call:
                    device.Events.Reload();
                }
                _homegearDevicesMutex.ReleaseMutex();
            }
            catch (Exception ex)
            {
                _homegearDevicesMutex.ReleaseMutex();
                _mainInstance.Error = ex.Message;
                Logging.WriteLog(ex.Message, ex.StackTrace);
            }
        }

        void _homegear_ReloadRequired(Homegear sender, ReloadType reloadType)
        {
            try
            {
                _mainInstance.Status = "RPC: Reload erforderlich (" + reloadType.ToString() + ")";
                if (reloadType == ReloadType.Full)
                {
                    try
                    {
                        _mainInstance.Get("RPC_InitComplete").Set(false);
                        Logging.WriteLog("Homegear is full-reloading");
                        _homegearDevicesMutex.WaitOne();
                        _homegear.Reload();
                        _homegearDevicesMutex.ReleaseMutex();
                    }
                    catch (Exception ex)
                    {
                        _mainInstance.Error = "Reload Thread ist tot";
                        Logging.WriteLog(ex.Message, ex.StackTrace);
                    }
                }
                else if (reloadType == ReloadType.SystemVariables)
                {
                    Logging.WriteLog("Homegear is reloading SystemVariables");
                    _homegear.SystemVariables.Reload();
                }
                else if (reloadType == ReloadType.Events)
                {
                    Logging.WriteLog("Homegear is reloading Events");
                    _homegear.TimedEvents.Reload();
                }
            }
            catch (Exception ex)
            {
                _mainInstance.Error = "Reload Thread ist tot";
                Logging.WriteLog(ex.Message, ex.StackTrace);
            }
        }

        void _homegear_DeviceLinkConfigParameterUpdated(Homegear sender, Device device, Channel channel, Link link, ConfigParameter parameter)
        {
            try
            {
                _instanzenMutex.WaitOne();
                if (_instanzen.ContainsKey(device.ID))
                    _instanzen[device.ID].Status = "Link-Parameter " + link.Name + " updated to " + link.ToString();
                _mainInstance.Status = "RPC: " + device.ID.ToString() + " " + link.RemotePeerID.ToString() + " " + link.RemoteChannel.ToString() + " " + parameter.Name + " = " + parameter.ToString();
                _instanzenMutex.ReleaseMutex();
            }
            catch (Exception ex)
            {
                _instanzenMutex.ReleaseMutex();
                _mainInstance.Error = ex.Message;
                Logging.WriteLog(ex.Message, ex.StackTrace);
            }
        }

        void _homegear_DeviceConfigParameterUpdated(Homegear sender, Device device, Channel channel, ConfigParameter parameter)
        {
            try
            {
                _instanzenMutex.WaitOne();
                if (_instanzen.ContainsKey(device.ID))
                    _instanzen[device.ID].Status = "Config-Parameter " + parameter.Name + " updated to " + parameter.ToString();
                _mainInstance.Status = "RPC: " + device.ID.ToString() + " " + parameter.Name + " = " + parameter.ToString();
                _instanzenMutex.ReleaseMutex();
            }
            catch (Exception ex)
            {
                _instanzenMutex.ReleaseMutex();
                _mainInstance.Error = ex.Message;
                Logging.WriteLog(ex.Message, ex.StackTrace);
            }
        }

        void _homegear_DeviceVariableUpdated(Homegear sender, Device device, Channel channel, Variable variable)
        {
            Int32 deviceID = 0;
            try
            {
                _homegearDevicesMutex.WaitOne();
                _instanzenMutex.WaitOne();
                deviceID = device.ID;
                _homegearDevicesMutex.ReleaseMutex();
            }
            catch (Exception ex)
            {
                _homegearDevicesMutex.ReleaseMutex();
                _mainInstance.Error = ex.Message;
                Logging.WriteLog(ex.Message, ex.StackTrace);
            }

            try
            {
                _mainInstance.Status = "RPC: " + deviceID.ToString() + " " + variable.Name + " = " + variable.ToString();

                if (_instanzen.ContainsKey(deviceID))
                {
                    AXInstance instanz = _instanzen[deviceID];
                    String varName = variable.Name + "_V" + channel.Index.ToString("D2");
                    if (instanz.VariableExists(varName))
                    {
                        AXVariable aktAXVar = instanz.Get(varName);
                        if (aktAXVar != null) _varConverter.SetAXVariable(aktAXVar, variable);
                    }
                    String subinstance = "V" + channel.Index.ToString("D2");
                    if (instanz.SubinstanceExists(subinstance))
                    {
                        AXVariable aktAXVar2 = instanz.GetSubinstance(subinstance).Get(variable.Name);
                        if (aktAXVar2 != null) _varConverter.SetAXVariable(aktAXVar2, variable);
                    }

                    AXVariable aXVariable_LastChange = instanz.Get("LastChange");
                    List<String> lastChange = new List<String>();
                    UInt16 x = 0;

                    lastChange.Add(DateTime.Now.Hour.ToString("D2") + ":" + DateTime.Now.Minute.ToString("D2") + ":" + DateTime.Now.Second.ToString("D2") + ": " + varName + " = " + variable.ToString());
                    for (x = 0; x < aXVariable_LastChange.Length; x++)
                        lastChange.Add(aXVariable_LastChange.GetString(x));
                    for (x = 0; x < aXVariable_LastChange.Length; x++)
                        aXVariable_LastChange.Set(x, lastChange[x]);


                    instanz.Status = varName + " = " + variable.ToString();
                    //Console.WriteLine(device.SerialNumber + ": " + variable.Name + ": " + variable.ToString());
                }
                _instanzenMutex.ReleaseMutex();
            }
            catch (Exception ex)
            {
                _instanzenMutex.ReleaseMutex();
                _mainInstance.Error = ex.Message;
                Logging.WriteLog(ex.Message, ex.StackTrace);
            }
        }


        void _homegear_MetadataUpdated(Homegear sender, Device device, MetadataVariable variable)
        {
            try
            {
                _mainInstance.Status = "RPC: Metadata Variable '" + variable.Name + "' geändert";
            }
            catch (Exception ex)
            {
                _mainInstance.Error = ex.Message;
                Logging.WriteLog(ex.Message, ex.StackTrace);
            }
        }

        void _homegear_SystemVariableUpdated(Homegear sender, SystemVariable variable)
        {
            try
            {
                _mainInstance.Status = "RPC: System Variable '" + variable.Name + "' geändert";
            }
            catch (Exception ex)
            {
                _mainInstance.Error = ex.Message;
                Logging.WriteLog(ex.Message, ex.StackTrace);
            }
        }

        void _homegear_ConnectError(Homegear sender, string message, string stackTrace)
        {
            try
            {
                _mainInstance.Error = message;
                Logging.WriteLog(message);
            }
            catch (Exception ex)
            {
                _mainInstance.Error = ex.Message;
                Logging.WriteLog(ex.Message, ex.StackTrace);
            }
        }

        void _rpc_ClientDisconnected(RPCClient sender)
        {
            try
            {
                Logging.WriteLog("RPC-Client Verbindung unterbrochen");
            }
            catch (Exception ex)
            {
                _mainInstance.Error = ex.Message;
                Logging.WriteLog(ex.Message, ex.StackTrace);
            }
        }

        void _rpc_ClientConnected(RPCClient sender, CipherAlgorithmType cipherAlgorithm, Int32 cipherStrength)
        {
            try
            {
                Logging.WriteLog("RPC-Client verbunden");
                _mainInstance.Get("err").Set(false);
            }
            catch (Exception ex)
            {
                _mainInstance.Error = ex.Message;
                Logging.WriteLog(ex.Message, ex.StackTrace);
            }
        }

        void _rpc_ServerDisconnected(RPCServer sender)
        {
            try
            {
                Logging.WriteLog("Verbindung von Homegear zu aX unterbrochen");
            }
            catch (Exception ex)
            {
                _mainInstance.Error = ex.Message;
                Logging.WriteLog(ex.Message, ex.StackTrace);
            }
        }

        void _rpc_ServerConnected(RPCServer sender, CipherAlgorithmType cipherAlgorithm, Int32 cipherStrength)
        {
            try
            {
                Logging.WriteLog("Eingehende Verbindung von Homegear hergestellt");
                _mainInstance.Get("err").Set(false);
            }
            catch (Exception ex)
            {
                _mainInstance.Error = ex.Message;
                Logging.WriteLog(ex.Message, ex.StackTrace);
            }
        }
    }
}
