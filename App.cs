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

    class App : IDisposable
    {
        //Globale Variablen
        bool _disposing = false;
        bool _initCompleted = false;
        AX _aX = null;
        AXInstance _mainInstance = null;

        VariableConverter _varConverter = null;

        AXVariable _deviceID = null;
        AXVariable _deviceInstance = null;
        AXVariable _deviceRemark = null;
        AXVariable _deviceTypeString = null;
        AXVariable _deviceState = null;
        AXVariable _deviceStateColor = null;
        HomegearLib.RPC.RPCController _rpc = null;
        HomegearLib.Homegear _homegear = null;

        //Instances _instanzen = null;
        Instances _instances = null;
        Mutex _homegearDevicesMutex = new Mutex();
        //Queue<AXInstance> _instancesToDispose = new Queue<AXInstance>();

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
                _instances = new Instances(_aX, _mainInstance);
                _instances.VariableValueChanged += OnInstanceVariableValueChanged;
                _instances.SubinstanceVariableValueChanged += OnSubinstanceVariableValueChanged;

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
                Int32 connectionTimeout = 0;

                AXVariable mainInstanceErr = _mainInstance.Get("err");

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
                            _mainInstance.Get("RPC_InitComplete").Set(false);
                            _initCompleted = false;

                            if (connectionTimeout > 0)
                                Logging.WriteLog("Waiting for RPC-Server connection... (" + (connectionTimeout * 5).ToString() + " s)");

                            if ((connectionTimeout > 6) && !mainInstanceErr.GetBool())
                                mainInstanceErr.Set(true);


                            Thread.Sleep(5000);
                            connectionTimeout++;
                            continue;
                        }
                        else
                            connectionTimeout = 0;

                        

                        //Allen Devices einen Lifetick senden um DataValid zu generieren
                        if (_initCompleted)
                            _instances.Lifetick();


                        if (_homegear != null && _initCompleted && j % 10 == 0)
                        {
                            Boolean serviceMessageVorhanden = false;
                            List<ServiceMessage> serviceMessages = _homegear.ServiceMessages;
                            UInt16 x = 0;
                            Int16 alarmCounter = 0;
                            Int16 warningCounter = 0;
                            AXVariable aX_serviceMessages = _mainInstance.Get("ServiceMessages");
                            foreach (ServiceMessage message in serviceMessages)
                            {
                                if ((message.Type == "UNREACH") || (message.Type == "STICKY_UNREACH") || (message.Type == "LOWBAT") || (message.Type == "CENTRAL_ADDRESS_SPOOFED"))
                                    alarmCounter++;
                                else
                                    warningCounter++;

                                String aktMessage = "Device ID: " + message.PeerID.ToString();
                                if (_homegear.Devices.ContainsKey(message.PeerID))
                                {
                                    if (_homegear.Devices[message.PeerID].Name.Length > 0)
                                        aktMessage += " " + _homegear.Devices[message.PeerID].Name;

                                    aktMessage += " (" + _homegear.Devices[message.PeerID].TypeString + ")";
                                }
                                aktMessage += " " + "Channel: " + message.Channel.ToString() + " " + message.Type + " = " + message.Value.ToString();
                                aX_serviceMessages.Set(x, aktMessage);
                                x++;
                                serviceMessageVorhanden = true;
                            }
                            for (; x < aX_serviceMessages.Length; x++)
                                aX_serviceMessages.Set(x, "");

                            _mainInstance.Get("ServiceMessageVorhanden").Set(serviceMessageVorhanden);
                            _mainInstance.Get("AlarmCounter").Set(alarmCounter);
                            _mainInstance.Get("WarningCounter").Set(warningCounter);
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

        String findVarInClass(Int32 id, String varName, VariableType varType,  String Typ, String Kanal)
        {
            try
            {
                bool instanceVarExists = false;
                bool subinstanceVarExists = false;
                AXVariableType type = AXVariableType.axUnsignedShortInteger;

                if (!_instances.ContainsKey(id))
                    return "Keine Instanz vorhanden";

                instanceVarExists = (_instances[id].VariableExists(varName + "_" + Typ + Kanal));
                subinstanceVarExists = _instances[id].SubinstanceExists(Typ + Kanal) && _instances[id].GetSubinstance(Typ + Kanal).VariableExists(varName);
            
                if (!instanceVarExists && ! subinstanceVarExists)
                    return "Variable nicht vorhanden";

                if (instanceVarExists && subinstanceVarExists)
                    return "Variable in instanz UND SubInstanz vorhanden!";

                ////// Variablentyp prüfen
                if (instanceVarExists)
                    type = _instances[id].Get(varName + "_" + Typ + Kanal).Type;
                if (subinstanceVarExists)
                    type = _instances[id].GetSubinstance(Typ + Kanal).Get(varName).Type;

                if (((varType == VariableType.tInteger) && (type != AXVariableType.axLongInteger)) ||
                    ((varType == VariableType.tDouble) && (type != AXVariableType.axLongReal)) ||
                    ((varType == VariableType.tBoolean) && (type != AXVariableType.axBool)) ||
                    ((varType == VariableType.tAction) && (type != AXVariableType.axBool)) ||
                    ((varType == VariableType.tEnum) && (type != AXVariableType.axLongInteger)) ||
                    ((varType == VariableType.tString) && (type != AXVariableType.axString)))
                    return "Falscher Variablen-Typ";


                return "OK";
            }
            catch (Exception ex)
            {
                Logging.WriteLog(ex.Message, ex.StackTrace);
                return ex.Message;
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
                AXVariable deviceVars_Unit = _mainInstance.Get("DeviceVars_Dimension");
                AXVariable deviceVars_VarVorhanden = _mainInstance.Get("DeviceVars_VarVorhanden");
                UInt16 x = 0;
                Int32 deviceID = _mainInstance.Get("ActionID").GetLongInteger();
       
                _homegearDevicesMutex.WaitOne();
                _instances.MutexLocked = true;
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
                                    _instances.MutexLocked = false;
                                    return;
                                }

                                var aktVar = configName.Value;
                                String minVar = "";
                                String maxVar = "";
                                String typ = "";
                                String defaultVar = "";
                                String rwVar = "";
                                String varOK = findVarInClass(deviceID, aktVar.Name, aktVar.Type, "C", aktChannel.Key.ToString("D2"));

                                _varConverter.ParseDeviceConfigVars(aktVar, out minVar, out maxVar, out typ, out defaultVar, out rwVar);

                                deviceVars_Name.Set(x, aktVar.Name + "_C" + aktChannel.Key.ToString("D2"));
                                deviceVars_Type.Set(x, typ);
                                deviceVars_Min.Set(x, minVar);
                                deviceVars_Max.Set(x, maxVar);
                                deviceVars_Default.Set(x, defaultVar);
                                deviceVars_RW.Set(x, rwVar);
                                deviceVars_Unit.Set(x, aktVar.Unit);
                                deviceVars_VarVorhanden.Set(x, varOK);
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
                                    _instances.MutexLocked = false;
                                    return;
                                }

                                var aktVar = varName.Value;
                                String minVar = "";
                                String maxVar = "";
                                String typ = "";
                                String defaultVar = "";
                                String rwVar = "";
                                String varOK = findVarInClass(deviceID, aktVar.Name, aktVar.Type, "V", aktChannel.Key.ToString("D2"));
                                _varConverter.ParseDeviceVars(aktVar, out minVar, out maxVar, out typ, out defaultVar, out rwVar);

                                deviceVars_Name.Set(x, aktVar.Name + "_V" + aktChannel.Key.ToString("D2"));
                                deviceVars_Type.Set(x, typ);
                                deviceVars_Min.Set(x, minVar);
                                deviceVars_Max.Set(x, maxVar);
                                deviceVars_Default.Set(x, defaultVar);
                                deviceVars_RW.Set(x, rwVar);
                                deviceVars_Unit.Set(x, aktVar.Unit);
                                deviceVars_VarVorhanden.Set(x, varOK);
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
                        deviceVars_Unit.Set(x, "");
                        deviceVars_VarVorhanden.Set(x, "");
                    }
                }
                _homegearDevicesMutex.ReleaseMutex();
                _instances.MutexLocked = false;
            }
            catch (Exception ex)
            {
                try { _homegearDevicesMutex.ReleaseMutex(); }
                catch (Exception) { }
                _instances.MutexLocked = false;
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
                try{ _homegearDevicesMutex.ReleaseMutex(); } catch (Exception) { }
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
                try { _homegearDevicesMutex.ReleaseMutex(); } catch (Exception) { }
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
                try { _homegearDevicesMutex.ReleaseMutex(); } catch (Exception) { }
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
                Logging.WriteLog(ex.Message, ex.StackTrace);
            }
            return className;
        }

        void init_ValueChanged(AXVariable sender)
        {
            try
            {
                if (!sender.GetBool())
                    return;
                
                _initCompleted = false;

                _homegearDevicesMutex.WaitOne();

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

                _instances.Reload(_homegear.Devices);
                _instances.MutexLocked = true;
                try
                {
                    x = 0;
                    foreach (KeyValuePair<Int32, Device> devicePair in _homegear.Devices)
                    {
                        _deviceID.Set(x, devicePair.Key);
                        if (_instances.ContainsKey(devicePair.Key))
                        {
                            AXInstance aktInstanz = _instances[devicePair.Key];
                            if (aktInstanz.ClassName == devicePair.Value.TypeString)
                            {
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
                                            if (aktAXVar != null)
                                            {
                                                _varConverter.SetAXVariable(aktAXVar, aktVar);
                                                setDeviceStatusInMaininstance(aktVar, devicePair.Key);
                                            }
                                        }
                                        String subinstance = "V" + aktChannel.Key.ToString("D2");
                                        if (aktInstanz.SubinstanceExists(subinstance))
                                        {
                                            AXVariable aktAXVar2 = aktInstanz.GetSubinstance(subinstance).Get(aktVar.Name);
                                            if (aktAXVar2 != null)
                                            {
                                                _varConverter.SetAXVariable(aktAXVar2, aktVar);
                                                setDeviceStatusInMaininstance(aktVar, devicePair.Key);
                                            }
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
                                _deviceState.Set(x, "Falsche Klasse! (" + aktInstanz.ClassName + ")");
                                _deviceStateColor.Set(x, (Int16)DeviceStatus.Fehler);
                                _instances.Remove(devicePair.Key, false);
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
                }
                catch (Exception ex)
                {
                    Logging.WriteLog(ex.Message, ex.StackTrace);
                }
                _instances.MutexLocked = false;

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
            }
            catch (Exception ex)
            {
                try { _homegearDevicesMutex.ReleaseMutex(); }  catch (Exception) { }
                sender.Instance.Error = ex.Message;
                Logging.WriteLog(ex.Message, ex.StackTrace);
            }
            finally
            {
                try
                {
                    sender.Set(false);
                    Logging.WriteLog("Init completely finished");
                    _mainInstance.Get("PolledVariables").Set(_mainInstance.PolledVariablesCount + _instances.PolledVariablesCount);
                }
                catch (Exception ex)
                {
                    Logging.WriteLog(ex.Message, ex.StackTrace);
                }
            }
        }

        void OnSubinstanceVariableValueChanged(AXVariable sender)
        {
            _homegearDevicesMutex.WaitOne();
            if (sender == null || sender.Instance == null)
            {
                _homegearDevicesMutex.ReleaseMutex();
                return;
            }
            try
            {
                AXInstance parentInstance = sender.Instance.Parent;
                if (_homegear.Devices.ContainsKey(parentInstance.Get("ID").GetLongInteger()))
                {
                    Logging.WriteLog("Variable " + sender.Path + " has changed to: " + _varConverter.AutomationXVarToString(sender));
                    Device aktDevice = _homegear.Devices[parentInstance.Get("ID").GetLongInteger()];
                    String name;
                    String type;
                    Int32 channelIndex;
                    name = sender.Name;
                    type = sender.Instance.Name.Substring(0, 1);
                    Int32.TryParse(sender.Instance.Name.Substring((sender.Instance.Name.Length - 2), 2), out channelIndex);

                    if (aktDevice.Channels.ContainsKey(channelIndex))
                    {
                        Channel channel = aktDevice.Channels[channelIndex];
                        if (type == "V")
                        {
                            if (channel.Variables.ContainsKey(name))
                            {
                                Logging.WriteLog("[aX -> HomeGear]: " + parentInstance.Name + "." + name + ", Channel:" + channelIndex.ToString() + " = " + _varConverter.AutomationXVarToString(sender));
                                SetLastChange(sender.Instance, "[aX -> HomeGear]: " + sender.Instance.Name + "." + name + ", Channel:" + channelIndex.ToString() + " = " + _varConverter.AutomationXVarToString(sender)); 
                                _varConverter.SetHomeGearVariable(channel.Variables[name], sender);
                            }
                        }
                        else if (type == "C")
                        {
                            if (channel.Config.ContainsKey(name))
                            {
                                Logging.WriteLog("[aX -> HomeGear]: " + parentInstance.Name + "." + name + ", Channel: " + channelIndex.ToString() + " = " + _varConverter.AutomationXVarToString(sender));
                                SetLastChange(sender.Instance, "[aX -> HomeGear]: " + sender.Instance.Name + "." + name + ", Channel: " + channelIndex.ToString() + " = " + _varConverter.AutomationXVarToString(sender)); 
                                _varConverter.SetHomeGearVariable(channel.Config[name], sender);
                            }
                        }
                    }
                }
                _homegearDevicesMutex.ReleaseMutex();
            }
            catch (Exception ex)
            {
                try { _homegearDevicesMutex.ReleaseMutex(); } catch (Exception) { }
                Logging.WriteLog(ex.Message, ex.StackTrace);
            }
        }

        void OnInstanceVariableValueChanged(AXVariable sender)
        {
            if ((sender.Name == "Lifetick") || (sender.Name == "DataValid"))
                return;

            try
            {
                _homegearDevicesMutex.WaitOne();
                if (sender == null || sender.Instance == null)
                {
                    _homegearDevicesMutex.ReleaseMutex();
                    return;
                }

                if(_homegear.Devices.ContainsKey(sender.Instance.Get("ID").GetLongInteger()))
                {
                    Logging.WriteLog("Variable " + sender.Path + " has changed to: " + _varConverter.AutomationXVarToString(sender));
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

                        foreach (Int32 aktchannelIndex in aktDevice.Channels.Keys)
                        {
                            aktDevice.Channels[aktchannelIndex].Config.Put();
                            Logging.WriteLog("Pushe Config für Kanal " + aktchannelIndex.ToString());
                            sender.Instance.Status = "Pushe Config für Kanal " + aktchannelIndex.ToString();
                        }
                    }
                    else if (_varConverter.ParseAXVariable(sender, out name, out type, out channelIndex)) 
                    {
                        if (aktDevice.Channels.ContainsKey(channelIndex))
                        {
                            Channel channel = aktDevice.Channels[channelIndex];
                            if (type == "V")
                            {
                                if (channel.Variables.ContainsKey(name))
                                {
                                    Logging.WriteLog("[aX -> HomeGear]: " + sender.Instance.Name + "." + name + ", Channel:" + channelIndex.ToString() + " = " + _varConverter.AutomationXVarToString(sender));
                                    SetLastChange(sender.Instance, "[aX -> HomeGear]: " + sender.Instance.Name + "." + name + ", Channel:" + channelIndex.ToString() + " = " + _varConverter.AutomationXVarToString(sender));
                                    _varConverter.SetHomeGearVariable(channel.Variables[name], sender);
                                }
                            }
                            else if (type == "C")
                            {
                                if (channel.Config.ContainsKey(name))
                                {
                                    Logging.WriteLog("[aX -> HomeGear]: " + sender.Instance.Name + "." + name + ", Channel: " + channelIndex.ToString() + " = " + _varConverter.AutomationXVarToString(sender));
                                    SetLastChange(sender.Instance, "[aX -> HomeGear]: " + sender.Instance.Name + "." + name + ", Channel: " + channelIndex.ToString() + " = " + _varConverter.AutomationXVarToString(sender));
                                    _varConverter.SetHomeGearVariable(channel.Config[name], sender);
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
                try { _homegearDevicesMutex.ReleaseMutex(); } catch (Exception) { }
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
                System.Diagnostics.Debug.WriteLine("SPS-ID change start");
                if (!_initCompleted) return;
                _initCompleted = false;

                Logging.WriteLog("SPS-ID Changed! Triggering Init!");
                AXVariable init = _mainInstance.Get("Init");
                init.Set(true);
                init_ValueChanged(init);
                System.Diagnostics.Debug.WriteLine("SPS-ID change end");
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
                _rpc.Dispose();

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
                else if (reloadType == DeviceReloadType.Variables)
                {
                    Logging.WriteLog("Device variables were updated: Device type: \"" + device.TypeString + "\", ID: " + device.ID.ToString() + ", Channel: " + channel.Index.ToString());
                    Logging.WriteLog("Reloading variables of channel " + channel.Index + " and device " + device.ID.ToString() + ".");
                    //Finish all operations on the channels's variables and then call:
                    channel.Variables.Reload();
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
                try { _homegearDevicesMutex.ReleaseMutex(); } catch (Exception) { }
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
                _instances.MutexLocked = true;
                if (_instances.ContainsKey(device.ID))
                    _instances[device.ID].Status = "Link-Parameter " + link.Name + " updated to " + link.ToString();
                _mainInstance.Status = "RPC: " + device.ID.ToString() + " " + link.RemotePeerID.ToString() + " " + link.RemoteChannel.ToString() + " " + parameter.Name + " = " + parameter.ToString();
            }
            catch (Exception ex)
            {
                Logging.WriteLog(ex.Message, ex.StackTrace);
            }
            _instances.MutexLocked = false;
        }

        void _homegear_DeviceConfigParameterUpdated(Homegear sender, Device device, Channel channel, ConfigParameter parameter)
        {
            try
            {
                _instances.MutexLocked = true;
                if (_instances.ContainsKey(device.ID))
                    _instances[device.ID].Status = "Config-Parameter " + parameter.Name + " updated to " + parameter.ToString();
                _mainInstance.Status = "RPC: " + device.ID.ToString() + " " + parameter.Name + " = " + parameter.ToString();
            }
            catch (Exception ex)
            {
                Logging.WriteLog(ex.Message, ex.StackTrace);
            }
            _instances.MutexLocked = false;
        }

        void setDeviceStatusInMaininstance(Variable variable, Int32 id)
        {
            if ((variable.Type != VariableType.tBoolean) || ((variable.Name != "UNREACH") && (variable.Name != "STICKY_UNREACH") && (variable.Name != "LOWBAT") && (variable.Name != "CENTRAL_ADDRESS_SPOOFED")))
                return;

            for (UInt16 x = 0; x < _deviceID.Length; x++)
            {
                if (_deviceID.GetLongInteger(x) == 0)
                    return;

                if (_deviceID.GetLongInteger(x) == id)
                {
                    String stateOld = _deviceState.GetString(x);
                    if (variable.BooleanValue)
                    {
                        _deviceState.Set(x, stateOld + ", " + variable.Name);
                        _deviceStateColor.Set(x, (Int16)DeviceStatus.Fehler);
                    }
                    else
                    {
                        _deviceState.Set(x, stateOld.Replace(", " + variable.Name, ""));
                    }
                }
            }
        }

        void _homegear_DeviceVariableUpdated(Homegear sender, Device device, Channel channel, Variable variable)
        {
            _instances.MutexLocked = true;
            try
            {
                Int32 deviceID = device.ID;
                _mainInstance.Status = "RPC: " + deviceID.ToString() + " " + variable.Name + " = " + variable.ToString();
                
                if (_instances.ContainsKey(deviceID))
                {
                    AXInstance instanz = _instances[deviceID];
                    String varName = variable.Name + "_V" + channel.Index.ToString("D2");
                    if (instanz.VariableExists(varName))
                    {
                        AXVariable aktAXVar = instanz.Get(varName);
                        if (aktAXVar != null)
                        {
                            _varConverter.SetAXVariable(aktAXVar, variable);
                            setDeviceStatusInMaininstance(variable, deviceID);
                            Logging.WriteLog("[HomeGear -> aX]: " + aktAXVar.Path + " = " + variable.ToString());
                        }
                    }
                    String subinstance = "V" + channel.Index.ToString("D2");
                    if (instanz.SubinstanceExists(subinstance))
                    {
                        AXVariable aktAXVar2 = instanz.GetSubinstance(subinstance).Get(variable.Name);
                        if (aktAXVar2 != null)
                        {
                            _varConverter.SetAXVariable(aktAXVar2, variable);
                            setDeviceStatusInMaininstance(variable, deviceID);
                            Logging.WriteLog("[HomeGear -> aX]: " + aktAXVar2.Path + " = " + variable.ToString());
                        }
                    }

                    SetLastChange(instanz, varName + " = " + variable.ToString());
                    
                    //Console.WriteLine(device.SerialNumber + ": " + variable.Name + ": " + variable.ToString());
                }
            }
            catch (Exception ex)
            {
                Logging.WriteLog(ex.Message, ex.StackTrace);
            }
            _instances.MutexLocked = false;
        }

        void SetLastChange(AXInstance instanz, String text)
        {
            try
            {
                if (!instanz.VariableExists("LastChange"))
                    return;

                AXVariable aXVariable_LastChange = instanz.Get("LastChange");
                List<String> lastChange = new List<String>();
                UInt16 x = 0;

                lastChange.Add(DateTime.Now.Hour.ToString("D2") + ":" + DateTime.Now.Minute.ToString("D2") + ":" + DateTime.Now.Second.ToString("D2") + ": " + text);
                for (x = 0; x < aXVariable_LastChange.Length; x++)
                    lastChange.Add(aXVariable_LastChange.GetString(x));
                for (x = 0; x < aXVariable_LastChange.Length; x++)
                    aXVariable_LastChange.Set(x, lastChange[x]);

                instanz.Status = text;
            }
            catch (Exception ex)
            {
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
                Logging.WriteLog(ex.Message, ex.StackTrace);
            }
        }

        void IDisposable.Dispose()
        {
            Dispose();
        }
    }
}
