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
using System.Diagnostics;

namespace StaKoTecHomeGear
{
    enum DeviceStatus
    {
        Nichts = 0,
        OK = 1,
        Fehler = 2,
        KeineInstanzVorhanden = 3,
        Warnung = 4
    }

    enum LogLevel
    {
        Error = 0,
        Warning = 1,
        Info = 2,
        Debug = 3,
        Always = 100
    }



    class App : IDisposable
    {
        //Globale Variablen
        Dictionary<DeviceStatus, String> _deviceStatusText = null;
        List<String> _errorStates = null;
        List<String> _warningStates = null;

        List<Int32> _firstInitDevices = null;
        Dictionary<Device, AxInstance> _getActualDeviceDataDictionary = null;
        List<String> _ignoredVariableChanged = null;

        bool _disposing = false;
        bool _initCompleted = false;
        Ax _aX = null;
        AxInstance _mainInstance = null;
        LogLevel _logLevel = LogLevel.Error;

        VariableConverter _varConverter = null;

        AxVariable _homegearInterfaces = null;
        AxVariable _deviceID = null;
        AxVariable _deviceInstance = null;
        AxVariable _deviceRemark = null;
        AxVariable _deviceTypeString = null;
        AxVariable _deviceState = null;
        AxVariable _deviceFirmware = null;
        AxVariable _deviceInterface = null;
        AxVariable _deviceStateColor = null;

        AxVariable _deviceVars_Name = null;
        AxVariable _deviceVars_Type = null;
        AxVariable _deviceVars_Min = null;
        AxVariable _deviceVars_Max = null;
        AxVariable _deviceVars_Default = null;
        AxVariable _deviceVars_Actual = null;
        AxVariable _deviceVars_RW = null;
        AxVariable _deviceVars_Unit = null;
        AxVariable _deviceVars_VarVorhanden = null;
        Int32 _deviceVars_DeviceID = 0;

        AxVariable _systemVariableName = null;
        Dictionary<UInt16, String> _tempSystemvariableNames = new Dictionary<UInt16, String>();
        AxVariable _systemVariableValue = null;

        HomegearLib.RPC.RPCController _rpc = null;
        HomegearLib.Homegear _homegear = null;
        Boolean _reloading = false;


        Mutex _queueConfigToPushMutex = new Mutex();
        Queue<Dictionary<AxInstance, Dictionary<Device, List<Int32>>>> _queueConfigToPush = new Queue<Dictionary<AxInstance, Dictionary<Device, List<Int32>>>>();
        Dictionary<AxInstance, Dictionary<Device, List<Int32>>> _aktQueue = null;
        Dictionary<String, List<Int32>> _instancesConfigChannels = null;
        Thread _pushConfigThread = null;
        Thread _getActualDeviceDataThread = null;

        Instances _instances = null;
        Mutex _homegearDevicesMutex = new Mutex();
        Mutex _homegearSystemVariablesMutex = new Mutex();

        public void Run(String instanceName, String version)
        {
            try
            {
                try
                {
                    foreach (Process proc in Process.GetProcessesByName("StaKoTecHomeGear"))
                    {
                        if (proc.Id != Process.GetCurrentProcess().Id)
                        {
                            Console.WriteLine("Kille vorhergehende StaKoTecHomeGear.exe (Proc-ID: " + proc.Id.ToString() + ")");
                            proc.Kill();
                        }
                    }
                }
                catch (Exception ex)
                {
                    Console.WriteLine(ex.Message);
                }

                try
                {
                    _aX = new Ax(50);
                    _aX.CycleTime = 20;
                }
                catch (AxException ex)
                {
                    Console.WriteLine(ex.Message + "\r\n" + ex.StackTrace);
                    Dispose();
                }
                
                
                // {{{ Status Texte laden
                _deviceStatusText = new Dictionary<DeviceStatus, String>();
                _deviceStatusText.Add(DeviceStatus.Nichts, "");
                _deviceStatusText.Add(DeviceStatus.OK, "OK");
                _deviceStatusText.Add(DeviceStatus.Fehler, "Falsche Klasse (%s)");
                _deviceStatusText.Add(DeviceStatus.KeineInstanzVorhanden, "Keine Instanz vorhanden");
                _deviceStatusText.Add(DeviceStatus.Warnung, "Warnung");

                // Error- und Warning-States laden
                _errorStates = new List<String>();
                _errorStates.Add("UNREACH");
                _errorStates.Add("STICKY_UNREACH");
                _errorStates.Add("LOWBAT");
                _errorStates.Add("CENTRAL_ADDRESS_SPOOFED");

                _warningStates = new List<String>();
                _warningStates.Add("CONFIG_PENDING");
                // }}}


                _firstInitDevices = new List<Int32>();
                _getActualDeviceDataDictionary = new Dictionary<Device, AxInstance>();
                _ignoredVariableChanged = new List<String>();

                _aX.ShuttingDown += aX_ShuttingDown;
                _aX.SpsIdChangedAfter += _aX_SpsIdChanged;

                _mainInstance = new AxInstance(_aX, instanceName); //Instanz-Objekt erstellen
                AxVariable[] variables = _mainInstance.Variables;
                foreach (AxVariable variable in variables)
                    variable.Events = true;
                Logging.Init(_aX, _mainInstance);
                Logging.WriteLog(LogLevel.Always, _mainInstance, "", "StaKoTecHomegear V " + version + " started");
                _varConverter = new VariableConverter(_mainInstance);
                _instances = new Instances(_aX, _mainInstance);
                _instances.VariableValueChanged += OnInstanceVariableValueChanged;
                _instances.SubinstanceVariableValueChanged += OnSubinstanceVariableValueChanged;

                _instancesConfigChannels = new Dictionary<String, List<Int32>>();

                AxVariable aXVersion = _mainInstance.Get("Version");
                aXVersion.Set(version);
                AxVariable init = _mainInstance.Get("Init");
                init.Set(false);
                init.ValueChanged += init_ValueChanged;
                AxVariable pairingMode = _mainInstance.Get("PairingMode");
                pairingMode.Set(false);
                pairingMode.ValueChanged += pairingMode_ValueChanged;
                AxVariable searchDevices = _mainInstance.Get("SearchDevices");
                searchDevices.Set(false);
                searchDevices.ValueChanged += searchDevices_ValueChanged;
                AxVariable deviceUnpair = _mainInstance.Get("DeviceUnpair");
                deviceUnpair.Set(false);
                deviceUnpair.ValueChanged += deviceUnpair_ValueChanged;
                AxVariable deviceReset = _mainInstance.Get("DeviceReset");
                deviceReset.Set(false);
                deviceReset.ValueChanged += deviceReset_ValueChanged;
                AxVariable deviceRemove = _mainInstance.Get("DeviceRemove");
                deviceRemove.Set(false);
                deviceRemove.ValueChanged += deviceRemove_ValueChanged;
                AxVariable deviceUpdate = _mainInstance.Get("DeviceUpdate");
                deviceUpdate.Set(false);
                deviceUpdate.ValueChanged += deviceUpdate_ValueChanged;
                AxVariable changeInterface = _mainInstance.Get("ChangeInterface");
                changeInterface.Set(false);
                changeInterface.ValueChanged += changeInterface_ValueChanged;
                AxVariable getDeviceVars = _mainInstance.Get("GetDeviceVars");
                getDeviceVars.Set(false);
                getDeviceVars.ValueChanged += getDeviceVars_ValueChanged;
                AxVariable getDeviceConfigVars = _mainInstance.Get("GetDeviceConfigVars");
                getDeviceConfigVars.Set(false);
                getDeviceConfigVars.ValueChanged += getDeviceVars_ValueChanged;
                AxVariable setAllRoaming = _mainInstance.Get("SetAllRoaming");
                setAllRoaming.Set(false);
                setAllRoaming.ValueChanged += setAllRoaming_ValueChanged;
                AxVariable homegearErrorQuit = _mainInstance.Get("HomegearErrorQuit");
                UInt16 x = 0;
                for (x = 0; x < homegearErrorQuit.Length; x++)
                    homegearErrorQuit.Set(x, false);
                homegearErrorQuit.ArrayValueChanged += homegearErrorQuit_ArrayValueChanged;
                AxVariable homegearErrorState = _mainInstance.Get("HomegearErrorState");
                Int16 errorCount = 0;
                for (x = 0; x < homegearErrorState.Length; x++)
                {
                    if (homegearErrorState.GetLongInteger(x) > 0)
                        errorCount++;
                }
                _mainInstance.Get("HomegearError").Set((errorCount > 0));

                //{{{ SystemVariables
                _systemVariableName = _mainInstance.Get("SystemVariableName");
                _systemVariableValue = _mainInstance.Get("SystemVariableValue");
                _systemVariableName.ArrayValueChanged += _systemVariableName_ArrayValueChanged;
                _systemVariableValue.ArrayValueChanged += _systemVariableValue_ArrayValueChanged;
                //}}}

                //{{{ DeviceVars
                _homegearInterfaces = _mainInstance.Get("HomegearInterfaces");
                _deviceID = _mainInstance.Get("DeviceID");
                _deviceInstance = _mainInstance.Get("DeviceInstance");
                _deviceRemark = _mainInstance.Get("DeviceRemark");
                _deviceTypeString = _mainInstance.Get("DeviceTypeString");
                _deviceState = _mainInstance.Get("DeviceState");
                _deviceFirmware = _mainInstance.Get("DeviceFirmware");
                _deviceInterface = _mainInstance.Get("DeviceInterface");
                _deviceStateColor = _mainInstance.Get("DeviceStateColor");

                _deviceVars_Name = _mainInstance.Get("DeviceVars_Name");
                _deviceVars_Type = _mainInstance.Get("DeviceVars_Type");
                _deviceVars_Min = _mainInstance.Get("DeviceVars_Min");
                _deviceVars_Max = _mainInstance.Get("DeviceVars_Max");
                _deviceVars_Default = _mainInstance.Get("DeviceVars_Default");
                _deviceVars_Actual = _mainInstance.Get("DeviceVars_Actual");
                _deviceVars_Actual.ArrayValueChanged += _deviceVars_Actual_ArrayValueChanged;
                _deviceVars_RW = _mainInstance.Get("DeviceVars_RW");
                _deviceVars_Unit = _mainInstance.Get("DeviceVars_Dimension");
                _deviceVars_VarVorhanden = _mainInstance.Get("DeviceVars_VarVorhanden");
                //}}}

                Int32 axStartID = _mainInstance.Get("StartID").GetInteger();
                Int32 axStartID_old = axStartID;

                AxVariable lifetick = _mainInstance.Get("Lifetick");
                AxVariable aXcycleCounter = _mainInstance.Get("CycleCounter");
                Int32 cycleCounter = 0;
                Int32 connectionTimeout = 0;

                AxVariable mainInstanceErr = _mainInstance.Get("err");

                AxVariable start_CAPI_Release = _mainInstance.Get("Start_CAPI_Release");
                start_CAPI_Release.ValueChanged += start_CAPI_Release_ValueChanged;
                if (!start_CAPI_Release.GetBool())
                    Dispose();

                _mainInstance.Get("RPC_InitComplete").Set(false);
                _mainInstance.Get("CAPI_Running").Set(true);

                HomeGearConnect();

                UInt32 j = 0;
                Stopwatch stopWatch = new Stopwatch();
                Stopwatch stopWatchThreadTimer = new Stopwatch();
                Double lastCycletime = 0;
                Double cycletimerServiceMessages = 999;
                Double cycletimerInterfaceCheck = 0;
                while (!_disposing)
                {
                    try
                    {
                        stopWatchThreadTimer.Reset();
                        stopWatchThreadTimer.Start();
                        stopWatch.Reset();
                        stopWatch.Start();

                        cycletimerServiceMessages += lastCycletime;
                        cycletimerInterfaceCheck += lastCycletime;

                        lifetick.Set(true);
                        aXcycleCounter.Set(cycleCounter);
                        cycleCounter++;

                        if (axStartID != axStartID_old)
                        {
                            Logging.WriteLog(LogLevel.Info, _mainInstance, "StartID has changed! Exiting!!!");
                            Dispose();
                            continue;
                        }

                        _logLevel = (LogLevel)_mainInstance.Get("LogLevel").GetLongInteger();
                        //Zu übertragende Config-Parameter abarbeiten
                        if ((_queueConfigToPush.Count > 0) && _initCompleted && (_pushConfigThread == null || !_pushConfigThread.IsAlive || _pushConfigThread.ThreadState == System.Threading.ThreadState.Aborted))
                        {
                            _pushConfigThread = new Thread(PushConfig);
                            _pushConfigThread.Start();

                        }

                        if (!_rpc.IsConnected)
                        {
                            _mainInstance.Get("RPC_InitComplete").Set(false);
                            _initCompleted = false;

                            _firstInitDevices.Clear();

                            if (connectionTimeout > 0)
                                Logging.WriteLog(LogLevel.Info, _mainInstance, "Waiting for RPC-Server connection... (" + (connectionTimeout * 5).ToString() + " s)");

                            if ((connectionTimeout > 6) && !mainInstanceErr.GetBool())
                            {
                                Logging.WriteLog(LogLevel.Error, _mainInstance, "Waiting for RPC-Server connection... (" + (connectionTimeout * 5).ToString() + " s)");
                            }


                            Thread.Sleep(5000);
                            connectionTimeout++;
                            continue;
                        }
                        else
                        {
                            if (connectionTimeout > 0)  //Wenn verbindung wieder hergestellt wurde, neu laden
                            {
                                _homegearDevicesMutex.WaitOne();
                                if (!_reloading)
                                {
                                    _reloading = true;
                                    _homegear.Reload();
                                }
                                _homegearDevicesMutex.ReleaseMutex();
                            }
                            connectionTimeout = 0;
                        }
                        

                        //Allen Devices einen Lifetick senden um DataValid zu generieren
                        if (_initCompleted)
                            _instances.Lifetick();


                        if (_homegear != null && (_pushConfigThread == null || !_pushConfigThread.IsAlive) && _initCompleted && (cycletimerServiceMessages >= 10))
                        {
                            cycletimerServiceMessages = 0;
                            Boolean serviceMessageVorhanden = false;
                            List<ServiceMessage> serviceMessages = _homegear.ServiceMessages;
                            Int16 alarmCounter = 0;
                            Int16 warningCounter = 0;
                            x = 0;
                            AxVariable aX_serviceMessages = _mainInstance.Get("ServiceMessages");
                            foreach (ServiceMessage message in serviceMessages)
                            {
                                if (_errorStates.Contains(message.Message))
                                    alarmCounter++;
                                if (_warningStates.Contains(message.Message))
                                    warningCounter++;

                                String aktMessage = "Device ID: " + message.PeerID.ToString();
                                if (_homegear.Devices.ContainsKey(message.PeerID))
                                {
                                    if (_homegear.Devices[message.PeerID].Name.Length > 0)
                                        aktMessage += " " + _homegear.Devices[message.PeerID].Name;

                                    aktMessage += " (" + _homegear.Devices[message.PeerID].TypeString + ")";
                                }
                                aktMessage += " " + "Channel: " + message.Channel.ToString() + " " + message.Type + " = " + message.Value.ToString();
                                if (x < aX_serviceMessages.Length)
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

                        //Logging.WriteLog("cycletimerInterfaceCheck: " + (cycletimerInterfaceCheck).ToString() + "s");
                        //Alle 60 Sekunden checken wie lange es her ist, dass die letzten Pakete Empfangen / gesendet wurden
                        Int32 currentTime = (Int32)(DateTime.UtcNow.Subtract(new DateTime(1970, 1, 1))).TotalSeconds;
                        if (_homegear != null && _initCompleted && (cycletimerInterfaceCheck >= 600))
                        {
                            cycletimerInterfaceCheck = 0;
                            //Firmwareupgrades prüfen
                            _homegearDevicesMutex.WaitOne();
                            lock (_instances._mutex)
                            {
                                foreach (KeyValuePair<Int32, AxInstance> aktInstance in _instances)
                                {
                                    deviceCheckFirmwareUpdates(aktInstance.Key);
                                }
                            }
                            _homegearDevicesMutex.ReleaseMutex();
                        }
                        //Logging.WriteLog("Cycle-Dauer: " + (lastCycletime).ToString() + "s");
                        j++;
                    }
                    catch (Exception ex)
                    {
                        Logging.WriteLog(LogLevel.Error, _mainInstance, ex.Message, ex.StackTrace);
                    }

                    stopWatchThreadTimer.Stop();
                    TimeSpan elapsedTotalThreadTimer = stopWatchThreadTimer.Elapsed;
                    Int32 sleep = 100 - (Int32)elapsedTotalThreadTimer.TotalMilliseconds;
                    if (sleep < 0)
                        sleep = 0;
                    Thread.Sleep(sleep);

                    stopWatch.Stop();
                    TimeSpan elapsedTotal = stopWatch.Elapsed;
                    lastCycletime = (elapsedTotal.TotalMilliseconds / 1000);
                    //Logging.WriteLog("Last Cycletime: " + lastCycletime.ToString());
                }
            }
            catch(Exception ex)
            {
                Logging.WriteLog(LogLevel.Error, _mainInstance, ex.Message, ex.StackTrace);
            }
        }

        void _deviceVars_Actual_ArrayValueChanged(AxVariable sender, ushort index, AxVariableValue value, DateTime timestamp)
        {
            try
            {
                Int32 deviceID = _mainInstance.Get("ActionID").GetLongInteger();
                String variableName = _deviceVars_Name.GetString(index);
                Logging.WriteLog(LogLevel.Info, _mainInstance, "Variable " + variableName + " for Device ID " + deviceID.ToString() + " has changed");

                String name = "";
                Int32 channel = -1;
                String type = "";
                try
                {
                    if (variableName.Length >= 5)
                    {
                        name = variableName.Substring(0, (variableName.Length - 4));
                        type = variableName.Substring((variableName.Length - 3), 1);
                        Int32.TryParse(variableName.Substring((variableName.Length - 2), 2), out channel);
                    }
                }
                catch (Exception ex)
                {
                    Logging.WriteLog(LogLevel.Error, _mainInstance, ex.Message, ex.StackTrace);
                }
                Logging.WriteLog(LogLevel.Debug, _mainInstance, "HomegearVariable " + name + " - Channel: " + channel.ToString() + "  - Type: " + type);
                if (name.Length > 0)
                {
                    _homegearDevicesMutex.WaitOne();
                    if (_homegear.Devices.ContainsKey(deviceID))
                    {
                        if (type == "V")
                        {
                            if (_homegear.Devices[deviceID].Channels.ContainsKey(channel) && _homegear.Devices[deviceID].Channels[channel].Variables.ContainsKey(name))
                            {
                                Variable aktVariable = _homegear.Devices[deviceID].Channels[channel].Variables[name];
                                if (aktVariable.Writeable)
                                {
                                    Logging.WriteLog(LogLevel.Debug, _mainInstance, "HomegearVariable: PeerID: " + aktVariable.PeerID + " - Name: " + aktVariable.Name + " - Channel: " + aktVariable.Channel.ToString() + " - Type: " + aktVariable.Type.ToString());
                                    if ((aktVariable.Type == VariableType.tAction) || (aktVariable.Type == VariableType.tBoolean))
                                    {
                                        aktVariable.BooleanValue = (sender.GetString(index).ToLower() == "true") ? true : false;
                                    }
                                    else if (aktVariable.Type == VariableType.tDouble)
                                    {
                                        aktVariable.DoubleValue = Convert.ToDouble(sender.GetString(index));
                                    }
                                    else if ((aktVariable.Type == VariableType.tEnum) || (aktVariable.Type == VariableType.tInteger))
                                    {
                                        aktVariable.IntegerValue = Convert.ToInt32(sender.GetString(index));
                                    }
                                    else if (aktVariable.Type == VariableType.tString)
                                    {
                                        aktVariable.StringValue = sender.GetString(index);
                                    }
                                    else
                                    {
                                        Logging.WriteLog(LogLevel.Error, _mainInstance, "VariableType " + aktVariable.Type.ToString() + " is not supported");
                                    }
                                }
                                else
                                {
                                    Logging.WriteLog(LogLevel.Error, _mainInstance, "HomegearVariable " + name + " from DeviceID " + deviceID.ToString() + " is not writeable!");
                                    sender.Set(index, aktVariable.ToString());
                                }
                            }
                            else
                                Logging.WriteLog(LogLevel.Error, _mainInstance, "HomegearVariable " + name + " or channel " + channel.ToString() + " from DeviceID " + deviceID.ToString() + " does not exist!");
                        }
                        else if (type == "C")
                        {
                            if (_homegear.Devices[deviceID].Channels.ContainsKey(channel) && _homegear.Devices[deviceID].Channels[channel].Config.ContainsKey(name))
                            {
                                ConfigParameter aktVariable = _homegear.Devices[deviceID].Channels[channel].Config[name];
                                if (aktVariable.Writeable)
                                {
                                    Logging.WriteLog(LogLevel.Debug, _mainInstance, "HomegearConfigVariable: PeerID: " + aktVariable.PeerID + " - Name: " + aktVariable.Name + " - Channel: " + aktVariable.Channel.ToString() + " - Type: " + aktVariable.Type.ToString());
                                    if ((aktVariable.Type == VariableType.tAction) || (aktVariable.Type == VariableType.tBoolean))
                                    {
                                        aktVariable.BooleanValue = (sender.GetString(index).ToLower() == "true") ? true : false;
                                    }
                                    else if (aktVariable.Type == VariableType.tDouble)
                                    {
                                        aktVariable.DoubleValue = Convert.ToDouble(sender.GetString(index));
                                    }
                                    else if ((aktVariable.Type == VariableType.tEnum) || (aktVariable.Type == VariableType.tInteger))
                                    {
                                        aktVariable.IntegerValue = Convert.ToInt32(sender.GetString(index));
                                    }
                                    else if (aktVariable.Type == VariableType.tString)
                                    {
                                        aktVariable.StringValue = sender.GetString(index);
                                    }
                                    else
                                    {
                                        Logging.WriteLog(LogLevel.Error, _mainInstance, "VariableType " + aktVariable.Type.ToString() + " is not supported");
                                    }
                                }
                                else
                                {
                                    Logging.WriteLog(LogLevel.Error, _mainInstance, "HomegearVariable " + name + " from DeviceID " + deviceID.ToString() + " is not writeable!");
                                    sender.Set(index, aktVariable.ToString());
                                }
                            }
                            else
                                Logging.WriteLog(LogLevel.Error, _mainInstance, "HomegearVariable " + name + " or channel " + channel.ToString() + " from DeviceID " + deviceID.ToString() + " does not exist!");
                        }
                        else
                            Logging.WriteLog(LogLevel.Error, _mainInstance, type + " is no valid Type!");
                    }
                    else
                        Logging.WriteLog(LogLevel.Error, _mainInstance, "No DeviceID " + deviceID.ToString() + " found!");
                }
                else
                    Logging.WriteLog(LogLevel.Error, _mainInstance, "HomegearVariable " + name + " is no valid VariableName");

                _homegearDevicesMutex.ReleaseMutex();
            }
            catch (Exception ex)
            {
                try { _homegearDevicesMutex.ReleaseMutex(); }
                catch (Exception) { }
                Logging.WriteLog(LogLevel.Error, _mainInstance, ex.Message, ex.StackTrace);
            }
        }

        void _systemVariableName_ArrayValueChanged(AxVariable sender, ushort index, AxVariableValue value, DateTime timestamp)
        {
            try
            {
                _homegearSystemVariablesMutex.WaitOne();
                String varName = _systemVariableName.GetString(index).Trim();
                if ((varName == "") && (_tempSystemvariableNames.ContainsKey(index)))
                {
                    Logging.WriteLog(LogLevel.Info, _mainInstance, "Systemvariable '" + _tempSystemvariableNames[index] + "' wurde von aX gelöscht");
                    _rpc.DeleteSystemVariable(new SystemVariable(_tempSystemvariableNames[index], 0));
                    _tempSystemvariableNames.Remove(index);
                }
                else if ((varName != "") && (!_tempSystemvariableNames.ContainsKey(index)))
                {
                    Logging.WriteLog(LogLevel.Info, _mainInstance, "Systemvariable '" + varName + "' wurde von aX erstellt");
                    _rpc.SetSystemVariable(new SystemVariable(varName, 0));
                    _tempSystemvariableNames.Add(index, varName);
                }
                else if ((varName != "") && (_tempSystemvariableNames.ContainsKey(index)))  //SystemVariable wurde umbenannt
                {
                    Logging.WriteLog(LogLevel.Info, _mainInstance, "Systemvariable '" + _tempSystemvariableNames[index] + "' wurde von aX umbenannt in '" + varName + "'");
                    _rpc.SetSystemVariable(new SystemVariable(varName, _systemVariableValue.GetString(index)));
                    _homegearSystemVariablesMutex.ReleaseMutex();
                    Thread.Sleep(1000);
                    _homegearSystemVariablesMutex.WaitOne();
                    _rpc.DeleteSystemVariable(new SystemVariable(_tempSystemvariableNames[index], 0));
                }
            }
            catch (Exception ex)
            {
                Logging.WriteLog(LogLevel.Error, _mainInstance, ex.Message, ex.StackTrace);
            }
            _homegearSystemVariablesMutex.ReleaseMutex();
        }

        void _systemVariableValue_ArrayValueChanged(AxVariable sender, ushort index, AxVariableValue value, DateTime timestamp)
        {
            try
            {
                _homegearSystemVariablesMutex.WaitOne();
                String varName = _systemVariableName.GetString(index).Trim();
                if (varName != "")
                {
                    Logging.WriteLog(LogLevel.Info, _mainInstance, "Wert von Systemvariable '" + varName + "' in aX geändert in: " + sender.GetString(index));
                    _rpc.SetSystemVariable(new SystemVariable(varName, sender.GetString(index)));
                    _homegear.SystemVariables.Reload();
                }
                else
                    sender.Set(index, "");
            }
            catch (Exception ex)
            {
                Logging.WriteLog(LogLevel.Error, _mainInstance, ex.Message, ex.StackTrace);
            }
            _homegearSystemVariablesMutex.ReleaseMutex();
        }

        void changeInterface_ValueChanged(AxVariable sender, AxVariableValue value, DateTime timestamp)
        {
            try
            {
                Int32 deviceID = _mainInstance.Get("ChangeInterfaceDeviceID").GetLongInteger();
                String deviceInterface = _mainInstance.Get("ChangeInterfaceName").GetString();

                Logging.WriteLog(LogLevel.Debug, _mainInstance, "Change interface from DeciveID " + deviceID.ToString() + " to: " + deviceInterface);
                _homegearDevicesMutex.WaitOne();
                if (_homegear.Devices.ContainsKey(deviceID))
                {
                    Dictionary<String, Interface> rpcInterfaces = _rpc.ListInterfaces();
                    foreach (KeyValuePair<String, Interface> aktInterface in rpcInterfaces)
                    {
                        if (aktInterface.Key == deviceInterface)
                            _rpc.SetInterface(deviceID, aktInterface.Value);
                    }
                }
                _homegearDevicesMutex.ReleaseMutex();
                sender.Set(false);
                _homegear.Reload();
                Logging.WriteLog(LogLevel.Debug, _mainInstance, "Change interface ready");
            }
            catch (Exception ex)
            {
                Logging.WriteLog(LogLevel.Error, _mainInstance, ex.Message, ex.StackTrace);
            }
        }

        Double getUnixTime()
        {
            Double currentTime = (DateTime.UtcNow.Subtract(new DateTime(1970, 1, 1))).TotalMilliseconds;
            return currentTime;
        }


        void setAllRoaming_ValueChanged(AxVariable sender, AxVariableValue value, DateTime timestamp)
        {
            try
            {
                sender.Set(false);
                Logging.WriteLog(LogLevel.Info, _mainInstance, "Setze Roaming aktiv für alle Geräte");
                foreach (KeyValuePair<long, Device> aktDevice in _homegear.Devices)
                {
                    if (!aktDevice.Value.Channels.ContainsKey(0))
                        continue;

                    if (!aktDevice.Value.Channels[0].Config.ContainsKey("ROAMING"))
                        continue;

                    if (!aktDevice.Value.Channels[0].Config["ROAMING"].BooleanValue)
                    {
                        aktDevice.Value.Channels[0].Config["ROAMING"].BooleanValue = true;
                        Logging.WriteLog(LogLevel.Info, _mainInstance, "Gerät " + aktDevice.Value.Name + " (ID: " + aktDevice.Value.ID + ") hinzugefügt");
                        aktDevice.Value.Channels[0].Config.Put();
                    }
                    else
                        Logging.WriteLog(LogLevel.Info, _mainInstance, "Gerät " + aktDevice.Value.Name + " (ID: " + aktDevice.Value.ID + ") war schon auf Roaming");
                }
            }
            catch (Exception ex)
            {
                Logging.WriteLog(LogLevel.Error, _mainInstance, ex.Message, ex.StackTrace);
            }
        }

        void PushConfig()
        {
            try
            {
                Logging.WriteLog(LogLevel.Info, _mainInstance, "Pushing Config-Thread gestartet");
                while (_queueConfigToPush.Count > 0) 
                {
                    _queueConfigToPushMutex.WaitOne();
                    _aktQueue = _queueConfigToPush.Dequeue();
                    _queueConfigToPushMutex.ReleaseMutex();
                    foreach (KeyValuePair<AxInstance, Dictionary<Device, List<Int32>>> aktInstance in _aktQueue)
                    {
                        foreach (KeyValuePair<Device, List<Int32>> aktDevice in aktInstance.Value)
                        {
                            foreach (Int32 aktChannel in aktDevice.Value)
                            {
                                Logging.WriteLog(LogLevel.Info, aktInstance.Key, "[" + aktInstance.Key.Path + "] Pushe Config für Kanal " + aktChannel.ToString());
                                aktInstance.Key["Status"].Set("Pushe Config für Kanal " + aktChannel.ToString());
                                aktDevice.Key.Channels[aktChannel].Config.Put();
                                Thread.Sleep(1000);  //Sendepause zuwischen den Geräten falls dirverse Geräte gleichzeitig pushen wollen
                            }
                        }
                        if (aktInstance.Key.VariableExists("ConfigValuesChanged"))
                            aktInstance.Key.Get("ConfigValuesChanged").Set(false);
                        Thread.Sleep(1000);  //Sendepause zuwischen den Geräten falls dirverse Geräte gleichzeitig pushen wollen
                   }
                    _aktQueue = null;
                }
            }
            catch(ThreadAbortException)
            {
                if (_aktQueue != null)
                {
                    _queueConfigToPushMutex.WaitOne();
                    _queueConfigToPush.Enqueue(_aktQueue);
                    _queueConfigToPushMutex.ReleaseMutex();
                }
                Logging.WriteLog(LogLevel.Info, _mainInstance, "PushConfig Thread abgebrochen. Geht gleich weiter...");
            }
            catch (Exception ex)
            {
                Logging.WriteLog(LogLevel.Error, _mainInstance, ex.Message, ex.StackTrace);
            }
        }

        String deviceCheckFirmwareUpdates(Int32 deviceID)
        {
            String firmware = "";
            if ((_homegear.Devices.ContainsKey(deviceID)) && (_instances.ContainsKey(deviceID)))
            {
                Device aktDevice = _homegear.Devices[deviceID];
                AxInstance aktInstance = _instances[deviceID];

                if (aktInstance.VariableExists("AktuelleFirmware"))
                    aktInstance.Get("AktuelleFirmware").Set(aktDevice.Firmware + " (" + aktDevice.AvailableFirmware + ")");
                if ((aktDevice.AvailableFirmware != "") && (aktDevice.Firmware != aktDevice.AvailableFirmware))
                {
                    firmware = aktDevice.Firmware + "/" + aktDevice.AvailableFirmware;
                    if (aktInstance.VariableExists("FirmwareupdateVorhanden"))
                        aktInstance.Get("FirmwareupdateVorhanden").Set(true);
                }
                else
                {
                    firmware = aktDevice.Firmware;
                    if (aktInstance.VariableExists("FirmwareupdateVorhanden"))
                        aktInstance.Get("FirmwareupdateVorhanden").Set(false);
                }
            }
            return firmware;
        }

        void homegearErrorQuit_ArrayValueChanged(AxVariable sender, ushort index, AxVariableValue value, DateTime timestamp)
        {
            if (sender.GetBool(index))
                aXremoveHomegearError(index);
            sender.Set(index, false);
        }

        void aXremoveHomegearError(UInt16 index)
        {
            try
            {
                //Logging.WriteLog("RemoveHomegearError index " + index.ToString());             
                UInt16 x = 0;
                UInt16 i = 0;
                Int16 errorCount = 0;
                Dictionary<UInt16, Dictionary<String, Int32>> dictHomegearErrors = new Dictionary<UInt16, Dictionary<String, Int32>>();
                AxVariable homegearErrorState = _mainInstance.Get("HomegearErrorState");
                AxVariable homegearErrors = _mainInstance.Get("HomegearErrors");
                for (x = 0, i = 0; x < homegearErrors.Length; x++)
                {
                    if (x != index)
                    {
                        //Logging.WriteLog("Read Index " + x.ToString() + " " + homegearErrors.GetString(x));
                        Dictionary<String, Int32> aktEintrag = new Dictionary<String, Int32>();
                        aktEintrag.Add(homegearErrors.GetString(x), homegearErrorState.GetLongInteger(x));
                        dictHomegearErrors.Add(i, aktEintrag);
                        i++;
                    }
                    //else
                      //  Logging.WriteLog("Skip Read Index " + x.ToString() + " " + homegearErrors.GetString(x));
                }
                for (; i < x; i++ )
                {
                    //Logging.WriteLog("Read Frei Index " + i.ToString());
                    Dictionary<String, Int32> aktEintrag = new Dictionary<String, Int32>();
                    aktEintrag.Add("", 0);
                    dictHomegearErrors.Add(i, aktEintrag);
                }

                foreach (KeyValuePair<UInt16, Dictionary<String, Int32>> aktZeile in dictHomegearErrors)
                {
                    foreach (KeyValuePair<String, Int32> aktError in aktZeile.Value)
                    {
                        //Logging.WriteLog("Write Index " + aktZeile.Key.ToString() + " " + aktError.Key);
                        homegearErrors.Set(aktZeile.Key, aktError.Key);
                        homegearErrorState.Set(aktZeile.Key, aktError.Value);
                    }
                }


                _mainInstance.Get("HomegearError").Set((errorCount > 0));
            }
            catch (Exception ex)
            {
                Logging.WriteLog(LogLevel.Error, _mainInstance, ex.Message, ex.StackTrace);
            }
        }

        void aXAddHomegearError(String message, Int16 level)
        {
            try
            {
                List<String> homegearErrorsTemp = new List<String>();
                List<Int32> homegearErrorStateTemp = new List<Int32>();
                AxVariable homegearErrorState = _mainInstance.Get("HomegearErrorState");
                AxVariable homegearErrors = _mainInstance.Get("HomegearErrors");
                UInt16 x = 0;

                homegearErrorsTemp.Add(DateTime.Now.Hour.ToString("D2") + ":" + DateTime.Now.Minute.ToString("D2") + ":" + DateTime.Now.Second.ToString("D2") + ": " + message);
                homegearErrorStateTemp.Add(level);
                for (x = 0; x < homegearErrors.Length; x++)
                {
                    homegearErrorsTemp.Add(homegearErrors.GetString(x));
                    homegearErrorStateTemp.Add(homegearErrorState.GetLongInteger(x));
                }
                for (x = 0; x < homegearErrors.Length; x++)
                {
                    homegearErrors.Set(x, homegearErrorsTemp[x]);
                    homegearErrorState.Set(x, homegearErrorStateTemp[x]);
                }
                _mainInstance.Get("HomegearError").Set(true);
            }
            catch (Exception ex)
            {
                Logging.WriteLog(LogLevel.Error, _mainInstance, ex.Message, ex.StackTrace);
            }
        }

        String findVarInClass(Int32 id, String varName, VariableType varType,  String Typ, String Kanal)
        {
            try
            {
                bool instanceVarExists = false;
                bool subinstanceVarExists = false;
                AxVariableType type = AxVariableType.axUnsignedShortInteger;

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

                if (((varType == VariableType.tInteger) && (type != AxVariableType.axLongInteger)) ||
                    ((varType == VariableType.tDouble) && (type != AxVariableType.axLongReal)) ||
                    ((varType == VariableType.tBoolean) && (type != AxVariableType.axBool)) ||
                    ((varType == VariableType.tAction) && (type != AxVariableType.axBool)) ||
                    ((varType == VariableType.tEnum) && (type != AxVariableType.axLongInteger)) ||
                    ((varType == VariableType.tString) && (type != AxVariableType.axString)))
                    return "Falscher Variablen-Typ";


                return "OK";
            }
            catch (Exception ex)
            {
                Logging.WriteLog(LogLevel.Error, _mainInstance, ex.Message, ex.StackTrace);
                return ex.Message;
            }
        }

        void getDeviceVars_ValueChanged(AxVariable sender, AxVariableValue value, DateTime timestamp)
        {
            lock (_instances._mutex)
            {
                try
                {
                    _deviceVars_DeviceID = _mainInstance.Get("ActionID").GetLongInteger();

                    UInt16 x = 0;

                    _homegearDevicesMutex.WaitOne();

                    if (_homegear.Devices.ContainsKey(_deviceVars_DeviceID))
                    {
                        Device aktDevice = _homegear.Devices[_deviceVars_DeviceID];
                        foreach (KeyValuePair<long, Channel> aktChannel in aktDevice.Channels)
                        {
                            if (sender.Name == "GetDeviceConfigVars")
                            {
                                sender.Instance["Status"].Set("Get VariableConfigNames for DeviceID: " + _deviceVars_DeviceID.ToString());
                                foreach (KeyValuePair<String, ConfigParameter> configName in aktDevice.Channels[aktChannel.Key].Config)
                                {
                                    if (x >= _deviceVars_Name.Length)
                                    {
                                        Logging.WriteLog(LogLevel.Error, _mainInstance, "Array-Index zu klein bei 'DeviceVars_Name'");
                                        _homegearDevicesMutex.ReleaseMutex();
                                        return;
                                    }

                                    var aktVar = configName.Value;
                                    String minVar = "";
                                    String maxVar = "";
                                    String typ = "";
                                    String defaultVar = "";
                                    String actualVar = "";
                                    String rwVar = "";
                                    String varOK = findVarInClass(_deviceVars_DeviceID, aktVar.Name, aktVar.Type, "C", aktChannel.Key.ToString("D2"));

                                    _varConverter.ParseDeviceConfigVars(aktVar, out minVar, out maxVar, out typ, out defaultVar, out rwVar);

                                    actualVar = _varConverter.HomegearVarToString(aktVar);
                                    _deviceVars_Name.Set(x, aktVar.Name + "_C" + aktChannel.Key.ToString("D2"));
                                    _deviceVars_Type.Set(x, typ);
                                    _deviceVars_Min.Set(x, minVar);
                                    _deviceVars_Max.Set(x, maxVar);
                                    _deviceVars_Default.Set(x, defaultVar);
                                    _deviceVars_Actual.Set(x, actualVar);
                                    _deviceVars_RW.Set(x, rwVar);
                                    _deviceVars_Unit.Set(x, aktVar.Unit);
                                    _deviceVars_VarVorhanden.Set(x, varOK);
                                    x++;
                                }
                            }

                            if (sender.Name == "GetDeviceVars")
                            {
                                sender.Instance["Status"].Set("Get VariableNames for DeviceID: " + _deviceVars_DeviceID.ToString());
                                foreach (KeyValuePair<String, Variable> varName in aktDevice.Channels[aktChannel.Key].Variables)
                                {
                                    if (x >= _deviceVars_Name.Length)
                                    {
                                        Logging.WriteLog(LogLevel.Error, _mainInstance, "Array-Index zu klein bei 'DeviceVars_Name'");
                                        _homegearDevicesMutex.ReleaseMutex();
                                        return;
                                    }

                                    var aktVar = varName.Value;
                                    String minVar = "";
                                    String maxVar = "";
                                    String typ = "";
                                    String defaultVar = "";
                                    String actualVar = "";
                                    String rwVar = "";
                                    String varOK = findVarInClass(_deviceVars_DeviceID, aktVar.Name, aktVar.Type, "V", aktChannel.Key.ToString("D2"));
                                    _varConverter.ParseDeviceVars(aktVar, out minVar, out maxVar, out typ, out defaultVar, out rwVar);

                                    actualVar = _varConverter.HomegearVarToString(aktVar);
                                    _deviceVars_Name.Set(x, aktVar.Name + "_V" + aktChannel.Key.ToString("D2"));
                                    _deviceVars_Type.Set(x, typ);
                                    _deviceVars_Min.Set(x, minVar);
                                    _deviceVars_Max.Set(x, maxVar);
                                    _deviceVars_Default.Set(x, defaultVar);
                                    _deviceVars_Actual.Set(x, actualVar);
                                    _deviceVars_RW.Set(x, rwVar);
                                    _deviceVars_Unit.Set(x, aktVar.Unit);
                                    _deviceVars_VarVorhanden.Set(x, varOK);
                                    x++;
                                }
                            }
                        }

                        for (; x < _deviceVars_Name.Length; x++)
                        {
                            _deviceVars_Name.Set(x, "");
                            _deviceVars_Type.Set(x, "");
                            _deviceVars_Min.Set(x, "");
                            _deviceVars_Max.Set(x, "");
                            _deviceVars_Default.Set(x, "");
                            _deviceVars_Actual.Set(x, "");
                            _deviceVars_RW.Set(x, "");
                            _deviceVars_Unit.Set(x, "");
                            _deviceVars_VarVorhanden.Set(x, "");
                        }
                    }
                    _homegearDevicesMutex.ReleaseMutex();
                }
                catch (Exception ex)
                {
                    try { _homegearDevicesMutex.ReleaseMutex(); }
                    catch (Exception) { }
                    Logging.WriteLog(LogLevel.Error, sender.Instance, ex.Message, ex.StackTrace);
                }
                finally
                {
                    try
                    {
                        sender.Set(false);
                    }
                    catch (Exception ex)
                    {
                        Logging.WriteLog(LogLevel.Error, _mainInstance, ex.Message, ex.StackTrace);
                    }
                }
            }
        }

        void deviceRemove_ValueChanged(AxVariable sender, AxVariableValue value, DateTime timestamp)
        {
            try
            {
                Int32 deviceID = _mainInstance.Get("ActionID").GetLongInteger();
                Logging.WriteLog(LogLevel.Info, _mainInstance, "Removing Device ID " + deviceID.ToString());
                _homegearDevicesMutex.WaitOne();

                _rpc.DeleteDevice(deviceID, RPCDeleteDeviceFlags.Force);

                _homegearDevicesMutex.ReleaseMutex();
                sender.Set(false);
                Logging.WriteLog(LogLevel.Info, _mainInstance, "Removing Device ID " + deviceID.ToString() + " complete");
            }
            catch (Exception ex)
            {
                try{ _homegearDevicesMutex.ReleaseMutex(); } catch (Exception) { }
                Logging.WriteLog(LogLevel.Error, _mainInstance, ex.Message, ex.StackTrace);
            }
        }

        void deviceReset_ValueChanged(AxVariable sender, AxVariableValue value, DateTime timestamp)
        {
            try
            {
                Int32 deviceID = _mainInstance.Get("ActionID").GetLongInteger();
                Logging.WriteLog(LogLevel.Info, _mainInstance, "Resetting Device ID " + deviceID.ToString());
                if (_firstInitDevices.Contains(deviceID))
                    _firstInitDevices.Remove(deviceID);

                _homegearDevicesMutex.WaitOne();

                _rpc.DeleteDevice(deviceID, RPCDeleteDeviceFlags.Reset | RPCDeleteDeviceFlags.Defer);

                _homegearDevicesMutex.ReleaseMutex();
                sender.Set(false);
                Logging.WriteLog(LogLevel.Info, _mainInstance, "Resetting Device ID " + deviceID.ToString() + " complete");
            }
            catch (Exception ex)
            {
                try { _homegearDevicesMutex.ReleaseMutex(); } catch (Exception) { }
                Logging.WriteLog(LogLevel.Error, _mainInstance, ex.Message, ex.StackTrace);
            }
        }

        void deviceUnpair_ValueChanged(AxVariable sender, AxVariableValue value, DateTime timestamp)
        {
            try
            {
                Int32 deviceID = _mainInstance.Get("ActionID").GetLongInteger();
                Logging.WriteLog(LogLevel.Info, _mainInstance, "Unpairing Device ID " + deviceID.ToString());
                if (_firstInitDevices.Contains(deviceID))
                    _firstInitDevices.Remove(deviceID);

                _homegearDevicesMutex.WaitOne();

                _rpc.DeleteDevice(deviceID, RPCDeleteDeviceFlags.Defer);

                _homegearDevicesMutex.ReleaseMutex();
                sender.Set(false);
                Logging.WriteLog(LogLevel.Info, _mainInstance, "Unpairing Device ID " + deviceID.ToString() + " complete");
            }
            catch (Exception ex)
            {
                try { _homegearDevicesMutex.ReleaseMutex(); } catch (Exception) { }
                Logging.WriteLog(LogLevel.Error, _mainInstance, ex.Message, ex.StackTrace);
            }
        }


        void deviceUpdate_ValueChanged(AxVariable sender, AxVariableValue value, DateTime timestamp)
        {
            try
            {
                Int32 deviceID = _mainInstance.Get("ActionID").GetLongInteger();
                Logging.WriteLog(LogLevel.Info, _mainInstance, "Firmwareupdate for Device ID " + deviceID.ToString());
  
                _homegearDevicesMutex.WaitOne();

                if (_homegear.Devices.ContainsKey(deviceID))
                {
                    Device aktDevice = _homegear.Devices[deviceID];
                    if (aktDevice.Firmware != aktDevice.AvailableFirmware)
                    {
                        Logging.WriteLog(LogLevel.Info, _mainInstance, "Firmwareupdate for Device ID " + deviceID.ToString() + " running. Please wait...");
                        aktDevice.UpdateFirmware(false);
                    }
                    else
                    {
                        Logging.WriteLog(LogLevel.Info, _mainInstance, "No Firmwareupdate for Device ID " + deviceID.ToString() + " available.");
                    }
                }
                _homegearDevicesMutex.ReleaseMutex();
                sender.Set(false);
            }
            catch (Exception ex)
            {
                try { _homegearDevicesMutex.ReleaseMutex(); }
                catch (Exception) { }
                Logging.WriteLog(LogLevel.Error, _mainInstance, ex.Message, ex.StackTrace);
            }
        }


        void pairingMode_ValueChanged(AxVariable sender, AxVariableValue value, DateTime timestamp)
        {
            try
            {
                _homegear.EnablePairingMode(sender.GetBool());
            }
            catch (Exception ex)
            {
                Logging.WriteLog(LogLevel.Error, _mainInstance, ex.Message, ex.StackTrace);
            }
        }

        void searchDevices_ValueChanged(AxVariable sender, AxVariableValue value, DateTime timestamp)
        {
            try
            {
                if (sender.GetBool())
                {
                    Logging.WriteLog(LogLevel.Info, _mainInstance, "Suche neue Geräte");
                    _rpc.SearchDevices();
                    sender.Set(false);
                }
            }
            catch (Exception ex)
            {
                Logging.WriteLog(LogLevel.Error, _mainInstance, ex.Message, ex.StackTrace);
            }
        }

        void start_CAPI_Release_ValueChanged(AxVariable sender, AxVariableValue value, DateTime timestamp)
        {
            try
            {
                if (!sender.GetBool())
                {
                    _mainInstance.Get("err").Set(false);
                    Logging.WriteLog(LogLevel.Always, _mainInstance, "Beende StaKoTecHomeGear.exe");

                    Dispose();
                }
            }
            catch (Exception ex)
            {
                Logging.WriteLog(LogLevel.Error, _mainInstance, ex.Message, ex.StackTrace);
            }
        }

        void getInterfaces()
        {
            UInt16 x = 0;
            foreach(KeyValuePair<String, Interface> aktInterface in _homegear.Interfaces)
            {
                if (x < _homegearInterfaces.Length)
                    _homegearInterfaces.Set(x, aktInterface.Value.ID);
                x++;
            }
            for(; x < _homegearInterfaces.Length; x++)
                _homegearInterfaces.Set(x, "");
        }

        void init_ValueChanged(AxVariable sender, AxVariableValue value, DateTime timestamp)
        {
            try
            {
                if (!sender.GetBool())
                    return;

                _initCompleted = false;
                if (_pushConfigThread != null && _pushConfigThread.IsAlive)
                    _pushConfigThread.Abort();


                if (_getActualDeviceDataThread != null && _getActualDeviceDataThread.IsAlive)
                    _getActualDeviceDataThread.Abort();

                if (_getActualDeviceDataDictionary.Count() > 0)
                    _getActualDeviceDataDictionary.Clear();

                UInt16 x = 0;
                Logging.WriteLog(LogLevel.Info, _mainInstance, "Init Devices");

                UpdateSystemVariables();

                _mainInstance.Get("HomeGearVersion").Set(_homegear.Version.Replace("Homegear ", ""));
                getInterfaces();

                _homegearDevicesMutex.WaitOne();
                Logging.WriteLog(LogLevel.Info, _mainInstance, "Reloading Instances");
                _instances.Reload(_homegear.Devices);
                lock (_instances._mutex)
                {
                    try
                    {
                        x = 0;
                        //Erstmal alle ColorStates löschen, damit sichtbar wird welches Device bereite geparst wurde
                        for (x = 0; x < _deviceStateColor.Length; x++)
                            _deviceStateColor.Set(x, (Int16)DeviceStatus.Nichts);

                        x = 0;
                        foreach (KeyValuePair<long, Device> devicePair in _homegear.Devices)
                        {
                            Int32 key = (Int32)devicePair.Key;
                            if (key > 1000000000)  //Teams nicht anzeigen
                                continue;

                            _deviceID.Set(x, devicePair.Key);

                            Logging.WriteLog(LogLevel.Debug, _mainInstance, "Parse DeviceID " + devicePair.Key.ToString());
                            if (_instances.ContainsKey(key))
                            {
                                AxInstance aktInstanz = _instances[key];
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

                                    _deviceFirmware.Set(x, deviceCheckFirmwareUpdates(key));
                                    try { _deviceInterface.Set(x, devicePair.Value.Interface.ID); }
                                    catch (Exception) { /* It's a virtual device - it has no interface */ }

                                    _deviceState.Set(x, _deviceStatusText.ContainsKey(DeviceStatus.OK) ? _deviceStatusText[DeviceStatus.OK] : "OK");
                                    _deviceStateColor.Set(x, (Int16)DeviceStatus.OK);
                                    if (aktInstanz.VariableExists("SerialNo"))
                                        aktInstanz.Get("SerialNo").Set(devicePair.Value.SerialNumber);
                                    if (aktInstanz.VariableExists("Name"))
                                        aktInstanz.Get("Name").Set(devicePair.Value.Name);
                                    if (aktInstanz.VariableExists("InterfaceID"))
                                        aktInstanz.Get("InterfaceID").Set(devicePair.Value.Interface.ID);

                                    //Aktuelle Config- und Statuswerte Werte auslesen
                                    if (!_firstInitDevices.Contains(key))
                                    {
                                        _getActualDeviceDataDictionary.Add(devicePair.Value, aktInstanz);
                                        _firstInitDevices.Add(key);
                                    }

                                    aktInstanz.Get("ConfigValuesChanged").Set(false);
                                    _instances.Lifetick();
                                    _mainInstance.Get("Lifetick").Set(true);
                                    Thread.Sleep(1);
                                }
                                else  //if (classnames[devicePair.Key] == devicePair.Value.TypeString)
                                {
                                    _deviceTypeString.Set(x, devicePair.Value.TypeString);
                                    _deviceInstance.Set(x, "");
                                    _deviceRemark.Set(x, "");
                                    _deviceState.Set(x, _deviceStatusText.ContainsKey(DeviceStatus.Fehler) ? _deviceStatusText[DeviceStatus.Fehler].Replace("%s", aktInstanz.ClassName) : "Falsche Klasse! (" + aktInstanz.ClassName + ")");
                                    _deviceStateColor.Set(x, (Int16)DeviceStatus.Fehler);
                                    _instances.Remove(key);
                                }
                            }
                            else
                            {
                                _deviceTypeString.Set(x, devicePair.Value.TypeString);
                                _deviceInstance.Set(x, "");
                                _deviceRemark.Set(x, "");
                                _deviceState.Set(x, _deviceStatusText.ContainsKey(DeviceStatus.KeineInstanzVorhanden) ? _deviceStatusText[DeviceStatus.KeineInstanzVorhanden] : "Keine Instanz gefunden");
                                _deviceStateColor.Set(x, (Int16)DeviceStatus.KeineInstanzVorhanden);
                            }
                            x++;
                        }
                    }
                    catch (Exception ex)
                    {
                        Logging.WriteLog(LogLevel.Error, _mainInstance, ex.Message, ex.StackTrace);
                    }
                    _instances.mutexIsLocked = false;
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
                _homegearDevicesMutex.ReleaseMutex();


                _getActualDeviceDataThread = new Thread(getActualDeviceData);
                _getActualDeviceDataThread.Start();
            }
            catch (Exception ex)
            {
                try { _homegearDevicesMutex.ReleaseMutex(); }  catch (Exception) { }
                Logging.WriteLog(LogLevel.Error, sender.Instance, ex.Message, ex.StackTrace);
            }
            finally
            {
                try
                {
                    sender.Set(false);
                    Logging.WriteLog(LogLevel.Info, _mainInstance, "Init completely finished");
                    //_mainInstance.Get("PolledVariables").Set(_mainInstance.PolledVariablesCount + _instances.PolledVariablesCount);
                }
                catch (Exception ex)
                {
                    Logging.WriteLog(LogLevel.Error, _mainInstance, ex.Message, ex.StackTrace);
                }
            }
        }

        void getActualDeviceData()
        {

            lock (_instances._mutex)
            {
                try
                {
                    _homegearDevicesMutex.WaitOne();
                    foreach (KeyValuePair<Device, AxInstance> aktPair in _getActualDeviceDataDictionary)
                    {
                        Device aktDevice = aktPair.Key;
                        AxInstance aktInstanz = aktPair.Value;

                        Logging.WriteLog(LogLevel.Debug, _mainInstance, "Hole aktuelle Variablen von Device " + aktDevice.ID.ToString() + " (" + aktDevice.Name + ")");
                        foreach (KeyValuePair<long, Channel> aktChannel in aktDevice.Channels)
                        {
                            Int32 id = (Int32)aktDevice.ID;
                            foreach (KeyValuePair<String, Variable> Wert in aktDevice.Channels[aktChannel.Key].Variables)
                            {
                                var aktVar = Wert.Value;
                                if (!aktVar.Writeable)  //Action-Variablen nicht beim Init auslesen (Wie z.B. PRESS_SHORT oder so), da HomeGear speichert, dass die Variable irgendwann mal auf 1 war und somit immer beim warmstart alle bisher gedrückten Taster noch einmal auf 1 gesetzt werden
                                    continue;

                                String aktVarName = aktVar.Name + "_V" + aktChannel.Key.ToString("D2");
                                if (aktInstanz.VariableExists(aktVarName))
                                {
                                    AxVariable aktAxVar = aktInstanz.Get(aktVarName);
                                    if (aktAxVar != null)
                                    {
                                        if (aktVar.Type == VariableType.tAction)
                                        {
                                            aktAxVar.Set(false);
                                            continue;
                                        }
                                        _varConverter.SetAxVariable(aktAxVar, aktVar);
                                        setDeviceStatusInMaininstance(aktVar, id);
                                    }
                                }
                                String subinstance = "V" + aktChannel.Key.ToString("D2");
                                if (aktInstanz.SubinstanceExists(subinstance))
                                {
                                    AxVariable aktAxVar2 = aktInstanz.GetSubinstance(subinstance).Get(aktVar.Name);
                                    if (aktAxVar2 != null)
                                    {
                                        if (aktVar.Type == VariableType.tAction)
                                        {
                                            aktAxVar2.Set(false);
                                            continue;
                                        }
                                        _varConverter.SetAxVariable(aktAxVar2, aktVar);
                                        setDeviceStatusInMaininstance(aktVar, id);
                                    }
                                }
                            }

                            foreach (KeyValuePair<String, ConfigParameter> configName in aktDevice.Channels[aktChannel.Key].Config)
                            {
                                var aktVar = configName.Value;
                                String aktVarName = aktVar.Name + "_C" + aktChannel.Key.ToString("D2");
                                if (aktInstanz.VariableExists(aktVarName))
                                {
                                    AxVariable aktAxVar = aktInstanz.Get(aktVarName);
                                    if (aktAxVar != null)
                                    {
                                        _varConverter.SetAxVariable(aktAxVar, aktVar);
                                    }
                                }
                                String subinstance = "C" + aktChannel.Key.ToString("D2");
                                if (aktInstanz.SubinstanceExists(subinstance))
                                {
                                    AxVariable aktAxVar2 = aktInstanz.GetSubinstance(subinstance).Get(aktVar.Name);
                                    if (aktAxVar2 != null)
                                    {
                                        _varConverter.SetAxVariable(aktAxVar2, aktVar);
                                    }
                                }
                            }
                        }
                    }
                    _homegearDevicesMutex.ReleaseMutex();
                    _initCompleted = true;
                    Logging.WriteLog(LogLevel.Info, _mainInstance, "Init Devices completed");
                }
                catch (Exception ex)
                {
                    Logging.WriteLog(LogLevel.Error, _mainInstance, ex.Message, ex.StackTrace);
                    _homegearDevicesMutex.ReleaseMutex();
                    _initCompleted = true;
                    Logging.WriteLog(LogLevel.Info,  _mainInstance, "Init Devices completed");
                }
            }
        }

        void OnSubinstanceVariableValueChanged(AxVariable sender)
        {
            if ((sender.Name == "Lifetick") || (sender.Name == "DataValid") || (sender.Name == "RSSI") || (sender.Name == "Status") || (sender.Name == "err") || (sender.Name == "LastChange"))
                return;

            if (_ignoredVariableChanged.Contains(sender.Path))  //Variablen die nicht in den Homegear-Devices vorkommen, werden einmal registriert und dann nie mehr ausgewertet
                return;

            _homegearDevicesMutex.WaitOne();
            if (sender == null || sender.Instance == null)
            {
                _homegearDevicesMutex.ReleaseMutex();
                return;
            }
            try
            {
                AxInstance parentInstance = sender.Instance.Parent;
                Boolean variableExists = false;
                if (_homegear.Devices.ContainsKey(parentInstance.Get("ID").GetLongInteger()))
                {
                    Logging.WriteLog(LogLevel.Debug, sender.Instance, "aX-Variable " + sender.Path + " has changed to: " + _varConverter.AutomationXVarToString(sender));
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
                                Logging.WriteLog(LogLevel.Debug, sender.Instance, "[aX -> HomeGear]: " + parentInstance.Name + "." + name + ", Channel:" + channelIndex.ToString() + " = " + _varConverter.AutomationXVarToString(sender));
                                SetLastChange(sender.Instance, "[aX -> HomeGear]: " + sender.Instance.Name + "." + name + ", Channel:" + channelIndex.ToString() + " = " + _varConverter.AutomationXVarToString(sender)); 
                                _varConverter.SetHomeGearVariable(channel.Variables[name], sender);
                                variableExists = true;
                                if (channel.Variables[name].Type == VariableType.tAction)  //Action Variablen sofort wieder rücksetzen
                                    sender.Set(false);
                            }
                        }
                        else if (type == "C")
                        {
                            if (channel.Config.ContainsKey(name))
                            {
                                Logging.WriteLog(LogLevel.Debug, sender.Instance, "[aX -> HomeGear]: " + parentInstance.Name + "." + name + ", Channel: " + channelIndex.ToString() + " = " + _varConverter.AutomationXVarToString(sender));
                                SetLastChange(sender.Instance, "[aX -> HomeGear]: " + sender.Instance.Name + "." + name + ", Channel: " + channelIndex.ToString() + " = " + _varConverter.AutomationXVarToString(sender));
                                AddConfigChannelChanged(parentInstance, channelIndex); 
                                _varConverter.SetHomeGearVariable(channel.Config[name], sender);
                                variableExists = true;
                            }
                        }
                    }
                }

                if (!variableExists)
                {
                    Logging.WriteLog(LogLevel.Info, sender.Instance, "Variable " + sender.Path + " ist nicht für Homegear relevont - wird ab jetzt ignoriert!");
                    _ignoredVariableChanged.Add(sender.Path);
                }
                _homegearDevicesMutex.ReleaseMutex();
            }
            catch (Exception ex)
            {
                try { _homegearDevicesMutex.ReleaseMutex(); } catch (Exception) { }
                Logging.WriteLog(LogLevel.Error, _mainInstance, ex.Message, ex.StackTrace);
            }
        }

        void AddConfigChannelChanged(AxInstance instance, Int32 channel)
        {
            String instancename = instance.Name;

            if (!_instancesConfigChannels.ContainsKey(instancename))
                _instancesConfigChannels.Add(instancename, new List<Int32>());

            if (!_instancesConfigChannels[instancename].Contains(channel))
                _instancesConfigChannels[instancename].Add(channel);

            if (instance.VariableExists("ConfigValuesChanged"))
                instance.Get("ConfigValuesChanged").Set(true);

            Logging.WriteLog(LogLevel.Debug, instance, "Adding ConfigChannel " + channel.ToString() + " from Instance " + instancename);
        }


        void OnInstanceVariableValueChanged(AxVariable sender)
        {
            if ((sender.Name == "Lifetick") || (sender.Name == "DataValid") || (sender.Name == "RSSI") || (sender.Name == "Status") || (sender.Name == "err") || (sender.Name == "LastChange"))
                return;

            if (_ignoredVariableChanged.Contains(sender.Path))  //Variablen die nicht in den Homegear-Devices vorkommen, werden einmal registriert und dann nie mehr ausgewertet
                return;

            lock (_instances._mutex)
            {
                try
                {
                    if (sender == null || sender.Instance == null)
                        return;

                    _homegearDevicesMutex.WaitOne();

                    Boolean variableExists = false;
                    if (_homegear.Devices.ContainsKey(sender.Instance.Get("ID").GetLongInteger()))
                    {
                        Logging.WriteLog(LogLevel.Debug, sender.Instance, "aX-Variable " + sender.Path + " has changed to: " + _varConverter.AutomationXVarToString(sender));
                        Device aktDevice = _homegear.Devices[sender.Instance.Get("ID").GetLongInteger()];
                        String name;
                        String type;
                        Int32 channelIndex;

                        if (sender.Name == "InterfaceID")
                        {
                            //aktDevice.Interface.ID = sender.GetString();
                            variableExists = true;
                        }
                        else if (sender.Name == "SetConfigValues")
                        {
                            variableExists = true;

                            if (!sender.GetBool())
                            {
                                _homegearDevicesMutex.ReleaseMutex();
                                return;
                            }

                            if (_instancesConfigChannels.ContainsKey(sender.Instance.Name))
                            {
                                //Muss Device, List<Int32> statt Device, int32 sein, da meherer Kanäle pro Gerät!!!!!
                                Dictionary<Device, List<Int32>> configToPush = new Dictionary<Device, List<Int32>>();
                                configToPush.Add(aktDevice, new List<Int32>());
                                foreach (Int32 aktChannel in _instancesConfigChannels[sender.Instance.Name])
                                {
                                    if (aktDevice.Channels.ContainsKey(aktChannel))
                                        configToPush[aktDevice].Add(aktChannel);
                                }
                                Dictionary<AxInstance, Dictionary<Device, List<Int32>>> queueAdd = new Dictionary<AxInstance, Dictionary<Device, List<Int32>>>();
                                queueAdd.Add(sender.Instance, configToPush);
                                _queueConfigToPushMutex.WaitOne();
                                _queueConfigToPush.Enqueue(queueAdd);
                                _queueConfigToPushMutex.ReleaseMutex();
                                _instancesConfigChannels.Remove(sender.Instance.Name);
                            }
                            if (sender.Instance.VariableExists("ConfigValuesChanged"))
                                sender.Instance.Get("ConfigValuesChanged").Set(false);
                        }
                        else if (sender.Name == "Name")
                        {
                            aktDevice.Name = sender.Instance.Get("Name").GetString();
                            variableExists = true;
                        }
                        else if (_varConverter.ParseAxVariable(sender, out name, out type, out channelIndex))
                        {
                            if (aktDevice.Channels.ContainsKey(channelIndex))
                            {
                                Channel channel = aktDevice.Channels[channelIndex];
                                if (type == "V")
                                {
                                    if (channel.Variables.ContainsKey(name))
                                    {
                                        Logging.WriteLog(LogLevel.Debug, sender.Instance, "[aX -> HomeGear]: " + sender.Instance.Name + "." + name + ", Channel:" + channelIndex.ToString() + " = " + _varConverter.AutomationXVarToString(sender));
                                        SetLastChange(sender.Instance, "[aX -> HomeGear]: " + sender.Instance.Name + "." + name + ", Channel:" + channelIndex.ToString() + " = " + _varConverter.AutomationXVarToString(sender));
                                        _varConverter.SetHomeGearVariable(channel.Variables[name], sender);
                                        variableExists = true;
                                        if (channel.Variables[name].Type == VariableType.tAction)  //Action Variablen sofort wieder rücksetzen
                                            sender.Set(false);
                                    }
                                }
                                else if (type == "C")
                                {
                                    if (channel.Config.ContainsKey(name))
                                    {
                                        Logging.WriteLog(LogLevel.Debug, sender.Instance, "[aX -> HomeGear]: " + sender.Instance.Name + "." + name + ", Channel: " + channelIndex.ToString() + " = " + _varConverter.AutomationXVarToString(sender));
                                        SetLastChange(sender.Instance, "[aX -> HomeGear]: " + sender.Instance.Name + "." + name + ", Channel: " + channelIndex.ToString() + " = " + _varConverter.AutomationXVarToString(sender));
                                        AddConfigChannelChanged(sender.Instance, channelIndex);
                                        _varConverter.SetHomeGearVariable(channel.Config[name], sender);
                                        variableExists = true;
                                    }
                                }
                            }
                        }
                    }
                    if (!variableExists)
                    {
                        Logging.WriteLog(LogLevel.Info, sender.Instance, "Variable " + sender.Path + " ist nicht für Homegear relevont - wird ab jetzt ignoriert!");
                        _ignoredVariableChanged.Add(sender.Path);
                    }
                    _homegearDevicesMutex.ReleaseMutex();
                }
                catch (Exception ex)
                {
                    try { _homegearDevicesMutex.ReleaseMutex(); }
                    catch (Exception) { }
                    Logging.WriteLog(LogLevel.Error, _mainInstance, ex.Message, ex.StackTrace);
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
                        Logging.WriteLog(LogLevel.Error, _mainInstance, ex.Message, ex.StackTrace);
                    }
                }
            }
        }

        void aX_ShuttingDown(Ax sender)
        {
            Dispose();
        }


        void _aX_SpsIdChanged(Ax sender)
        {
            try
            {
                _instances.Reload(_homegear.Devices);
            }
            catch (Exception ex)
            {
                Logging.WriteLog(LogLevel.Error, _mainInstance, ex.Message, ex.StackTrace);
            }
        }


        void Dispose()
        {
            try
            {
                if (_disposing) return;
                _disposing = true;

                Console.WriteLine("Aus, Ende!");

                Logging.WriteLog(LogLevel.Always, _mainInstance, "Beende RPC-Server...");
                _homegear.Dispose();
                _rpc.Dispose();

                _mainInstance.Get("err").Set(false);
                _mainInstance.Get("Init").Set(false);
                _mainInstance.Get("RPC_InitComplete").Set(false);
                _mainInstance.Get("ServiceMessageVorhanden").Set(false);
                _mainInstance.Get("PairingMode").Set(false);
                _mainInstance.Get("CAPI_Running").Set(false);
                Logging.WriteLog(LogLevel.Always, _mainInstance, "StaKoTecHomeGear.exe beendet");
                Console.WriteLine("Und aus!!");
            }
            catch (Exception ex)
            {
                Logging.WriteLog(LogLevel.Error, _mainInstance, ex.Message, ex.StackTrace);
            }
            finally
            {
                Environment.Exit(0);
            }
        }

        private void UpdateSystemVariables()
        {
            UInt16 x = 0;
            _homegearSystemVariablesMutex.WaitOne();
            _tempSystemvariableNames.Clear();
            try
            {
                foreach (KeyValuePair<String, SystemVariable> aktSystemVar in _homegear.SystemVariables)
                {
                    if (x > _systemVariableName.Length)
                        break;
                    _systemVariableName.Set(x, aktSystemVar.Key);
                    _systemVariableValue.Set(x, aktSystemVar.Value.ToString());
                    _tempSystemvariableNames.Add(x, aktSystemVar.Key.ToString());
                    x++;
                }
                for (; x < _systemVariableName.Length; x++)
                {
                    _systemVariableName.Set(x, "");
                    _systemVariableValue.Set(x, "");
                }
                _homegearSystemVariablesMutex.ReleaseMutex();
            }
            catch (Exception ex)
            {
                Logging.WriteLog(LogLevel.Error, _mainInstance, ex.Message, ex.StackTrace);
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

                if (sslEnable)
                {
                    SslInfo sslInfo = new SslInfo(
                        new Tuple<string, string>(sslHomeGearUsername, sslHomeGearPassword),
                        sslVerifyCertificate			//Enable hostname verification
                    );
                    _rpc = new RPCController(homegearHostName, homegearPort, sslInfo);
                }
                else
                    _rpc = new RPCController(homegearHostName, homegearPort);
                _rpc.AsciiDeviceTypeIdString = true;
                _rpc.Client.Connected += _rpc_ClientConnected;
                _rpc.Client.Disconnected += _rpc_ClientDisconnected;

                _homegear = new Homegear(_rpc, true);
                _homegear.ConnectError += _homegear_ConnectError;
                _homegear.DeviceConfigParameterUpdated += _homegear_DeviceConfigParameterUpdated;
                _homegear.DeviceLinkConfigParameterUpdated += _homegear_DeviceLinkConfigParameterUpdated;
                _homegear.DeviceReloadRequired += _homegear_DeviceReloadRequired;
                _homegear.DeviceVariableUpdated += _homegear_DeviceVariableUpdated;
                _homegear.EventUpdated += _homegear_EventUpdated;
                _homegear.HomegearError += _homegear_HomegearError;
                _homegear.MetadataUpdated += _homegear_MetadataUpdated;
                _homegear.Reloaded += _homegear_Reloaded;
                _homegear.ReloadRequired += _homegear_ReloadRequired;
                _homegear.SystemVariableUpdated += _homegear_SystemVariableUpdated;
            }
            catch (Exception ex)
            {
                Logging.WriteLog(LogLevel.Error, _mainInstance, ex.Message, ex.StackTrace);
            }
        }

        void _homegear_HomegearError(Homegear sender, long level, string message)
        {
            try
            {
                //Level: 
                //1: critical,
                //2: error,
                //3: warning
                if (level < 3)
                    aXAddHomegearError(message, (Int16)level);
                LogLevel loglevel = LogLevel.Error;
                if (level == 3)
                    loglevel = LogLevel.Warning;
                else if (level > 3)
                    loglevel = LogLevel.Info;
                Logging.WriteLog(loglevel, _mainInstance, "[HomeGear-Error-Handler]: " + message, "");
            }
            catch (Exception ex)
            {
                Logging.WriteLog(LogLevel.Error, _mainInstance, ex.Message, ex.StackTrace);
            }
        }

        void _homegear_EventUpdated(Homegear sender, Event homegearEvent)
        {
            try
            {
                Logging.WriteLog(LogLevel.Info, _mainInstance, "HomeGear Event " + homegearEvent.ToString() + " is updated");
            }
            catch (Exception ex)
            {
                Logging.WriteLog(LogLevel.Error, _mainInstance, ex.Message, ex.StackTrace);
            }
        }

        void _homegear_Reloaded(Homegear sender)
        {
            _reloading = false;
            try
            {
                _mainInstance.Get("RPC_InitComplete").Set(true);
                Logging.WriteLog(LogLevel.Info, _mainInstance, "RPC Init Complete");
            }
            catch (Exception ex)
            {
                Logging.WriteLog(LogLevel.Error, _mainInstance, ex.Message, ex.StackTrace);
            }
        }

        void _homegear_DeviceReloadRequired(Homegear sender, Device device, Channel channel, DeviceReloadType reloadType)
        {
            try
            {
                Logging.WriteLog(LogLevel.Info, _mainInstance, "RPC: Reload erforderlich (" + reloadType.ToString() + ")");
                _homegearDevicesMutex.WaitOne();
                if (reloadType == DeviceReloadType.Full)
                {
                    Logging.WriteLog(LogLevel.Info, _mainInstance, "Reloading device " + device.ID.ToString() + ".");
                    //Finish all operations on the device and then call:
                    device.Reload();
                }
                else if (reloadType == DeviceReloadType.Metadata)
                {
                    Logging.WriteLog(LogLevel.Info, _mainInstance, "Reloading metadata of device " + device.ID.ToString() + ".");
                    //Finish all operations on the device's metadata and then call:
                    device.Metadata.Reload();
                }
                else if (reloadType == DeviceReloadType.Channel)
                {
                    Logging.WriteLog(LogLevel.Info, _mainInstance, "Reloading channel " + channel.Index + " of device " + device.ID.ToString() + ".");
                    //Finish all operations on the device's channel and then call:
                    channel.Reload();
                }
                else if (reloadType == DeviceReloadType.Variables)
                {
                    Logging.WriteLog(LogLevel.Info, _mainInstance, "Device variables were updated: Device type: \"" + device.TypeString + "\", ID: " + device.ID.ToString() + ", Channel: " + channel.Index.ToString());
                    Logging.WriteLog(LogLevel.Info, _mainInstance, "Reloading variables of channel " + channel.Index + " and device " + device.ID.ToString() + ".");
                    //Finish all operations on the channels's variables and then call:
                    channel.Variables.Reload();
                }
                else if (reloadType == DeviceReloadType.Links)
                {
                    Logging.WriteLog(LogLevel.Info, _mainInstance, "Device links were updated: Device type: \"" + device.TypeString + "\", ID: " + device.ID.ToString() + ", Channel: " + channel.Index.ToString());
                    Logging.WriteLog(LogLevel.Info, _mainInstance, "Reloading links of channel " + channel.Index + " and device " + device.ID.ToString() + ".");
                    //Finish all operations on the channels's links and then call:
                    channel.Links.Reload();
                }
                else if (reloadType == DeviceReloadType.Team)
                {
                    Logging.WriteLog(LogLevel.Info, _mainInstance, "Device team was updated: Device type: \"" + device.TypeString + "\", ID: " + device.ID.ToString() + ", Channel: " + channel.Index.ToString());
                    Logging.WriteLog(LogLevel.Info, _mainInstance, "Reloading channel " + channel.Index + " of device " + device.ID.ToString() + ".");
                    //Finish all operations on the device's channel and then call:
                    channel.Reload();
                }
                else if (reloadType == DeviceReloadType.Events)
                {
                    Logging.WriteLog(LogLevel.Info, _mainInstance, "Device events were updated: Device type: \"" + device.TypeString + "\", ID: " + device.ID.ToString() + ", Channel: " + channel.Index.ToString());
                    Logging.WriteLog(LogLevel.Info, _mainInstance, "Reloading events of device " + device.ID.ToString() + ".");
                    //Finish all operations on the device's events and then call:
                    device.Events.Reload();
                }
                _homegearDevicesMutex.ReleaseMutex();
            }
            catch (Exception ex)
            {
                try { _homegearDevicesMutex.ReleaseMutex(); } catch (Exception) { }
                Logging.WriteLog(LogLevel.Error, _mainInstance, ex.Message, ex.StackTrace);
            }
        }

        void _homegear_ReloadRequired(Homegear sender, ReloadType reloadType)
        {
            try
            {
                Logging.WriteLog(LogLevel.Info, _mainInstance, "RPC: Reload erforderlich (" + reloadType.ToString() + ")");
                if (reloadType == ReloadType.Full)
                {
                    try
                    {
                        _mainInstance.Get("RPC_InitComplete").Set(false);
                        while (_reloading)
                        {
                            Logging.WriteLog(LogLevel.Info, _mainInstance, "Wait for homegear.Reload()");
                            Thread.Sleep(10);
                        }
                        Logging.WriteLog(LogLevel.Info, _mainInstance, "Homegear is full-reloading");
                        _homegearDevicesMutex.WaitOne();
                        _reloading = true;
                        _homegear.Reload();
                        _homegearDevicesMutex.ReleaseMutex();
                    }
                    catch (Exception ex)
                    {
                        Logging.WriteLog(LogLevel.Error, _mainInstance, "Reload Thread ist tot");
                        Logging.WriteLog(LogLevel.Error, _mainInstance, ex.Message, ex.StackTrace);
                    }
                }
                else if (reloadType == ReloadType.SystemVariables)
                {
                    Logging.WriteLog(LogLevel.Info, _mainInstance, "Homegear is reloading SystemVariables");
                    _homegearSystemVariablesMutex.WaitOne();
                    _homegear.SystemVariables.Reload();
                    _homegearSystemVariablesMutex.ReleaseMutex();
                    UpdateSystemVariables();
                }
                else if (reloadType == ReloadType.Events)
                {
                    Logging.WriteLog(LogLevel.Info, _mainInstance, "Homegear is reloading Events");
                    _homegear.TimedEvents.Reload();
                }
            }
            catch (Exception ex)
            {
                Logging.WriteLog(LogLevel.Error, _mainInstance, "Reload Thread ist tot");
                Logging.WriteLog(LogLevel.Error, _mainInstance, ex.Message, ex.StackTrace);
            }
        }

        void _homegear_DeviceLinkConfigParameterUpdated(Homegear sender, Device device, Channel channel, Link link, ConfigParameter parameter)
        {
            lock (_instances._mutex)
            {
                try
                {
                    Int32 id = (Int32)device.ID;
                    Logging.WriteLog(LogLevel.Debug, _mainInstance, "RPC: " + device.ID.ToString() + " " + link.RemotePeerID.ToString() + " " + link.RemoteChannel.ToString() + " " + parameter.Name + " = " + parameter.ToString());
                    if (_instances.ContainsKey(id))
                        Logging.WriteLog(LogLevel.Info, _instances[id], "Link-Parameter " + link.Name + " updated to " + link.ToString());
                }
                catch (Exception ex)
                {
                    Logging.WriteLog(LogLevel.Error, _mainInstance, ex.Message, ex.StackTrace);
                }
            }
        }

        void _homegear_DeviceConfigParameterUpdated(Homegear sender, Device device, Channel channel, ConfigParameter parameter)
        {
            lock (_instances._mutex)
            {
                try
                {
                    Int32 id = (Int32)device.ID;
                    Logging.WriteLog(LogLevel.Debug, _mainInstance, "RPC: " + device.ID.ToString() + " " + parameter.Name + " = " + parameter.ToString());
                    if (_instances.ContainsKey(id))
                        _instances[id]["Status"].Set("Config-Parameter " + parameter.Name + " updated to " + parameter.ToString());
                    Logging.WriteLog(LogLevel.Info, _mainInstance, "Config-Parameter " + parameter.Name + " updated to " + parameter.ToString());
                }
                catch (Exception ex)
                {
                    Logging.WriteLog(LogLevel.Error, _mainInstance, ex.Message, ex.StackTrace);
                }
            }
        }

        void setDeviceStatusInMaininstance(Variable variable, Int32 id)
        {
            if (!_errorStates.Contains(variable.Name) && !_warningStates.Contains(variable.Name))
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
                        if (!stateOld.Contains(", " + variable.Name))   //Nicht doppelt eintragen falls z.B. UNREACH im aX quittiert wurde und dann nochmal UNREACH aus dem Gerät ausgelesen wird
                            _deviceState.Set(x, stateOld + ", " + variable.Name);
                        if (_warningStates.Contains(variable.Name) && _deviceStateColor.GetInteger(x) != (Int16)DeviceStatus.Fehler)
                            _deviceStateColor.Set(x, (Int16)DeviceStatus.Warnung);
                        else
                            _deviceStateColor.Set(x, (Int16)DeviceStatus.Fehler);
                    }
                    else
                    {
                        _deviceState.Set(x, stateOld.Replace(", " + variable.Name, ""));
                    }

                    foreach (KeyValuePair<DeviceStatus, String> aktdeviceStatusText in _deviceStatusText)
                    {
                        if (aktdeviceStatusText.Value == _deviceState.GetString(x))  //Wenn der String wieder genau z.B. "OK" ist, dann en jeweiligen Status der Zeile dementsprechend setzen
                            _deviceStateColor.Set(x, (Int16)aktdeviceStatusText.Key);
                    }

                    break;
                }
            }
        }

        void _homegear_DeviceVariableUpdated(Homegear sender, Device device, Channel channel, Variable variable, string eventSource)
        {
            lock (_instances._mutex)
            {
                try
                {
                    Int32 deviceID = (Int32)device.ID;
                    String varName = variable.Name + "_V" + channel.Index.ToString("D2");
                    Logging.WriteLog(LogLevel.Debug, _mainInstance, "RPC: " + deviceID.ToString() + " " + variable.Name + " = " + variable.ToString());

                    if (variable.Name == "CURRENT_TRACK_METADATA")  //Sonos zerschiesst sonst das aX-Log
                        return;

                    if (_instances.ContainsKey(deviceID))
                    {
                        AxInstance instanz = _instances[deviceID];
                        String subinstance = "V" + channel.Index.ToString("D2");
                        if (instanz.VariableExists(varName))
                        {
                            AxVariable aktAxVar = instanz.Get(varName);
                            if (aktAxVar != null)
                            {
                                _varConverter.SetAxVariable(aktAxVar, variable);
                                setDeviceStatusInMaininstance(variable, deviceID);
                                Logging.WriteLog(LogLevel.Debug, aktAxVar.Instance, "[HomeGear -> aX]: " + aktAxVar.Path + " = " + variable.ToString());
                            }
                            SetLastChange(instanz, varName + " = " + variable.ToString());
                        }
                        else if (instanz.SubinstanceExists(subinstance))
                        {
                            AxVariable aktAxVar2 = instanz.GetSubinstance(subinstance).Get(variable.Name);
                            if (aktAxVar2 != null)
                            {
                                _varConverter.SetAxVariable(aktAxVar2, variable);
                                setDeviceStatusInMaininstance(variable, deviceID);
                                Logging.WriteLog(LogLevel.Debug, aktAxVar2.Instance, "[HomeGear -> aX]: " + aktAxVar2.Path + " = " + variable.ToString());
                            }
                            SetLastChange(instanz, subinstance + "." + variable.Name + " = " + variable.ToString());
                        }
                        else
                            SetLastChange(instanz, varName + " = " + variable.ToString() + " - Variable nicht vorhanden!");

                        //Console.WriteLine(device.SerialNumber + ": " + variable.Name + ": " + variable.ToString());
                    }

                    // Variable in _mainInstnace aktualisieren
                    if (deviceID == _deviceVars_DeviceID)
                    {
                        for (UInt16 x = 0; x < _deviceVars_Name.Length; x++)
                        {
                            if (_deviceVars_Name.GetString(x) == varName)
                            {
                                _deviceVars_Actual.Set(x, variable.ToString());
                                break;
                            }
                        }
                    }
                }
                catch (Exception ex)
                {
                    Logging.WriteLog(LogLevel.Error, _mainInstance, ex.Message, ex.StackTrace);
                }
            }
        }

        void SetLastChange(AxInstance instanz, String text)
        {
            try
            {
                if (!instanz.VariableExists("LastChange"))
                    return;

                AxVariable aXVariable_LastChange = instanz.Get("LastChange");
                List<String> lastChange = new List<String>();
                UInt16 x = 0;

                lastChange.Add(DateTime.Now.Hour.ToString("D2") + ":" + DateTime.Now.Minute.ToString("D2") + ":" + DateTime.Now.Second.ToString("D2") + ": " + text);
                for (x = 0; x < aXVariable_LastChange.Length; x++)
                    lastChange.Add(aXVariable_LastChange.GetString(x));
                for (x = 0; x < aXVariable_LastChange.Length; x++)
                    aXVariable_LastChange.Set(x, lastChange[x]);

                Logging.WriteLog(LogLevel.Error, instanz, text);
            }
            catch (Exception ex)
            {
                Logging.WriteLog(LogLevel.Error, _mainInstance, ex.Message, ex.StackTrace);
            }
        }


        void _homegear_MetadataUpdated(Homegear sender, Device device, MetadataVariable variable)
        {
            try
            {
                Logging.WriteLog(LogLevel.Info, _mainInstance, "RPC: Metadata Variable '" + variable.Name + "' geändert");
            }
            catch (Exception ex)
            {
                Logging.WriteLog(LogLevel.Error, _mainInstance, ex.Message, ex.StackTrace);
            }
        }

        void _homegear_SystemVariableUpdated(Homegear sender, SystemVariable variable)
        {
            try
            {
                Logging.WriteLog(LogLevel.Info, _mainInstance, "RPC: System Variable '" + variable.Name + "' geändert");
                UpdateSystemVariables();
            }
            catch (Exception ex)
            {
                Logging.WriteLog(LogLevel.Error, _mainInstance, ex.Message, ex.StackTrace);
            }
        }

        void _homegear_ConnectError(Homegear sender, string message, string stackTrace)
        {
            try
            {
                Logging.WriteLog(LogLevel.Error, _mainInstance, message);
            }
            catch (Exception ex)
            {
                Logging.WriteLog(LogLevel.Error, _mainInstance, ex.Message, ex.StackTrace);
            }
        }

        void _rpc_ClientDisconnected(RPCClient sender)
        {
            try
            {
                Logging.WriteLog(LogLevel.Warning, _mainInstance, "RPC-Client Verbindung unterbrochen");
            }
            catch (Exception ex)
            {
                Logging.WriteLog(LogLevel.Error, _mainInstance, ex.Message, ex.StackTrace);
            }
        }

        void _rpc_ClientConnected(RPCClient sender, CipherAlgorithmType cipherAlgorithm, Int32 cipherStrength)
        {
            try
            {
                Logging.WriteLog(LogLevel.Info, _mainInstance, "RPC-Client verbunden");
                _mainInstance.Get("err").Set(false);
                if (!_reloading && !_mainInstance.Get("RPC_InitComplete").GetBool())
                {
                    _reloading = true;
                    _homegear.Reload();
                }
            }
            catch (Exception ex)
            {
                Logging.WriteLog(LogLevel.Error, _mainInstance, ex.Message, ex.StackTrace);
            }
        }


        void IDisposable.Dispose()
        {
            Dispose();
        }
    }
}
