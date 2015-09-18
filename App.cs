﻿using HomegearLib;
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

        bool _disposing = false;
        bool _initCompleted = false;
        AX _aX = null;
        AXInstance _mainInstance = null;
        LogLevel _logLevel = LogLevel.Error;

        VariableConverter _varConverter = null;

        AXVariable _homegearInterfaces = null;
        AXVariable _deviceID = null;
        AXVariable _deviceInstance = null;
        AXVariable _deviceRemark = null;
        AXVariable _deviceTypeString = null;
        AXVariable _deviceState = null;
        AXVariable _deviceFirmware = null;
        AXVariable _deviceInterface = null;
        AXVariable _deviceStateColor = null;

        AXVariable _systemVariableName = null;
        AXVariable _systemVariableValue = null;

        HomegearLib.RPC.RPCController _rpc = null;
        HomegearLib.Homegear _homegear = null;
        Boolean _reloading = false;


        Mutex _queueConfigToPushMutex = new Mutex();
        Queue<Dictionary<AXInstance, Dictionary<Device, List<Int32>>>> _queueConfigToPush = new Queue<Dictionary<AXInstance, Dictionary<Device, List<Int32>>>>();
        Dictionary<AXInstance, Dictionary<Device, List<Int32>>> _aktQueue = null;
        Dictionary<String, List<Int32>> _instancesConfigChannels = null;
        Thread _pushConfigThread = null;
        
        Instances _instances = null;
        Mutex _homegearDevicesMutex = new Mutex();

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
                ////////////////////////////////////////////////////////////////////////////////////////////
                // Status Texte laden
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

                _firstInitDevices = new List<Int32>();

                _aX.ShuttingDown += aX_ShuttingDown;
                _aX.SpsIdChanged += _aX_SpsIdChanged;

                _mainInstance = new AXInstance(_aX, instanceName, "Status", "err"); //Instanz-Objekt erstellen
                _mainInstance.PollingInterval = 200;
                _mainInstance.SetVariableEvents(true);
                Logging.Init(_aX, _mainInstance);
                _varConverter = new VariableConverter(_mainInstance);
                _instances = new Instances(_aX, _mainInstance);
                _instances.VariableValueChanged += OnInstanceVariableValueChanged;
                _instances.SubinstanceVariableValueChanged += OnSubinstanceVariableValueChanged;

                _instancesConfigChannels = new Dictionary<String, List<Int32>>();

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
                AXVariable changeInterface = _mainInstance.Get("ChangeInterface");
                changeInterface.Set(false);
                changeInterface.ValueChanged += changeInterface_ValueChanged;
                AXVariable getDeviceVars = _mainInstance.Get("GetDeviceVars");
                getDeviceVars.Set(false);
                getDeviceVars.ValueChanged += getDeviceVars_ValueChanged;
                AXVariable getDeviceConfigVars = _mainInstance.Get("GetDeviceConfigVars");
                getDeviceConfigVars.Set(false);
                getDeviceConfigVars.ValueChanged += getDeviceVars_ValueChanged;
                AXVariable homegearErrorQuit = _mainInstance.Get("HomegearErrorQuit");
                UInt16 x = 0;
                for (x = 0; x < homegearErrorQuit.Length; x++)
                    homegearErrorQuit.Set(x, false);
                homegearErrorQuit.ArrayValueChanged += homegearErrorQuit_ArrayValueChanged;
                AXVariable homegearErrorState = _mainInstance.Get("HomegearErrorState");
                Int16 errorCount = 0;
                for (x = 0; x < homegearErrorState.Length; x++)
                {
                    if (homegearErrorState.GetLongInteger(x) > 0)
                        errorCount++;
                }
                _mainInstance.Get("HomegearError").Set((errorCount > 0));


                Int32 axStartID = _mainInstance.Get("StartID").GetInteger();
                Int32 axStartID_old = axStartID;

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

                HomeGearConnect();

                UInt32 j = 0;
                Stopwatch stopWatch = new Stopwatch();
                Stopwatch stopWatchThreadTimer = new Stopwatch();
                Double lastCycletime = 0;
                Double cycletimerServiceMessages = 0;
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
                            Logging.WriteLog(LogLevel.Info, "StartID has changed! Exiting!!!");
                            Dispose();
                            continue;
                        }

                        _logLevel = (LogLevel)_mainInstance.Get("LogLevel").GetLongInteger();
                        //Zu übertragende Config-Parameter abarbeiten
                        //Logging.WriteLog(cycleCounter.ToString() + " start");
                        if ((_queueConfigToPush.Count > 0) && _initCompleted && (_pushConfigThread == null || !_pushConfigThread.IsAlive || _pushConfigThread.ThreadState == System.Threading.ThreadState.Aborted))
                        {
                            //Logging.WriteLog(cycleCounter.ToString() + " geht los");
                            _pushConfigThread = new Thread(PushConfig);
                            _pushConfigThread.Start();
                            
                        }
                        //Logging.WriteLog(cycleCounter.ToString() + " feddich");

                        if (!_rpc.IsConnected)
                        {
                            _mainInstance.Get("RPC_InitComplete").Set(false);
                            _initCompleted = false;

                            _firstInitDevices.Clear();

                            if (connectionTimeout > 0)
                                Logging.WriteLog(LogLevel.Info, "Waiting for RPC-Server connection... (" + (connectionTimeout * 5).ToString() + " s)");

                            if ((connectionTimeout > 6) && !mainInstanceErr.GetBool())
                            {
                                Logging.WriteLog(LogLevel.Info, "Waiting for RPC-Server connection... (" + (connectionTimeout * 5).ToString() + " s)", "", true);
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
                            AXVariable aX_serviceMessages = _mainInstance.Get("ServiceMessages");
                            foreach (ServiceMessage message in serviceMessages)
                            {
                                if (_errorStates.Contains(message.Type))
                                    alarmCounter++;
                                if (_warningStates.Contains(message.Type))
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
                            foreach(KeyValuePair<String, Interface> aktInterface in _rpc.ListInterfaces())
                            {
                                Logging.WriteLog(LogLevel.Info, "Interface " + aktInterface.Value.ID + ": LastPacketReceived: " + (currentTime - aktInterface.Value.LastPacketReceived).ToString() + "s ago");
                                Logging.WriteLog(LogLevel.Info, "Interface " + aktInterface.Value.ID + ": LastPacketSent: " + (currentTime - aktInterface.Value.LastPacketSent).ToString() + "s ago");
                            }

                            //Firmwareupgrades prüfen
                            /*_homegearDevicesMutex.WaitOne();
                            _instances.MutexLocked = true;
                            foreach (KeyValuePair<Int32, AXInstance> aktInstance in _instances)
                            {
                                deviceCheckFirmwareUpdates(aktInstance.Key);
                            }
                            _homegearDevicesMutex.ReleaseMutex();
                            _instances.MutexLocked = false;*/
                        }
                        //Logging.WriteLog("Cycle-Dauer: " + (lastCycletime).ToString() + "s");
                        j++;
                    }
                    catch (Exception ex)
                    {
                        Logging.WriteLog(LogLevel.Error, ex.Message, ex.StackTrace);
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
                Logging.WriteLog(LogLevel.Error, ex.Message, ex.StackTrace);
            }
        }

        void changeInterface_ValueChanged(AXVariable sender)
        {
            try
            {
                Int32 deviceID = _mainInstance.Get("ChangeInterfaceDeviceID").GetLongInteger();
                String deviceInterface = _mainInstance.Get("ChangeInterfaceName").GetString();

                Logging.WriteLog(LogLevel.Debug, "Change interface from DeciveID " + deviceID.ToString() + " to: " + deviceInterface);
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
                Logging.WriteLog(LogLevel.Debug, "Change interface ready");
            }
            catch (Exception ex)
            {
                Logging.WriteLog(LogLevel.Error, ex.Message, ex.StackTrace);
            }
        }

        Double getUnixTime()
        {
            Double currentTime = (DateTime.UtcNow.Subtract(new DateTime(1970, 1, 1))).TotalMilliseconds;
            return currentTime;
        }
        

        void PushConfig()
        {
            try
            {
                Logging.WriteLog(LogLevel.Info, "Pushing Config-Thread gestartet");
                while (_queueConfigToPush.Count > 0) 
                {
                    _queueConfigToPushMutex.WaitOne();
                    _aktQueue = _queueConfigToPush.Dequeue();
                    _queueConfigToPushMutex.ReleaseMutex();
                    foreach (KeyValuePair<AXInstance, Dictionary<Device, List<Int32>>> aktInstance in _aktQueue)
                    {
                        foreach (KeyValuePair<Device, List<Int32>> aktDevice in aktInstance.Value)
                        {
                            foreach (Int32 aktChannel in aktDevice.Value)
                            {
                                Logging.WriteLog(LogLevel.Info, "[" + aktInstance.Key.Path + "] Pushe Config für Kanal " + aktChannel.ToString());
                                aktInstance.Key.Status = "Pushe Config für Kanal " + aktChannel.ToString();
                                aktDevice.Key.Channels[aktChannel].Config.Put();
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
                Logging.WriteLog(LogLevel.Info, "PushConfig Thread abgebrochen. Geht gleich weiter...");
            }
            catch (Exception ex)
            {
                Logging.WriteLog(LogLevel.Error, ex.Message, ex.StackTrace);
            }
        }

        String deviceCheckFirmwareUpdates(Int32 deviceID)
        {
            String firmware = "";
            if ((_homegear.Devices.ContainsKey(deviceID)) && (_instances.ContainsKey(deviceID)))
            {
                Device aktDevice = _homegear.Devices[deviceID];
                AXInstance aktInstance = _instances[deviceID];

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

        void homegearErrorQuit_ArrayValueChanged(AXVariable sender, ushort index)
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
                AXVariable homegearErrorState = _mainInstance.Get("HomegearErrorState");
                AXVariable homegearErrors = _mainInstance.Get("HomegearErrors");
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
                Logging.WriteLog(LogLevel.Error, ex.Message, ex.StackTrace);
            }
        }

        void aXAddHomegearError(String message, Int16 level)
        {
            try
            {
                List<String> homegearErrorsTemp = new List<String>();
                List<Int32> homegearErrorStateTemp = new List<Int32>();
                AXVariable homegearErrorState = _mainInstance.Get("HomegearErrorState");
                AXVariable homegearErrors = _mainInstance.Get("HomegearErrors");
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
                Logging.WriteLog(LogLevel.Error, ex.Message, ex.StackTrace);
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
                Logging.WriteLog(LogLevel.Error, ex.Message, ex.StackTrace);
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
                AXVariable deviceVars_Actual = _mainInstance.Get("DeviceVars_Actual");
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
                                    Logging.WriteLog(LogLevel.Error, "Array-Index zu klein bei 'DeviceVars_Name'");
                                    _homegearDevicesMutex.ReleaseMutex();
                                    _instances.MutexLocked = false;
                                    return;
                                }

                                var aktVar = configName.Value;
                                String minVar = "";
                                String maxVar = "";
                                String typ = "";
                                String defaultVar = "";
                                String actualVar = "";
                                String rwVar = "";
                                String varOK = findVarInClass(deviceID, aktVar.Name, aktVar.Type, "C", aktChannel.Key.ToString("D2"));

                                _varConverter.ParseDeviceConfigVars(aktVar, out minVar, out maxVar, out typ, out defaultVar, out rwVar);

                                actualVar = _varConverter.HomegearVarToString(aktVar);
                                deviceVars_Name.Set(x, aktVar.Name + "_C" + aktChannel.Key.ToString("D2"));
                                deviceVars_Type.Set(x, typ);
                                deviceVars_Min.Set(x, minVar);
                                deviceVars_Max.Set(x, maxVar);
                                deviceVars_Default.Set(x, defaultVar);
                                deviceVars_Actual.Set(x, actualVar);
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
                                    Logging.WriteLog(LogLevel.Error, "Array-Index zu klein bei 'DeviceVars_Name'");
                                    _homegearDevicesMutex.ReleaseMutex();
                                    _instances.MutexLocked = false;
                                    return;
                                }

                                var aktVar = varName.Value;
                                String minVar = "";
                                String maxVar = "";
                                String typ = "";
                                String defaultVar = "";
                                String actualVar = "";
                                String rwVar = "";
                                String varOK = findVarInClass(deviceID, aktVar.Name, aktVar.Type, "V", aktChannel.Key.ToString("D2"));
                                _varConverter.ParseDeviceVars(aktVar, out minVar, out maxVar, out typ, out defaultVar, out rwVar);

                                actualVar = _varConverter.HomegearVarToString(aktVar);
                                deviceVars_Name.Set(x, aktVar.Name + "_V" + aktChannel.Key.ToString("D2"));
                                deviceVars_Type.Set(x, typ);
                                deviceVars_Min.Set(x, minVar);
                                deviceVars_Max.Set(x, maxVar);
                                deviceVars_Default.Set(x, defaultVar);
                                deviceVars_Actual.Set(x, actualVar);
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
                        deviceVars_Actual.Set(x, "");
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
                Logging.WriteLog(LogLevel.Error, ex.Message, ex.StackTrace);
            }
            finally
            {
                try
                {
                    sender.Set(false);
                }
                catch (Exception ex)
                {
                    Logging.WriteLog(LogLevel.Error, ex.Message, ex.StackTrace);
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
                Logging.WriteLog(LogLevel.Error, ex.Message, ex.StackTrace);
            }
        }

        void deviceReset_ValueChanged(AXVariable sender)
        {
            try
            {
                Int32 deviceID = _mainInstance.Get("ActionID").GetLongInteger();
                sender.Instance.Status = "Resetting Device ID " + deviceID.ToString();
                if (_firstInitDevices.Contains(deviceID))
                    _firstInitDevices.Remove(deviceID);

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
                Logging.WriteLog(LogLevel.Error, ex.Message, ex.StackTrace);
            }
        }

        void deviceUnpair_ValueChanged(AXVariable sender)
        {
            try
            {
                Int32 deviceID = _mainInstance.Get("ActionID").GetLongInteger();
                sender.Instance.Status = "Unpairing Device ID " + deviceID.ToString();
                if (_firstInitDevices.Contains(deviceID))
                    _firstInitDevices.Remove(deviceID);

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
                Logging.WriteLog(LogLevel.Error, ex.Message, ex.StackTrace);
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
                Logging.WriteLog(LogLevel.Error, ex.Message, ex.StackTrace);
            }
        }

        void start_CAPI_Release_ValueChanged(AXVariable sender)
        {
            try
            {
                if (!sender.GetBool())
                {
                    _mainInstance.Get("err").Set(false);
                    Logging.WriteLog(LogLevel.Always, "Beende StaKoTecHomeGear.exe");

                    Dispose();
                }
            }
            catch (Exception ex)
            {
                sender.Instance.Error = ex.Message;
                sender.Instance.Status = ex.Message;
                Logging.WriteLog(LogLevel.Error, ex.Message, ex.StackTrace);
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
                Logging.WriteLog(LogLevel.Error, ex.Message, ex.StackTrace);
            }
            return className;
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

        void init_ValueChanged(AXVariable sender)
        {
            try
            {
                if (!sender.GetBool())
                    return;
                
                _initCompleted = false;
                if (_pushConfigThread != null && _pushConfigThread.IsAlive)
                    _pushConfigThread.Abort();


                UInt16 x = 0;
                Logging.WriteLog(LogLevel.Info, "Init Devices");

                _homegearInterfaces = sender.Instance.Get("HomegearInterfaces");
                _deviceID = sender.Instance.Get("DeviceID");
                _deviceInstance = sender.Instance.Get("DeviceInstance");
                _deviceRemark = sender.Instance.Get("DeviceRemark");
                _deviceTypeString = sender.Instance.Get("DeviceTypeString");
                _deviceState = sender.Instance.Get("DeviceState");
                _deviceFirmware = sender.Instance.Get("DeviceFirmware");
                _deviceInterface = sender.Instance.Get("DeviceInterface");
                _deviceStateColor = sender.Instance.Get("DeviceStateColor");

                _systemVariableName = sender.Instance.Get("SystemVariableName");
                _systemVariableValue = sender.Instance.Get("SystemVariableValue");
                UpdateSystemVariables();

                _mainInstance.Get("HomeGearVersion").Set(_homegear.Version);
                getInterfaces();

                _homegearDevicesMutex.WaitOne();
                Logging.WriteLog(LogLevel.Info, "Reloading Instances");
                _instances.Reload(_homegear.Devices);
                _instances.MutexLocked = true;
                try
                {
                    x = 0;
                    //Erstmal alle ColorStates löschen, damit sichtbar wird welches Device bereite geparst wurde
                    for (x = 0; x < _deviceStateColor.Length; x++ )
                        _deviceStateColor.Set(x, (Int16)DeviceStatus.Nichts);

                    x = 0;
                    foreach (KeyValuePair<Int32, Device> devicePair in _homegear.Devices)
                    {
                        if (devicePair.Key > 1000000000)  //Teams nicht anzeigen
                            continue;

                        _deviceID.Set(x, devicePair.Key);

                        Logging.WriteLog(LogLevel.Debug, "Parse DeviceID " + devicePair.Key.ToString());
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

                                _deviceFirmware.Set(x, deviceCheckFirmwareUpdates(devicePair.Key));
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
                                if (!_firstInitDevices.Contains(devicePair.Key))
                                {
                                    Logging.WriteLog(LogLevel.Debug, "getActualDeviceData for Device ID " + devicePair.Key.ToString());
                                    getActualDeviceData(devicePair.Value, aktInstanz);
                                    _firstInitDevices.Add(devicePair.Key);
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
                                _instances.Remove(devicePair.Key, false);
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
                    Logging.WriteLog(LogLevel.Error, ex.Message, ex.StackTrace);
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
                Logging.WriteLog(LogLevel.Info, "Init Devices completed");
                _homegearDevicesMutex.ReleaseMutex();
            }
            catch (Exception ex)
            {
                try { _homegearDevicesMutex.ReleaseMutex(); }  catch (Exception) { }
                sender.Instance.Error = ex.Message;
                Logging.WriteLog(LogLevel.Error, ex.Message, ex.StackTrace);
            }
            finally
            {
                try
                {
                    sender.Set(false);
                    Logging.WriteLog(LogLevel.Info, "Init completely finished");
                    _mainInstance.Get("PolledVariables").Set(_mainInstance.PolledVariablesCount + _instances.PolledVariablesCount);
                }
                catch (Exception ex)
                {
                    Logging.WriteLog(LogLevel.Error, ex.Message, ex.StackTrace);
                }
            }
        }

        void getActualDeviceData(Device aktDevice, AXInstance aktInstanz)
        {
            foreach (KeyValuePair<Int32, Channel> aktChannel in aktDevice.Channels)
            {
                foreach (KeyValuePair<String, Variable> Wert in aktDevice.Channels[aktChannel.Key].Variables)
                {
                    var aktVar = Wert.Value;
                    if (aktVar.Type == VariableType.tAction)  //Action-Variablen nicht beim Init auslesen (Wie z.B. PRESS_SHORT oder so), da HomeGear speichert, dass die Variable irgendwann mal auf 1 war und somit immer beim warmstart alle bisher gedrückten Taster noch einmal auf 1 gesetzt werden
                        continue;
                    String aktVarName = aktVar.Name + "_V" + aktChannel.Key.ToString("D2");
                    if (aktInstanz.VariableExists(aktVarName))
                    {
                        AXVariable aktAXVar = aktInstanz.Get(aktVarName);
                        if (aktAXVar != null)
                        {
                            _varConverter.SetAXVariable(aktAXVar, aktVar);
                            setDeviceStatusInMaininstance(aktVar, aktDevice.ID);
                        }
                    }
                    String subinstance = "V" + aktChannel.Key.ToString("D2");
                    if (aktInstanz.SubinstanceExists(subinstance))
                    {
                        AXVariable aktAXVar2 = aktInstanz.GetSubinstance(subinstance).Get(aktVar.Name);
                        if (aktAXVar2 != null)
                        {
                            _varConverter.SetAXVariable(aktAXVar2, aktVar);
                            setDeviceStatusInMaininstance(aktVar, aktDevice.ID);
                        }
                    }
                }
                
                foreach (KeyValuePair<String, ConfigParameter> configName in aktDevice.Channels[aktChannel.Key].Config)
                {
                    var aktVar = configName.Value;
                    String aktVarName = aktVar.Name + "_C" + aktChannel.Key.ToString("D2");
                    if (aktInstanz.VariableExists(aktVarName))
                    {
                        AXVariable aktAXVar = aktInstanz.Get(aktVarName);
                        if (aktAXVar != null)
                        {
                            _varConverter.SetAXVariable(aktAXVar, aktVar);
                        }
                    }
                    String subinstance = "C" + aktChannel.Key.ToString("D2");
                    if (aktInstanz.SubinstanceExists(subinstance))
                    {
                        AXVariable aktAXVar2 = aktInstanz.GetSubinstance(subinstance).Get(aktVar.Name);
                        if (aktAXVar2 != null)
                        {
                            _varConverter.SetAXVariable(aktAXVar2, aktVar);
                        }
                    }
                }
            }
        }

        void OnSubinstanceVariableValueChanged(AXVariable sender)
        {
            if ((sender.Name == "Lifetick") || (sender.Name == "DataValid") || (sender.Name == "RSSI") || (sender.Name == "Status") || (sender.Name == "err") || (sender.Name == "LastChange"))
                return;

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
                    Logging.WriteLog(LogLevel.Debug, "aX-Variable " + sender.Path + " has changed to: " + _varConverter.AutomationXVarToString(sender));
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
                                Logging.WriteLog(LogLevel.Debug, "[aX -> HomeGear]: " + parentInstance.Name + "." + name + ", Channel:" + channelIndex.ToString() + " = " + _varConverter.AutomationXVarToString(sender));
                                SetLastChange(sender.Instance, "[aX -> HomeGear]: " + sender.Instance.Name + "." + name + ", Channel:" + channelIndex.ToString() + " = " + _varConverter.AutomationXVarToString(sender)); 
                                _varConverter.SetHomeGearVariable(channel.Variables[name], sender);
                            }
                        }
                        else if (type == "C")
                        {
                            if (channel.Config.ContainsKey(name))
                            {
                                Logging.WriteLog(LogLevel.Debug, "[aX -> HomeGear]: " + parentInstance.Name + "." + name + ", Channel: " + channelIndex.ToString() + " = " + _varConverter.AutomationXVarToString(sender));
                                SetLastChange(sender.Instance, "[aX -> HomeGear]: " + sender.Instance.Name + "." + name + ", Channel: " + channelIndex.ToString() + " = " + _varConverter.AutomationXVarToString(sender));
                                AddConfigChannelChanged(parentInstance, channelIndex); 
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
                Logging.WriteLog(LogLevel.Error, ex.Message, ex.StackTrace);
            }
        }

        void AddConfigChannelChanged(AXInstance instance, Int32 channel)
        {
            String instancename = instance.Name;

            if (!_instancesConfigChannels.ContainsKey(instancename))
                _instancesConfigChannels.Add(instancename, new List<Int32>());

            if (!_instancesConfigChannels[instancename].Contains(channel))
                _instancesConfigChannels[instancename].Add(channel);

            if (instance.VariableExists("ConfigValuesChanged"))
                instance.Get("ConfigValuesChanged").Set(true);
        }


        void OnInstanceVariableValueChanged(AXVariable sender)
        {
            if ((sender.Name == "Lifetick") || (sender.Name == "DataValid") || (sender.Name == "RSSI") || (sender.Name == "Status") || (sender.Name == "err") || (sender.Name == "LastChange"))
                return;

            try
            {
                if (sender == null || sender.Instance == null)
                    return;
                
                _homegearDevicesMutex.WaitOne();
                _instances.MutexLocked = true;
                if(_homegear.Devices.ContainsKey(sender.Instance.Get("ID").GetLongInteger()))
                {
                    Logging.WriteLog(LogLevel.Debug, "aX-Variable " + sender.Path + " has changed to: " + _varConverter.AutomationXVarToString(sender));
                    Device aktDevice = _homegear.Devices[sender.Instance.Get("ID").GetLongInteger()];
                    String name;
                    String type;
                    Int32 channelIndex;

                    if (sender.Name == "InterfaceID")
                    {
                        //aktDevice.Interface.ID = sender.GetString();
                    }
                    else if (sender.Name == "SetConfigValues")
                    {
                        if (!sender.GetBool())
                        {
                            _homegearDevicesMutex.ReleaseMutex();
                            _instances.MutexLocked = false;
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
                            Dictionary<AXInstance, Dictionary<Device, List<Int32>>> queueAdd = new Dictionary<AXInstance, Dictionary<Device, List<Int32>>>();
                            queueAdd.Add(sender.Instance, configToPush);
                            _queueConfigToPushMutex.WaitOne();
                            _queueConfigToPush.Enqueue(queueAdd);
                            _queueConfigToPushMutex.ReleaseMutex();
                            _instancesConfigChannels.Remove(sender.Instance.Name);
                        }
                        if (sender.Instance.VariableExists("ConfigValuesChanged"))
                            sender.Instance.Get("ConfigValuesChanged").Set(false);
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
                                    Logging.WriteLog(LogLevel.Debug, "[aX -> HomeGear]: " + sender.Instance.Name + "." + name + ", Channel:" + channelIndex.ToString() + " = " + _varConverter.AutomationXVarToString(sender));
                                    SetLastChange(sender.Instance, "[aX -> HomeGear]: " + sender.Instance.Name + "." + name + ", Channel:" + channelIndex.ToString() + " = " + _varConverter.AutomationXVarToString(sender));
                                    _varConverter.SetHomeGearVariable(channel.Variables[name], sender);
                                }
                            }
                            else if (type == "C")
                            {
                                if (channel.Config.ContainsKey(name))
                                {
                                    Logging.WriteLog(LogLevel.Debug, "[aX -> HomeGear]: " + sender.Instance.Name + "." + name + ", Channel: " + channelIndex.ToString() + " = " + _varConverter.AutomationXVarToString(sender));
                                    SetLastChange(sender.Instance, "[aX -> HomeGear]: " + sender.Instance.Name + "." + name + ", Channel: " + channelIndex.ToString() + " = " + _varConverter.AutomationXVarToString(sender));
                                    AddConfigChannelChanged(sender.Instance, channelIndex);
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
                _instances.MutexLocked = false;
            }
            catch (Exception ex)
            {
                try { _homegearDevicesMutex.ReleaseMutex(); } catch (Exception) { }
                Logging.WriteLog(LogLevel.Error, ex.Message, ex.StackTrace);
                _instances.MutexLocked = false;
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
                    Logging.WriteLog(LogLevel.Error, ex.Message, ex.StackTrace);
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

                Logging.WriteLog(LogLevel.Info, "SPS-ID Changed! Triggering Init!");
                AXVariable init = _mainInstance.Get("Init");
                init.Set(true);
                init_ValueChanged(init);
                System.Diagnostics.Debug.WriteLine("SPS-ID change end");
            }
            catch (Exception ex)
            {
                Logging.WriteLog(LogLevel.Error, ex.Message, ex.StackTrace);
            }
        }


        void Dispose()
        {
            try
            {
                if (_disposing) return;
                _disposing = true;

                Console.WriteLine("Aus, Ende!");

                Logging.WriteLog(LogLevel.Always, "Beende RPC-Server...");
                _homegear.Dispose();
                _rpc.Dispose();

                _mainInstance.Get("err").Set(false);
                _mainInstance.Get("Init").Set(false);
                _mainInstance.Get("RPC_InitComplete").Set(false);
                _mainInstance.Get("ServiceMessageVorhanden").Set(false);
                _mainInstance.Get("PairingMode").Set(false);
                _mainInstance.Get("CAPI_Running").Set(false);
                Logging.WriteLog(LogLevel.Always, "StaKoTecHomeGear.exe beendet");
                Console.WriteLine("Und aus!!");
            }
            catch (Exception ex)
            {
                Logging.WriteLog(LogLevel.Error, ex.Message, ex.StackTrace);
            }
            finally
            {
                Environment.Exit(0);
            }
        }

        private void UpdateSystemVariables()
        {
            UInt16 x = 0;
            foreach (KeyValuePair<String, SystemVariable> aktSystemVar in _homegear.SystemVariables)
            {
                if (x > _systemVariableName.Length)
                    break;
                _systemVariableName.Set(x, aktSystemVar.Key);
                _systemVariableValue.Set(x, aktSystemVar.Value.ToString());
                x++;
            }
            for(; x < _systemVariableName.Length; x++)
            {
                _systemVariableName.Set(x, "");
                _systemVariableValue.Set(x, "");
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
                Logging.WriteLog(LogLevel.Error, ex.Message, ex.StackTrace);
            }
        }

        void _homegear_HomegearError(Homegear sender, int level, string message)
        {
            try
            {
                //Level: 
                //1: critical,
                //2: error,
                //3: warning
                if (level < 3)
                    aXAddHomegearError(message, (Int16)level);
                Logging.WriteLog(LogLevel.Error, "[HomeGear-Error-Handler] (Level: " + level.ToString() + ") " + message, "", (level < 3));
            }
            catch (Exception ex)
            {
                Logging.WriteLog(LogLevel.Error, ex.Message, ex.StackTrace);
            }
        }

        void _homegear_EventUpdated(Homegear sender, Event homegearEvent)
        {
            try
            {
                Logging.WriteLog(LogLevel.Info, "HomeGear Event " + homegearEvent.ToString() + " is updated");
            }
            catch (Exception ex)
            {
                Logging.WriteLog(LogLevel.Error, ex.Message, ex.StackTrace);
            }
        }

        void _homegear_Reloaded(Homegear sender)
        {
            _reloading = false;
            try
            {
                _mainInstance.Get("RPC_InitComplete").Set(true);
                Logging.WriteLog(LogLevel.Info, "RPC Init Complete");
            }
            catch (Exception ex)
            {
                Logging.WriteLog(LogLevel.Error, ex.Message, ex.StackTrace);
            }
        }

        void _homegear_DeviceReloadRequired(Homegear sender, Device device, Channel channel, DeviceReloadType reloadType)
        {
            try
            {
                Logging.WriteLog(LogLevel.Info, "RPC: Reload erforderlich (" + reloadType.ToString() + ")");
                _homegearDevicesMutex.WaitOne();
                if (reloadType == DeviceReloadType.Full)
                {
                    Logging.WriteLog(LogLevel.Info, "Reloading device " + device.ID.ToString() + ".");
                    //Finish all operations on the device and then call:
                    device.Reload();
                }
                else if (reloadType == DeviceReloadType.Metadata)
                {
                    Logging.WriteLog(LogLevel.Info, "Reloading metadata of device " + device.ID.ToString() + ".");
                    //Finish all operations on the device's metadata and then call:
                    device.Metadata.Reload();
                }
                else if (reloadType == DeviceReloadType.Channel)
                {
                    Logging.WriteLog(LogLevel.Info, "Reloading channel " + channel.Index + " of device " + device.ID.ToString() + ".");
                    //Finish all operations on the device's channel and then call:
                    channel.Reload();
                }
                else if (reloadType == DeviceReloadType.Variables)
                {
                    Logging.WriteLog(LogLevel.Info, "Device variables were updated: Device type: \"" + device.TypeString + "\", ID: " + device.ID.ToString() + ", Channel: " + channel.Index.ToString());
                    Logging.WriteLog(LogLevel.Info, "Reloading variables of channel " + channel.Index + " and device " + device.ID.ToString() + ".");
                    //Finish all operations on the channels's variables and then call:
                    channel.Variables.Reload();
                }
                else if (reloadType == DeviceReloadType.Links)
                {
                    Logging.WriteLog(LogLevel.Info, "Device links were updated: Device type: \"" + device.TypeString + "\", ID: " + device.ID.ToString() + ", Channel: " + channel.Index.ToString());
                    Logging.WriteLog(LogLevel.Info, "Reloading links of channel " + channel.Index + " and device " + device.ID.ToString() + ".");
                    //Finish all operations on the channels's links and then call:
                    channel.Links.Reload();
                }
                else if (reloadType == DeviceReloadType.Team)
                {
                    Logging.WriteLog(LogLevel.Info, "Device team was updated: Device type: \"" + device.TypeString + "\", ID: " + device.ID.ToString() + ", Channel: " + channel.Index.ToString());
                    Logging.WriteLog(LogLevel.Info, "Reloading channel " + channel.Index + " of device " + device.ID.ToString() + ".");
                    //Finish all operations on the device's channel and then call:
                    channel.Reload();
                }
                else if (reloadType == DeviceReloadType.Events)
                {
                    Logging.WriteLog(LogLevel.Info, "Device events were updated: Device type: \"" + device.TypeString + "\", ID: " + device.ID.ToString() + ", Channel: " + channel.Index.ToString());
                    Logging.WriteLog(LogLevel.Info, "Reloading events of device " + device.ID.ToString() + ".");
                    //Finish all operations on the device's events and then call:
                    device.Events.Reload();
                }
                _homegearDevicesMutex.ReleaseMutex();
            }
            catch (Exception ex)
            {
                try { _homegearDevicesMutex.ReleaseMutex(); } catch (Exception) { }
                Logging.WriteLog(LogLevel.Error, ex.Message, ex.StackTrace);
            }
        }

        void _homegear_ReloadRequired(Homegear sender, ReloadType reloadType)
        {
            try
            {
                Logging.WriteLog(LogLevel.Info, "RPC: Reload erforderlich (" + reloadType.ToString() + ")");
                if (reloadType == ReloadType.Full)
                {
                    try
                    {
                        _mainInstance.Get("RPC_InitComplete").Set(false);
                        while (_reloading)
                        {
                            Logging.WriteLog(LogLevel.Info, "Wait for homegear.Reload()");
                            Thread.Sleep(10);
                        }
                        Logging.WriteLog(LogLevel.Info, "Homegear is full-reloading");
                        _homegearDevicesMutex.WaitOne();
                        _reloading = true;
                        _homegear.Reload();
                        _homegearDevicesMutex.ReleaseMutex();
                    }
                    catch (Exception ex)
                    {
                        _mainInstance.Error = "Reload Thread ist tot";
                        Logging.WriteLog(LogLevel.Error, ex.Message, ex.StackTrace);
                    }
                }
                else if (reloadType == ReloadType.SystemVariables)
                {
                    Logging.WriteLog(LogLevel.Info, "Homegear is reloading SystemVariables");
                    _homegear.SystemVariables.Reload();
                }
                else if (reloadType == ReloadType.Events)
                {
                    Logging.WriteLog(LogLevel.Info, "Homegear is reloading Events");
                    _homegear.TimedEvents.Reload();
                }
            }
            catch (Exception ex)
            {
                _mainInstance.Error = "Reload Thread ist tot";
                Logging.WriteLog(LogLevel.Error, ex.Message, ex.StackTrace);
            }
        }

        void _homegear_DeviceLinkConfigParameterUpdated(Homegear sender, Device device, Channel channel, Link link, ConfigParameter parameter)
        {
            try
            {
                _instances.MutexLocked = true; 
                Logging.WriteLog(LogLevel.Debug, "RPC: " + device.ID.ToString() + " " + link.RemotePeerID.ToString() + " " + link.RemoteChannel.ToString() + " " + parameter.Name + " = " + parameter.ToString());
                if (_instances.ContainsKey(device.ID))
                    _instances[device.ID].Status = "Link-Parameter " + link.Name + " updated to " + link.ToString();
                Logging.WriteLog(LogLevel.Info, "Link-Parameter " + link.Name + " updated to " + link.ToString());
            }
            catch (Exception ex)
            {
                Logging.WriteLog(LogLevel.Error, ex.Message, ex.StackTrace);
            }
            _instances.MutexLocked = false;
        }

        void _homegear_DeviceConfigParameterUpdated(Homegear sender, Device device, Channel channel, ConfigParameter parameter)
        {
            try
            {
                _instances.MutexLocked = true;
                Logging.WriteLog(LogLevel.Debug, "RPC: " + device.ID.ToString() + " " + parameter.Name + " = " + parameter.ToString());
                if (_instances.ContainsKey(device.ID))
                    _instances[device.ID].Status = "Config-Parameter " + parameter.Name + " updated to " + parameter.ToString();
                Logging.WriteLog(LogLevel.Info, "Config-Parameter " + parameter.Name + " updated to " + parameter.ToString());
            }
            catch (Exception ex)
            {
                Logging.WriteLog(LogLevel.Error, ex.Message, ex.StackTrace);
            }
            _instances.MutexLocked = false;
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

        void _homegear_DeviceVariableUpdated(Homegear sender, Device device, Channel channel, Variable variable)
        {
            _instances.MutexLocked = true;
            try
            {
                Int32 deviceID = device.ID;
                Logging.WriteLog(LogLevel.Debug, "RPC: " + deviceID.ToString() + " " + variable.Name + " = " + variable.ToString());
                
                if (_instances.ContainsKey(deviceID))
                {
                    AXInstance instanz = _instances[deviceID];
                    String varName = variable.Name + "_V" + channel.Index.ToString("D2");
                    String subinstance = "V" + channel.Index.ToString("D2");
                    if (instanz.VariableExists(varName))
                    {
                        AXVariable aktAXVar = instanz.Get(varName);
                        if (aktAXVar != null)
                        {
                            _varConverter.SetAXVariable(aktAXVar, variable);
                            setDeviceStatusInMaininstance(variable, deviceID);
                            Logging.WriteLog(LogLevel.Debug, "[HomeGear -> aX]: " + aktAXVar.Path + " = " + variable.ToString());
                        }
                        SetLastChange(instanz, varName + " = " + variable.ToString());
                    }
                    else if (instanz.SubinstanceExists(subinstance))
                    {
                        AXVariable aktAXVar2 = instanz.GetSubinstance(subinstance).Get(variable.Name);
                        if (aktAXVar2 != null)
                        {
                            _varConverter.SetAXVariable(aktAXVar2, variable);
                            setDeviceStatusInMaininstance(variable, deviceID);
                            Logging.WriteLog(LogLevel.Debug, "[HomeGear -> aX]: " + aktAXVar2.Path + " = " + variable.ToString());
                        }
                        SetLastChange(instanz, subinstance + "." + variable.Name + " = " + variable.ToString());
                    }
                    else
                        SetLastChange(instanz, varName + " = " + variable.ToString() + " - Variable nicht vorhanden!");
                    
                    //Console.WriteLine(device.SerialNumber + ": " + variable.Name + ": " + variable.ToString());
                }
            }
            catch (Exception ex)
            {
                Logging.WriteLog(LogLevel.Error, ex.Message, ex.StackTrace);
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
                Logging.WriteLog(LogLevel.Error, ex.Message, ex.StackTrace);
            }
        }


        void _homegear_MetadataUpdated(Homegear sender, Device device, MetadataVariable variable)
        {
            try
            {
                Logging.WriteLog(LogLevel.Info, "RPC: Metadata Variable '" + variable.Name + "' geändert");
            }
            catch (Exception ex)
            {
                Logging.WriteLog(LogLevel.Error, ex.Message, ex.StackTrace);
            }
        }

        void _homegear_SystemVariableUpdated(Homegear sender, SystemVariable variable)
        {
            try
            {
                Logging.WriteLog(LogLevel.Info, "RPC: System Variable '" + variable.Name + "' geändert");
                UpdateSystemVariables();
            }
            catch (Exception ex)
            {
                Logging.WriteLog(LogLevel.Error, ex.Message, ex.StackTrace);
            }
        }

        void _homegear_ConnectError(Homegear sender, string message, string stackTrace)
        {
            try
            {
                _mainInstance.Error = message;
                Logging.WriteLog(LogLevel.Error, message);
            }
            catch (Exception ex)
            {
                Logging.WriteLog(LogLevel.Error, ex.Message, ex.StackTrace);
            }
        }

        void _rpc_ClientDisconnected(RPCClient sender)
        {
            try
            {
                Logging.WriteLog(LogLevel.Warning, "RPC-Client Verbindung unterbrochen");
            }
            catch (Exception ex)
            {
                Logging.WriteLog(LogLevel.Error, ex.Message, ex.StackTrace);
            }
        }

        void _rpc_ClientConnected(RPCClient sender, CipherAlgorithmType cipherAlgorithm, Int32 cipherStrength)
        {
            try
            {
                Logging.WriteLog(LogLevel.Info, "RPC-Client verbunden");
                _mainInstance.Get("err").Set(false);
            }
            catch (Exception ex)
            {
                Logging.WriteLog(LogLevel.Error, ex.Message, ex.StackTrace);
            }
        }

        void _rpc_ServerDisconnected(RPCServer sender)
        {
            try
            {
                Logging.WriteLog(LogLevel.Warning, "Verbindung von Homegear zu aX unterbrochen");
            }
            catch (Exception ex)
            {
                Logging.WriteLog(LogLevel.Error, ex.Message, ex.StackTrace);
            }
        }

        void _rpc_ServerConnected(RPCServer sender, CipherAlgorithmType cipherAlgorithm, Int32 cipherStrength)
        {
            try
            {
                Logging.WriteLog(LogLevel.Info, "Eingehende Verbindung von Homegear hergestellt");
                _mainInstance.Get("err").Set(false);
                if (!_reloading && !_mainInstance.Get("RPC_InitComplete").GetBool())
                {
                    _reloading = true;
                    _homegear.Reload();
                }
            }
            catch (Exception ex)
            {
                Logging.WriteLog(LogLevel.Error, ex.Message, ex.StackTrace);
            }
        }

        void IDisposable.Dispose()
        {
            Dispose();
        }
    }
}
