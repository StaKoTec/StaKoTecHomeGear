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
                    Console.WriteLine(ex.Message);
                    Dispose();
                }

                _aX.ShuttingDown += aX_ShuttingDown;

                _mainInstance = new AXInstance(_aX, instanceName, "Status", "err"); //Instanz-Objekt erstellen
                _mainInstance.PollingInterval = 10;
                _mainInstance.VariableEvents = true;
                Logging.Init(_aX, _mainInstance);
                _varConverter = new VariableConverter(_mainInstance);

                AXVariable init = _mainInstance.Get("Init");
                _mainInstance.Get("Init").Set(false);
                init.ValueChanged += init_ValueChanged;
                AXVariable pairingMode = _mainInstance.Get("PairingMode");
                _mainInstance.Get("PairingMode").Set(false);
                pairingMode.ValueChanged += pairingMode_ValueChanged;
                AXVariable deviceUnpair = _mainInstance.Get("DeviceUnpair");
                _mainInstance.Get("DeviceUnpair").Set(false);
                deviceUnpair.ValueChanged += deviceUnpair_ValueChanged;
                AXVariable deviceReset = _mainInstance.Get("DeviceReset");
                _mainInstance.Get("DeviceReset").Set(false);
                deviceReset.ValueChanged += deviceReset_ValueChanged;
                AXVariable deviceRemove = _mainInstance.Get("DeviceRemove");
                _mainInstance.Get("DeviceRemove").Set(false);
                deviceRemove.ValueChanged += deviceRemove_ValueChanged;
                AXVariable getDeviceVars = _mainInstance.Get("GetDeviceVars");
                _mainInstance.Get("GetDeviceVars").Set(false);
                getDeviceVars.ValueChanged += getDeviceVars_ValueChanged;

                AXVariable lifetick = _mainInstance.Get("Lifetick");
                AXVariable startStaKoTCPIPRelease = _mainInstance.Get("StartStaKo_TCPIPRelease");
                startStaKoTCPIPRelease.ValueChanged += startStaKoTCPIPRelease_ValueChanged;
                if (!startStaKoTCPIPRelease.GetBool())
                    Dispose();

                _mainInstance.Get("RPC_InitComplete").Set(false);
                _mainInstance.Get("StaKo_TCPIP_Running").Set(true);


                Logging.WriteLog("HomeGear started");

                HomeGearConnect();
                Int32 i = 0;
                while(!_rpc.IsConnected)
                {
                    Console.WriteLine(i.ToString() + ": Waiting for RPC-Server connection");
                    Thread.Sleep(10);
                    i++;
                }

                while (!_disposing)
                {
                    try
                    {
                        lifetick.Set(true);

                        //Allen Devices einen Lifetick senden um DataValid zu generieren
                        if (_instanzen != null && _instanzen.Count() > 0)
                        {
                            foreach (KeyValuePair<Int32, AXInstance> instance in _instanzen)
                            {
                                instance.Value.Get("Lifetick").Set(true);
                            }
                        }

                        if (_homegear != null && _initCompleted)
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
                    }
                    catch (Exception ex)
                    {
                        Logging.WriteLog(ex.Message, ex.StackTrace);
                        Console.WriteLine(ex.Message + "\r\nStack Trace:\r\n" + ex.StackTrace);
                    }
                    Thread.Sleep(1000);
                }
            }
            catch(Exception ex)
            {
                Logging.WriteLog(ex.Message, ex.StackTrace);
                Console.WriteLine(ex.Message + "\r\nStack Trace:\r\n" + ex.StackTrace);
            }
        }

        void getDeviceVars_ValueChanged(AXVariable sender)
        {
            try
            {
                AXVariable deviceVars = _mainInstance.Get("DeviceVars");
                AXVariable deviceConfigVars = _mainInstance.Get("DeviceConfigVars");
                UInt16 x = 0;
                UInt16 y = 0;
                Int32 deviceID = _mainInstance.Get("ActionID").GetLongInteger();
                sender.Instance.Status = "Get VariableNames for DeviceID: " + deviceID.ToString();

                _homegearDevicesMutex.WaitOne();
                if (_homegear.Devices.ContainsKey(deviceID))
                {
                    Device aktDevice = _homegear.Devices[deviceID];
                    foreach (KeyValuePair<Int32, Channel> aktChannel in aktDevice.Channels)
                    {
                        foreach (KeyValuePair<String, ConfigParameter> configName in aktDevice.Channels[aktChannel.Key].Config)
                        {
                            if (x >= deviceConfigVars.Length)
                            {
                                _mainInstance.Error = "Array-Index zu klein bei 'DeviceConfigVars'";
                                return;
                            }
                            String textMinMax = "";
                            var aktVar = configName.Value;
                            String minVar = "";
                            String maxVar = "";
                            if (aktVar.Type == VariableType.tInteger)
                            {
                                minVar = aktVar.MinInteger.ToString();
                                maxVar = aktVar.MaxInteger.ToString();
                            }
                            if (aktVar.Type == VariableType.tDouble)
                            {
                                minVar = aktVar.MinDouble.ToString();
                                maxVar = aktVar.MaxDouble.ToString();
                            }
                            if (minVar.Length > 0)
                                textMinMax = " Min: " + minVar + " Max: " + maxVar;

                            deviceConfigVars.Set(x, aktVar.Name + " (" + aktVar.Type.ToString() + ")" + textMinMax);
                            x++;
                        }
                        foreach (KeyValuePair<String, Variable> varName in aktDevice.Channels[aktChannel.Key].Variables)
                        {
                            if (y >= deviceVars.Length)
                            {
                                _mainInstance.Error = "Array-Index zu klein bei 'DeviceVars'";
                                return;
                            }
                            String textMinMax = "";
                            var aktVar = varName.Value;
                            String minVar = "";
                            String maxVar = "";
                            if (aktVar.Type == VariableType.tInteger)
                            {
                                minVar = aktVar.MinInteger.ToString();
                                maxVar = aktVar.MaxInteger.ToString();
                            }
                            if (aktVar.Type == VariableType.tDouble)
                            {
                                minVar = aktVar.MinDouble.ToString();
                                maxVar = aktVar.MaxDouble.ToString();
                            }
                            if (minVar.Length > 0)
                                textMinMax = " Min: " + minVar + " Max: " + maxVar;

                            deviceVars.Set(y, aktVar.Name + " (" + aktVar.Type.ToString() + ")" + textMinMax);
                            y++;
                        }
                    }
                    for (; x < deviceConfigVars.Length; x++)
                        deviceConfigVars.Set(x, "");
                    for (; y < deviceConfigVars.Length; y++)
                        deviceConfigVars.Set(y, "");

                }
            }
            catch (Exception ex)
            {
                
                sender.Instance.Error = ex.Message;
                sender.Instance.Status = ex.Message;
            }
            finally
            {
                _homegearDevicesMutex.ReleaseMutex();
                sender.Set(false);
            }
        }

        void deviceRemove_ValueChanged(AXVariable sender)
        {
            try
            {
                Int32 deviceID = _mainInstance.Get("ActionID").GetLongInteger();
                sender.Instance.Status = "Removing Device ID " + deviceID.ToString();
                
                _rpc.DeleteDevice(deviceID, RPCDeleteDeviceFlags.Force);

                sender.Set(false);
                sender.Instance.Status = "Removing Device ID " + deviceID.ToString() + " complete";
            }
            catch (Exception ex)
            {
                sender.Instance.Error = ex.Message;
                sender.Instance.Status = ex.Message;
            }
        }

        void deviceReset_ValueChanged(AXVariable sender)
        {
            try
            {
                Int32 deviceID = _mainInstance.Get("ActionID").GetLongInteger();
                sender.Instance.Status = "Resetting Device ID " + deviceID.ToString();
                
                _rpc.DeleteDevice(deviceID, RPCDeleteDeviceFlags.Reset | RPCDeleteDeviceFlags.Defer);
                
                sender.Set(false);
                sender.Instance.Status = "Resetting Device ID " + deviceID.ToString() + " complete";
            }
            catch (Exception ex)
            {
                sender.Instance.Error = ex.Message;
                sender.Instance.Status = ex.Message;
            }
        }

        void deviceUnpair_ValueChanged(AXVariable sender)
        {
            try
            {
                Int32 deviceID = _mainInstance.Get("ActionID").GetLongInteger();
                sender.Instance.Status = "Unpairing Device ID " + deviceID.ToString();

                _rpc.DeleteDevice(deviceID, RPCDeleteDeviceFlags.Defer);

                sender.Set(false);
                sender.Instance.Status = "Unpairing Device ID " + deviceID.ToString() + " complete";
            }
            catch (Exception ex)
            {
                sender.Instance.Error = ex.Message;
                sender.Instance.Status = ex.Message;
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
            }
        }

        void startStaKoTCPIPRelease_ValueChanged(AXVariable sender)
        {
            try
            {
                if (!sender.GetBool())
                {
                    _mainInstance.Status = "Beende StaKoTecHomeGear.exe";
                    _mainInstance.Get("err").Set(false);
                    Console.WriteLine("Beende StaKoTecHomeGear.exe");
                    Dispose();
                }
            }
            catch (Exception ex)
            {
                sender.Instance.Error = ex.Message;
                sender.Instance.Status = ex.Message;
            }
        }

        void init_ValueChanged(AXVariable sender)
        {
            try
            {
                if (!sender.GetBool())
                    return;

                UInt16 x = 0;
                _mainInstance.Status = "Init";

                _deviceID = sender.Instance.Get("DeviceID");
                _deviceInstance = sender.Instance.Get("DeviceInstance");
                _deviceRemark = sender.Instance.Get("DeviceRemark");
                _deviceTypeString = sender.Instance.Get("DeviceTypeString");
                _deviceState = sender.Instance.Get("DeviceState");
                _deviceStateColor = sender.Instance.Get("DeviceStateColor");

                _mainInstance.Get("HomeGearVersion").Set(_homegear.Version);
                
                //////////////////////////////
                // Alle HomeGear Instanzen auslesen
                AXVariable homeGearKlassen = _mainInstance.Get("HomeGearKlassen");
                List<String> klassen = new List<string>();
                for (x = 0; x < homeGearKlassen.Length; x++)
                {
                    String temp = homeGearKlassen.GetString(x).Trim();
                    if (temp.Length > 0) klassen.Add(temp);
                    else break;
                }

                _instanzenMutex.WaitOne();
                _instanzen = new Dictionary<Int32, AXInstance>();
                foreach(String name in klassen)
                {
                    List<String> instanceNames = _aX.GetInstanceNames(name);
                    foreach(String name2 in instanceNames)
                    {
                        AXInstance instanz = new AXInstance(_aX, name2, "Status", "err");
                        _instanzen.Add(instanz.Get("ID").GetLongInteger(), instanz);
                    }
                }
                

                _homegearDevicesMutex.WaitOne();
                x = 0;
                foreach(KeyValuePair<Int32, Device> devicePair in _homegear.Devices)
                {
                    _deviceID.Set(x, devicePair.Key);
                    if (_instanzen.ContainsKey(devicePair.Key))
                    {
                        AXInstance aktInstanz = _instanzen[devicePair.Key];
                        _deviceTypeString.Set(x, devicePair.Value.TypeString);
                        _deviceInstance.Set(x, aktInstanz.Name);
                        _deviceRemark.Set(x, aktInstanz.Remark);
                        _deviceState.Set(x, "OK");
                        _deviceStateColor.Set(x, (Int16)DeviceStatus.OK);

                        aktInstanz.Get("SerialNo").Set(devicePair.Value.SerialNumber);
                        aktInstanz.VariableEvents = true;
                        aktInstanz.VariableValueChanged += aktInstanz_VariableValueChanged;
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
 
                _homegearDevicesMutex.ReleaseMutex();
                for (; x < _deviceID.Length; x++)
                {
                    _deviceID.Set(x, 0);
                    _deviceTypeString.Set(x, "");
                    _deviceInstance.Set(x, "");
                    _deviceRemark.Set(x, "");
                    _deviceState.Set(x, "");
                    _deviceStateColor.Set(x, (Int16)DeviceStatus.Nichts);

                }
                _instanzenMutex.ReleaseMutex();
                sender.Set(false);
                _initCompleted = true;
            }
            catch (Exception ex)
            {
                sender.Instance.Error = ex.Message;
                _homegearDevicesMutex.ReleaseMutex();
                _instanzenMutex.ReleaseMutex();
                sender.Set(false);
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
                            return;

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
                                }
                            }
                        }
                        sender.Set(false);
                    }
                    else
                    {
                        if (_varConverter.ParseAXVariable(sender, out name, out type, out channelIndex))
                        {
                            if (aktDevice.Channels.ContainsKey(channelIndex))
                            {
                                Channel channel = aktDevice.Channels[channelIndex];
                                if (type == "V")
                                {
                                    if (channel.Variables.ContainsKey(name))
                                    {

                                        String status = "Set Variable " + sender.Instance.Name + "." + name + ", Channel:" + channelIndex.ToString() + "";
                                        _mainInstance.Status = status;
                                        Console.WriteLine(status);
                                        _varConverter.SetHomeGearVariable(channel.Variables[name], sender);
                                    }
                                }
                                else if (type == "C")
                                {
                                    if (channel.Config.ContainsKey(name))
                                    {
                                        String status = "Set Config " + sender.Instance.Name + "." + name + ", Channel: " + channelIndex.ToString() + "";
                                        _mainInstance.Status = status;
                                        Console.WriteLine(status);
                                        _varConverter.SetHomeGearVariable(channel.Config[name], sender);
                                    }
                                }
                            }
                        }
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
        }

        void aX_ShuttingDown(AX sender)
        {
            Dispose();
        }


        void Dispose()
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
            _mainInstance.Status = "StaKoTecHomeGear.exe beendet";
            Console.WriteLine("Und aus!!");
            Environment.Exit(0);
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
            }
            catch (Exception ex)
            {
                _mainInstance.Error = ex.Message;
                Logging.WriteLog(ex.Message, ex.StackTrace);
            }
        }

        void _homegear_Reloaded(Homegear sender)
        {
            try
            {
                _mainInstance.Status = "RPC: Reload feddich";
                _mainInstance.Get("RPC_InitComplete").Set(true);
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
                if (reloadType == DeviceReloadType.Channel) channel.Reload();
                else if (reloadType == DeviceReloadType.Full) device.Reload();
                else if (reloadType == DeviceReloadType.Links) channel.Links.Reload();
                else if (reloadType == DeviceReloadType.Metadata) device.Metadata.Reload();
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
            _mainInstance.Status = "RPC: Reload erforderlich (" + reloadType.ToString() + ")";
            try
            {
                if (reloadType == ReloadType.Full)
                {
                    _mainInstance.Get("RPC_InitComplete").Set(false);
                    _homegearDevicesMutex.WaitOne();
                    _homegear.Reload();
                    _homegearDevicesMutex.ReleaseMutex();
                }
                else if (reloadType == ReloadType.SystemVariables) _homegear.SystemVariables.Reload();
            }
            catch (Exception ex)
            {
                _homegearDevicesMutex.ReleaseMutex();
                _mainInstance.Error = "Reload Thread ist tot";
                Logging.WriteLog(ex.Message, ex.StackTrace);
            }
        }

        void _homegear_DeviceLinkConfigParameterUpdated(Homegear sender, Device device, Channel channel, Link link, ConfigParameter parameter)
        {
            try
            {
                _mainInstance.Status = "RPC: " + device.ID.ToString() + " " + link.RemotePeerID.ToString() + " " + link.RemoteChannel.ToString() + " " + parameter.Name + " = " + parameter.ToString();
            }
            catch (Exception ex)
            {
                _mainInstance.Error = ex.Message;
                Logging.WriteLog(ex.Message, ex.StackTrace);
            }
        }

        void _homegear_DeviceConfigParameterUpdated(Homegear sender, Device device, Channel channel, ConfigParameter parameter)
        {
            try
            {
                _mainInstance.Status = "RPC: " + device.ID.ToString() + " " + parameter.Name + " = " + parameter.ToString();
            }
            catch (Exception ex)
            {
                _mainInstance.Error = ex.Message;
                Logging.WriteLog(ex.Message, ex.StackTrace);
            }
        }

        void _homegear_DeviceVariableUpdated(Homegear sender, Device device, Channel channel, Variable variable)
        {
            try
            {
                _mainInstance.Status = "RPC: " + device.ID.ToString() + " " + variable.Name + " = " + variable.ToString();
                _homegearDevicesMutex.WaitOne();
                if(_instanzen.ContainsKey(device.ID))
                {
                    AXInstance instanz = _instanzen[device.ID];
                    AXVariable aXVariable = instanz.Get(variable.Name + "_V" + channel.Index.ToString("D2"));
                    if (aXVariable != null) _varConverter.SetAXVariable(aXVariable, variable);
                    //Console.WriteLine(device.SerialNumber + ": " + variable.Name + ": " + variable.ToString());
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
                _mainInstance.Status = message;
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
                _mainInstance.Status = "RPC-Client Verbindung unterbrochen";
                Console.WriteLine("RPC-Client Verbindung unterbrochen");
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
                _mainInstance.Status = "RPC-Client verbunden";
                Console.WriteLine("RPC-Client verbunden");
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
                _mainInstance.Status = "Verbindung von Homegear zu aX unterbrochen";
                Console.WriteLine("Verbindung von Homegear zu aX unterbrochen");
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
                _mainInstance.Status = "Eingehende Verbindung von Homegear hergestellt";
                Console.WriteLine("Eingehende Verbindung von Homegear hergestellt");
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
