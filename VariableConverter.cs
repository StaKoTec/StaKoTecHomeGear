using AutomationX;
using HomegearLib;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace StaKoTecHomeGear
{
    class VariableConverter
    {
        private AxInstance _mainInstance = null;


        public VariableConverter(AxInstance mainInstance)
        {
            _mainInstance = mainInstance;
        }


        public void ParseDeviceConfigVars(ConfigParameter aktVar, out String minVar, out String maxVar, out String typ, out String defaultVar, out String rwVar)
        {
            minVar = "";
            maxVar = "";
            typ = "";
            defaultVar = "";
            rwVar = "";
            try
            {
                if (aktVar.Type == VariableType.tInteger)
                {
                    minVar = aktVar.MinInteger.ToString();
                    maxVar = aktVar.MaxInteger.ToString();
                    defaultVar = aktVar.DefaultInteger.ToString();
                    typ = "DINT";
                }
                else if (aktVar.Type == VariableType.tDouble)
                {
                    minVar = aktVar.MinDouble.ToString();
                    maxVar = aktVar.MaxDouble.ToString();
                    defaultVar = aktVar.DefaultDouble.ToString();
                    typ = "LREAL";
                }
                else if (aktVar.Type == VariableType.tBoolean)
                {
                    defaultVar = aktVar.DefaultBoolean.ToString();
                    typ = "BOOL";
                }
                else if (aktVar.Type == VariableType.tAction)
                {
                    defaultVar = "";
                    typ = "ACTION (BOOL)";
                }
                else if (aktVar.Type == VariableType.tEnum)
                {
                    defaultVar = aktVar.DefaultInteger.ToString();
                    typ = "ENUM (DINT) {";
                    foreach(KeyValuePair<int, String> aktPair in aktVar.ValueList)
                    {
                        if (aktPair.Value == "")
                            continue;
                        typ += "(" + aktPair.Key.ToString() + ": " + aktPair.Value + "), ";
                    }
                    typ += "}";
                }
                else
                {
                    defaultVar = aktVar.DefaultString;
                    typ = "STRING";
                }
                if (aktVar.Readable)
                    rwVar = "R";
                else
                    rwVar = "-";

                if (aktVar.Writeable)
                    rwVar += "/W";
                else
                    rwVar += "/-";
            }
            catch (Exception ex)
            {
                Logging.WriteLog(LogLevel.Error, _mainInstance, ex.Message, ex.StackTrace);
            }
        }

        public void ParseDeviceVars(Variable aktVar, out String minVar, out String maxVar, out String typ, out String defaultVar, out String rwVar)
        {
            minVar = "";
            maxVar = "";
            typ = "";
            defaultVar = "";
            rwVar = "";
            try
            {
                if (aktVar.Type == VariableType.tInteger)
                {
                    minVar = aktVar.MinInteger.ToString();
                    maxVar = aktVar.MaxInteger.ToString();
                    defaultVar = aktVar.DefaultInteger.ToString();
                    typ = "DINT";
                }
                else if (aktVar.Type == VariableType.tDouble)
                {
                    minVar = aktVar.MinDouble.ToString();
                    maxVar = aktVar.MaxDouble.ToString();
                    defaultVar = aktVar.DefaultDouble.ToString();
                    typ = "LREAL";
                }
                else if (aktVar.Type == VariableType.tBoolean)
                {
                    defaultVar = aktVar.DefaultBoolean.ToString();
                    typ = "BOOL";
                }
                else if (aktVar.Type == VariableType.tAction)
                {
                    defaultVar = "";
                    typ = "ACTION (BOOL)";
                }
                else if (aktVar.Type == VariableType.tEnum)
                {
                    defaultVar = aktVar.DefaultInteger.ToString();
                    typ = "ENUM (DINT) {";
                    foreach (KeyValuePair<int, String> aktPair in aktVar.ValueList)
                    {
                        if (aktPair.Value == "")
                            continue;
                        typ += "(" + aktPair.Key.ToString() + ": " + aktPair.Value + "), ";
                    }
                    typ += "}";
                }
                else
                {
                    defaultVar = aktVar.DefaultString;
                    typ = "STRING";
                }
                if (aktVar.Readable)
                    rwVar = "R";
                else
                    rwVar = "-";

                if (aktVar.Writeable)
                    rwVar += "/W";
                else
                    rwVar += "/-";
            }
            catch (Exception ex)
            {
                Logging.WriteLog(LogLevel.Error, _mainInstance, ex.Message, ex.StackTrace);
            }
        }


        public Boolean ParseAxVariable(AxVariable var, out String name, out String type, out Int32 channel)
        {
            name = "";
            channel = -1;
            type = "";
            try
            {
                if ((var == null) || (var.Name.Length < 5))
                    return false;

                name = var.Name.Substring(0, (var.Name.Length - 4));
                type = var.Name.Substring((var.Name.Length - 3), 1);
                if (!Int32.TryParse(var.Name.Substring((var.Name.Length - 2), 2), out channel))
                    return false;
            }
            catch (Exception ex)
            {
                Logging.WriteLog(LogLevel.Error, _mainInstance, ex.Message, ex.StackTrace);
            }
            return true;
        }

        public void SetHomeGearVariable(Variable homegearVar, AxVariable aXVar)
        {
            try
            {
                if (!homegearVar.Writeable)
                    return;

                switch (aXVar.Type)
                {
                    case AxVariableType.axBool:
                        if ((homegearVar.Type == VariableType.tAction) && !aXVar.GetBool()) return;
                        homegearVar.BooleanValue = aXVar.GetBool();
                        break;
                    case AxVariableType.axInteger:
                        homegearVar.IntegerValue = aXVar.GetInteger();
                        break;
                    case AxVariableType.axLongInteger:
                        homegearVar.IntegerValue = aXVar.GetLongInteger();
                        break;
                    case AxVariableType.axLongReal:
                        homegearVar.DoubleValue = aXVar.GetLongReal();
                        break;
                    case AxVariableType.axString:
                        homegearVar.StringValue = aXVar.GetString();
                        break;
                }
            }
            catch (Exception ex)
            {
                Logging.WriteLog(LogLevel.Error, _mainInstance, ex.Message, ex.StackTrace);
            }

        }

        public String AutomationXVarToString(AxVariable Var)
        {
            String stringVar = "AUTOMATIONX VAR-TYPE NOT FOUND";
            try
            {
                switch (Var.Type)
                {
                    case AxVariableType.axBool:
                        stringVar = Var.GetBool().ToString();
                        break;
                    case AxVariableType.axByte:
                        stringVar = Var.GetByte().ToString();
                        break;
                    case AxVariableType.axInteger:
                        stringVar = Var.GetInteger().ToString();
                        break;
                    case AxVariableType.axLongInteger:
                        stringVar = Var.GetLongInteger().ToString();
                        break;
                    case AxVariableType.axUnsignedInteger:
                        stringVar = Var.GetUnsignedInteger().ToString();
                        break;
                    case AxVariableType.axUnsignedLongInteger:
                        stringVar = Var.GetUnsignedLongInteger().ToString();
                        break;
                    case AxVariableType.axReal:
                        stringVar = Var.GetReal().ToString();
                        break;
                    case AxVariableType.axLongReal:
                        stringVar = Var.GetLongReal().ToString();
                        break;
                    case AxVariableType.axShortInteger:
                        stringVar = Var.GetShortInteger().ToString();
                        break;
                    case AxVariableType.axString:
                        stringVar = Var.GetString();
                        break;
                }
            }
            catch (Exception ex)
            {
                Logging.WriteLog(LogLevel.Error, _mainInstance, ex.Message, ex.StackTrace);
            }
            return stringVar;
        }

        public String HomegearVarToString(Variable Var)
        {
            String stringVar = "HOMEGEAR VAR-TYPE NOT FOUND";
            try
            {
                switch (Var.Type)
                {
                    case VariableType.tBoolean:
                    case VariableType.tAction:
                        stringVar = Var.BooleanValue.ToString();
                        break;
                    case VariableType.tDouble:
                        stringVar = Var.DoubleValue.ToString();
                        break;
                    case VariableType.tEnum:
                    case VariableType.tInteger:
                        stringVar = Var.IntegerValue.ToString();
                        break;
                    case VariableType.tString:
                        stringVar = Var.StringValue;
                        break;
                }
            }
            catch (Exception ex)
            {
                Logging.WriteLog(LogLevel.Error, _mainInstance, ex.Message, ex.StackTrace);
            }
            return stringVar;
        }


        public void SetAxVariable(AxVariable aXVar, Variable homegearVar)
        {
            try
            {
                switch (homegearVar.Type)
                {
                    case VariableType.tBoolean:
                        aXVar.Set(homegearVar.BooleanValue);
                        break;
                    case VariableType.tAction:
                        aXVar.Set(homegearVar.BooleanValue);
                        break;
                    case VariableType.tDouble:
                        aXVar.Set(homegearVar.DoubleValue);
                        break;
                    case VariableType.tEnum:
                        aXVar.Set(homegearVar.IntegerValue);
                        break;
                    case VariableType.tInteger:
                        aXVar.Set(homegearVar.IntegerValue);
                        break;
                    case VariableType.tString:
                        aXVar.Set(homegearVar.StringValue);
                        break;
                }
            }
            catch (Exception ex)
            {
                Logging.WriteLog(LogLevel.Error, _mainInstance, ex.Message, ex.StackTrace);
            }
        }
    }
}
