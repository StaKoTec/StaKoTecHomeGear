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
        private AXInstance _mainInstance = null;


        public VariableConverter(AXInstance mainInstance)
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
                    for (int i = 0; i < aktVar.ValueList.Length; i++)
                    {
                        if (aktVar.ValueList[i] == "")
                            continue;
                        typ += "(" + i.ToString() + ": " + aktVar.ValueList[i] + ")" + ((i == aktVar.ValueList.Length - 1) ? "" : ",");
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
                Logging.WriteLog(ex.Message, ex.StackTrace);
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
                    for (int i = 0; i < aktVar.ValueList.Length; i++)
                    {
                        if (aktVar.ValueList[i] == "")
                            continue;
                        typ += "(" + i.ToString() + ": " + aktVar.ValueList[i] + ")" + ((i == aktVar.ValueList.Length - 1) ? "" : ",");
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
                Logging.WriteLog(ex.Message, ex.StackTrace);
            }
        }


        public Boolean ParseAXVariable(AXVariable var, out String name, out String type, out Int32 channel)
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
                Logging.WriteLog(ex.Message, ex.StackTrace);
            }
            return true;
        }

        public void SetHomeGearVariable(Variable homegearVar, AXVariable aXVar)
        {
            try
            {
                if (!homegearVar.Writeable)
                    return;

                switch (aXVar.Type)
                {
                    case AXVariableType.axBool:
                        if ((homegearVar.Type == VariableType.tAction) && !aXVar.GetBool()) return;
                        homegearVar.BooleanValue = aXVar.GetBool();
                        break;
                    case AXVariableType.axInteger:
                        homegearVar.IntegerValue = aXVar.GetInteger();
                        break;
                    case AXVariableType.axLongInteger:
                        homegearVar.IntegerValue = aXVar.GetLongInteger();
                        break;
                    case AXVariableType.axLongReal:
                        homegearVar.DoubleValue = aXVar.GetLongReal();
                        break;
                    case AXVariableType.axString:
                        homegearVar.StringValue = aXVar.GetString();
                        break;
                }
            }
            catch (Exception ex)
            {
                _mainInstance.Error = ex.Message;
                Logging.WriteLog(ex.Message, ex.StackTrace);
            }

        }

        public String AutomationXVarToString(AXVariable Var)
        {
            String stringVar = "AUTOMATIONX VAR-TYPE NOT FOUND";
            try
            {
                switch (Var.Type)
                {
                    case AXVariableType.axBool:
                        stringVar = Var.GetBool().ToString();
                        break;
                    case AXVariableType.axByte:
                        stringVar = Var.GetByte().ToString();
                        break;
                    case AXVariableType.axInteger:
                        stringVar = Var.GetInteger().ToString();
                        break;
                    case AXVariableType.axLongInteger:
                        stringVar = Var.GetLongInteger().ToString();
                        break;
                    case AXVariableType.axUnsignedInteger:
                        stringVar = Var.GetUnsignedInteger().ToString();
                        break;
                    case AXVariableType.axUnsignedLongInteger:
                        stringVar = Var.GetUnsignedLongInteger().ToString();
                        break;
                    case AXVariableType.axReal:
                        stringVar = Var.GetReal().ToString();
                        break;
                    case AXVariableType.axLongReal:
                        stringVar = Var.GetLongReal().ToString();
                        break;
                    case AXVariableType.axShortInteger:
                        stringVar = Var.GetShortInteger().ToString();
                        break;
                    case AXVariableType.axString:
                        stringVar = Var.GetString();
                        break;
                }
            }
            catch (Exception ex)
            {
                _mainInstance.Error = ex.Message;
                Logging.WriteLog(ex.Message, ex.StackTrace);
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
                _mainInstance.Error = ex.Message;
                Logging.WriteLog(ex.Message, ex.StackTrace);
            }
            return stringVar;
        }


        public void SetAXVariable(AXVariable aXVar, Variable homegearVar)
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
                _mainInstance.Error = ex.Message;
                Logging.WriteLog(ex.Message, ex.StackTrace);
            }
        }
    }
}
