using System;
using System.IO;
using System.Collections.Generic;
using System.Windows.Forms;
using Wpm.Schema.Kernel;
using Wpm.Implement.Manager;
using Wpm.Implement.ModelSetting;
using Alma.BaseUI.DescriptionEditor;

namespace Wpm.Implement.ModelSetting
{
    public partial class ImportUserCommandType : ScriptModelCustomization, IScriptModelCustomization
    {
        public override bool Execute(IContext context, IContext hostContext)
        {
            #region Import_OF
            
            {
                ICommandTypeFactory commandTypeFactory = new CommandTypeFactory(context, 1, null, true);
                commandTypeFactory.Key = "AF_JOHN_DEERE_DLL";
                commandTypeFactory.Name = "Import_OF";
                commandTypeFactory.PlugInFileName = @"AF_JOHN_DEERE_Dll.dll";
                commandTypeFactory.PlugInNameSpace = "AF_JOHN_DEERE";
                commandTypeFactory.PlugInClassName = "JohnDeereIE";
                commandTypeFactory.Shortcut = Shortcut.None;
                commandTypeFactory.LargeImage = true;
                commandTypeFactory.ImageFile = "";
                
                {
                    IParameterDescription parameterDescription = new ParameterDescription(context.Kernel.UnitSystem, true);
                    parameterDescription.Key = "IMPORT_CDA";
                    parameterDescription.Name = "OF";
                    parameterDescription.ParameterDescriptionType = ParameterDescriptionType.String;
                    parameterDescription.DefaultValue = @"C:\temp";
                    commandTypeFactory.ParameterList.Add(parameterDescription);
                    
                }
                {
                    IParameterDescription parameterDescription = new ParameterDescription(context.Kernel.UnitSystem, true);
                    parameterDescription.Key = "EXPORT_Rp";
                    parameterDescription.Name = "Export Sage";
                    parameterDescription.ParameterDescriptionType = ParameterDescriptionType.String;
                    parameterDescription.DefaultValue = @"C:\temp";
                    commandTypeFactory.ParameterList.Add(parameterDescription);
                    
                }
                {
                    IParameterDescription parameterDescription = new ParameterDescription(context.Kernel.UnitSystem, true);
                    parameterDescription.Key = "MODEL_CA";
                    parameterDescription.Name = "Model OF";
                    parameterDescription.ParameterDescriptionType = ParameterDescriptionType.String;
                    parameterDescription.DefaultValue = @"0#AF_PT#string;1#AF_COMPOSANT#string;2#AF_ORDRE#string;3#_NAME#string;4#_QUANTITY#int;5#AF_BESOIN#string;6#AF_DESTINATION#date;7#AF_BROSS_EBAV#string;8#AF_KB_MRP#string;9#AF_AVANCE_RETARD#string;10#AF_CONTENANT#string;11#AF_CONTENANT_TH#string";
                    commandTypeFactory.ParameterList.Add(parameterDescription);
                    
                }
                {
                    IParameterDescription parameterDescription = new ParameterDescription(context.Kernel.UnitSystem, true);
                    parameterDescription.Key = "MODEL_PATH";
                    parameterDescription.Name = "Model Chemin export";
                    parameterDescription.ParameterDescriptionType = ParameterDescriptionType.String;
                    parameterDescription.DefaultValue = @"0#TECHNOLOGIE#string;1#ImportCda#string;0#ImportDM#string;2#ExportRp#string;3#ExportDT#string;4#Centredefrais#string;5#Destination_Path#string;6#Source_Path#string";
                    commandTypeFactory.ParameterList.Add(parameterDescription);
                    
                }
                {
                    IParameterDescription parameterDescription = new ParameterDescription(context.Kernel.UnitSystem, true);
                    parameterDescription.Key = "STRING_FORMAT_DOUBLE";
                    parameterDescription.Name = "Format double";
                    parameterDescription.ParameterDescriptionType = ParameterDescriptionType.String;
                    parameterDescription.DefaultValue = @"{0:0.00###}";
                    commandTypeFactory.ParameterList.Add(parameterDescription);
                    
                }
                {
                    IParameterDescription parameterDescription = new ParameterDescription(context.Kernel.UnitSystem, true);
                    parameterDescription.Key = "VERBOSE_LOG";
                    parameterDescription.Name = "Log Verbeux";
                    parameterDescription.ParameterDescriptionType = ParameterDescriptionType.Boolean;
                    parameterDescription.DefaultValue = false;
                    commandTypeFactory.ParameterList.Add(parameterDescription);
                    
                }
                {
                    IParameterDescription parameterDescription = new ParameterDescription(context.Kernel.UnitSystem, true);
                    parameterDescription.Key = "IMPORT_AUTO";
                    parameterDescription.Name = "Activer import auto";
                    parameterDescription.ParameterDescriptionType = ParameterDescriptionType.Boolean;
                    parameterDescription.DefaultValue = false;
                    commandTypeFactory.ParameterList.Add(parameterDescription);
                    
                }
                if (!commandTypeFactory.UpdateModel())
                {
                    foreach (ModelSettingError error in commandTypeFactory.ErrorList)
                    {
                        hostContext.TraceLogger.TraceError(error.Message, true);
                    }
                    return false;
                }
                
            }
            
            #endregion
            return true;
        }
    }
}
