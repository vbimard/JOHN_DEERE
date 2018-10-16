using System;
using System.IO;
using System.Collections.Generic;
using Wpm.Schema.Kernel;
using Wpm.Implement.Manager;
using Wpm.Implement.ModelSetting;
using Alma.BaseUI.DescriptionEditor;

namespace Wpm.Implement.ModelSetting
{
    public partial class ImportUserEntityType : ScriptModelCustomization, IScriptModelCustomization
    {
        public override bool Execute(IContext context, IContext hostContext)
        {
            #region Pièces 2D
            
            {
                IEntityType entityType = context.Kernel.GetEntityType("_REFERENCE");
                IEntityTypeFactory entityTypeFactory = new EntityTypeFactory(context, 1, entityType , null, "_REFERENCE", null);
                entityTypeFactory.Key = "_REFERENCE";
                entityTypeFactory.Name = "Pièces 2D";
                entityTypeFactory.DefaultDisplayKey = "_NAME";
                entityTypeFactory.ActAsEnvironment = false;
                
                {
                    IFieldDescription fieldDescription = new FieldDescription(context.Kernel.UnitSystem, true);
                    fieldDescription.Key = "AF_DPR_FOLDER";
                    fieldDescription.Name = "Répertoire DPR";
                    fieldDescription.Editable = FieldDescriptionEditableType.AllSection;
                    fieldDescription.Visible = FieldDescriptionVisibilityType.AllSection;
                    fieldDescription.Mandatory = false;
                    fieldDescription.FieldDescriptionType = FieldDescriptionType.String;
                    entityTypeFactory.EntityTypeAttributList.Add(fieldDescription);
                    
                }
                {
                    IFieldDescription fieldDescription = new FieldDescription(context.Kernel.UnitSystem, true);
                    fieldDescription.Key = "AF_LOCKED_REFERENCE";
                    fieldDescription.Name = "Bloquer Reference";
                    fieldDescription.Editable = FieldDescriptionEditableType.AllSection;
                    fieldDescription.Visible = FieldDescriptionVisibilityType.AllSection;
                    fieldDescription.Mandatory = false;
                    fieldDescription.FieldDescriptionType = FieldDescriptionType.Boolean;
                    fieldDescription.DefaultValue = false;
                    entityTypeFactory.EntityTypeAttributList.Add(fieldDescription);
                    
                }
                if (!entityTypeFactory.UpdateModel())
                {
                    foreach (ModelSettingError error in entityTypeFactory.ErrorList)
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
