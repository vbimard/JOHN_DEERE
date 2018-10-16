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
            #region Pièces à produire
            
            {
                IEntityType entityType = context.Kernel.GetEntityType("_TO_PRODUCE_REFERENCE");
                IEntityTypeFactory entityTypeFactory = new EntityTypeFactory(context, 1, entityType , null, "_ACTCUT", null);
                entityTypeFactory.Key = "_TO_PRODUCE_REFERENCE";
                entityTypeFactory.Name = "Pièces à produire";
                entityTypeFactory.DefaultDisplayKey = "_NAME";
                entityTypeFactory.ActAsEnvironment = false;
                
                {
                    IFieldDescription fieldDescription = new FieldDescription(context.Kernel.UnitSystem, true);
                    fieldDescription.Key = "AF_PT";
                    fieldDescription.Name = "AF Poste";
                    fieldDescription.Editable = FieldDescriptionEditableType.AllSection;
                    fieldDescription.Visible = FieldDescriptionVisibilityType.AllSection;
                    fieldDescription.Mandatory = false;
                    fieldDescription.FieldDescriptionType = FieldDescriptionType.String;
                    entityTypeFactory.EntityTypeAttributList.Add(fieldDescription);
                    
                }
                {
                    IFieldDescription fieldDescription = new FieldDescription(context.Kernel.UnitSystem, true);
                    fieldDescription.Key = "AF_COMPOSANT";
                    fieldDescription.Name = "AF Composants";
                    fieldDescription.Editable = FieldDescriptionEditableType.AllSection;
                    fieldDescription.Visible = FieldDescriptionVisibilityType.AllSection;
                    fieldDescription.Mandatory = false;
                    fieldDescription.FieldDescriptionType = FieldDescriptionType.String;
                    entityTypeFactory.EntityTypeAttributList.Add(fieldDescription);
                    
                }
                {
                    IFieldDescription fieldDescription = new FieldDescription(context.Kernel.UnitSystem, true);
                    fieldDescription.Key = "AF_ORDRE";
                    fieldDescription.Name = "AF Ordre";
                    fieldDescription.Editable = FieldDescriptionEditableType.AllSection;
                    fieldDescription.Visible = FieldDescriptionVisibilityType.AllSection;
                    fieldDescription.Mandatory = false;
                    fieldDescription.FieldDescriptionType = FieldDescriptionType.String;
                    entityTypeFactory.EntityTypeAttributList.Add(fieldDescription);
                    
                }
                {
                    IFieldDescription fieldDescription = new FieldDescription(context.Kernel.UnitSystem, true);
                    fieldDescription.Key = "AF_BESOIN";
                    fieldDescription.Name = "AF Besoin";
                    fieldDescription.Editable = FieldDescriptionEditableType.AllSection;
                    fieldDescription.Visible = FieldDescriptionVisibilityType.AllSection;
                    fieldDescription.Mandatory = false;
                    fieldDescription.FieldDescriptionType = FieldDescriptionType.String;
                    entityTypeFactory.EntityTypeAttributList.Add(fieldDescription);
                    
                }
                {
                    IFieldDescription fieldDescription = new FieldDescription(context.Kernel.UnitSystem, true);
                    fieldDescription.Key = "AF_DESTINATION";
                    fieldDescription.Name = "AF Destination";
                    fieldDescription.Editable = FieldDescriptionEditableType.AllSection;
                    fieldDescription.Visible = FieldDescriptionVisibilityType.AllSection;
                    fieldDescription.Mandatory = false;
                    fieldDescription.FieldDescriptionType = FieldDescriptionType.String;
                    entityTypeFactory.EntityTypeAttributList.Add(fieldDescription);
                    
                }
                {
                    IFieldDescription fieldDescription = new FieldDescription(context.Kernel.UnitSystem, true);
                    fieldDescription.Key = "AF_BROSS_EBAV";
                    fieldDescription.Name = "AF Brossage / Ebavurage";
                    fieldDescription.Editable = FieldDescriptionEditableType.AllSection;
                    fieldDescription.Visible = FieldDescriptionVisibilityType.AllSection;
                    fieldDescription.Mandatory = false;
                    fieldDescription.FieldDescriptionType = FieldDescriptionType.String;
                    entityTypeFactory.EntityTypeAttributList.Add(fieldDescription);
                    
                }
                {
                    IFieldDescription fieldDescription = new FieldDescription(context.Kernel.UnitSystem, true);
                    fieldDescription.Key = "AF_KB_MRP";
                    fieldDescription.Name = "AF KB / MRP";
                    fieldDescription.Editable = FieldDescriptionEditableType.AllSection;
                    fieldDescription.Visible = FieldDescriptionVisibilityType.AllSection;
                    fieldDescription.Mandatory = false;
                    fieldDescription.FieldDescriptionType = FieldDescriptionType.String;
                    entityTypeFactory.EntityTypeAttributList.Add(fieldDescription);
                    
                }
                {
                    IFieldDescription fieldDescription = new FieldDescription(context.Kernel.UnitSystem, true);
                    fieldDescription.Key = "AF_AVANCE_RETARD";
                    fieldDescription.Name = "AF avance retard";
                    fieldDescription.Editable = FieldDescriptionEditableType.AllSection;
                    fieldDescription.Visible = FieldDescriptionVisibilityType.AllSection;
                    fieldDescription.Mandatory = false;
                    fieldDescription.FieldDescriptionType = FieldDescriptionType.String;
                    entityTypeFactory.EntityTypeAttributList.Add(fieldDescription);
                    
                }
                {
                    IFieldDescription fieldDescription = new FieldDescription(context.Kernel.UnitSystem, true);
                    fieldDescription.Key = "AF_CONTENANT";
                    fieldDescription.Name = "AF Contenant";
                    fieldDescription.Editable = FieldDescriptionEditableType.AllSection;
                    fieldDescription.Visible = FieldDescriptionVisibilityType.AllSection;
                    fieldDescription.Mandatory = false;
                    fieldDescription.FieldDescriptionType = FieldDescriptionType.String;
                    entityTypeFactory.EntityTypeAttributList.Add(fieldDescription);
                    
                }
                {
                    IFieldDescription fieldDescription = new FieldDescription(context.Kernel.UnitSystem, true);
                    fieldDescription.Key = "AF_CONTENANT_TH";
                    fieldDescription.Name = "AF Contenant Th";
                    fieldDescription.Editable = FieldDescriptionEditableType.AllSection;
                    fieldDescription.Visible = FieldDescriptionVisibilityType.AllSection;
                    fieldDescription.Mandatory = false;
                    fieldDescription.FieldDescriptionType = FieldDescriptionType.String;
                    entityTypeFactory.EntityTypeAttributList.Add(fieldDescription);
                    
                }
                {
                    IFieldDescription fieldDescription = new FieldDescription(context.Kernel.UnitSystem, true);
                    fieldDescription.Key = "AF_TIME_STAMP";
                    fieldDescription.Name = "AF_TIMESTAMP";
                    fieldDescription.Editable = FieldDescriptionEditableType.AllSection;
                    fieldDescription.Visible = FieldDescriptionVisibilityType.AllSection;
                    fieldDescription.Mandatory = false;
                    fieldDescription.FieldDescriptionType = FieldDescriptionType.String;
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
