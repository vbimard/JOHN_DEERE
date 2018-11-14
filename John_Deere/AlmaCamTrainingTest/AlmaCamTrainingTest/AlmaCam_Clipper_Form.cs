using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

using Wpm.Implement.Manager;
using Wpm.Implement.Processor;
using Wpm.Implement.ComponentEditor;  // ouverture de fenetres de selection almacam
using Wpm.Schema.Kernel;

//using System.ComponentModel;
//using Alma.BaseUI.Utils;
//dll personnalisées
using AF_JOHN_DEERE;
//using AF_ImportTools;
using System.IO;
using Microsoft.Win32;
using System.Diagnostics;
using Actcut.ResourceManager;
using Actcut.ActcutModelManager;

//suppression de wpm.schema.component.dll et remplacement par wpm.schema.componenteditor.dll 

namespace AlmaCamTrainingTest
{

    public partial class AlmaCam_Clipper_Form : Form
    {

        //initialisation des listes
        IContext _Context = null;
        
        string DbName = Alma_RegitryInfos.GetLastDataBase();

        public AlmaCam_Clipper_Form()
        {
            try
            {

                InitializeComponent();
            }
            catch (Exception e)
            {
                MessageBox.Show(e.Message);
            }

        }
        /// <summary>
        /// icone notification,
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        /// 

            
        private void Form1_Resize(object sender, EventArgs e)
        {

        }


        private void notifyIcon_MouseDoubleClick(object sender, MouseEventArgs e)
        {
            Show();
            this.WindowState = FormWindowState.Normal;
            notifyIcon.Visible = false;
        }

        /// <summary>
        /// form load
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void Form1_Load(object sender, EventArgs e)
        {


            

            //creation du model repository
            IModelsRepository modelsRepository = new ModelsRepository();
            _Context = modelsRepository.GetModelContext(DbName);  //nom de la base;
            int i = _Context.ModelsRepository.ModelList.Count();
            string infosPasserelle;
            
            infosPasserelle= DbName + "-P." + AF_JOHN_DEERE.JohnDeere_Param.getClipperDllVersion() + "-CAM." + AF_JOHN_DEERE.JohnDeere_Param.getAlmaCAMCompatibleVerion();
            this.Text = this.Name;
            this.InfosLabel.Text = infosPasserelle;
            this.Text = "Passerelle SAP : " + infosPasserelle;
        }

     
    

        //private void button4_Click(object sender, EventArgs e)
        private void ImportOF_Click(object sender, EventArgs e)
        {

            //import of
            //chargement de sparamètres
            // bool SansDt=false;

            JohnDeere_Param.GetlistParam(_Context);
            string csvImportPath = JohnDeere_Param.GetPath("IMPORT_CDA");
            //recuperation du nom de fichier
            string csvFileName = Path.GetFileNameWithoutExtension(csvImportPath);
            string csvDirectory = Path.GetDirectoryName(csvImportPath);
            string csvImportSandDt = csvDirectory + "\\" + csvFileName + "_SANSDT.csv";
          

            string dataModelstring = JohnDeere_Param.GetModelCA();


            using (JohnDeere_OF Import_OF = new JohnDeere_OF())
            {
                Import_OF.Import(_Context, csvImportPath, dataModelstring);
                //CahierAffaire.Import(_Context, csvImportSandDt, dataModelstring, true);
                //}

            }


        }


       

        private void button7_Click(object sender, EventArgs e)
        {
            JohnDeere_Param.GetlistParam(_Context);

            string csvImportPath = JohnDeere_Param.GetPath("IMPORT_CDA");
            ProcessStartInfo start = new ProcessStartInfo();
            ProcessStartInfo start_ca = new ProcessStartInfo();
            start_ca.Arguments = "OF " + csvImportPath;
            //start.FileName = @"C:\AlmaCAM\Bin\Clipper_Import.exe";
            start_ca.FileName = JohnDeere_Param.get_application1();
            //start.WindowStyle = ProcessWindowStyle.Normal;
            start_ca.CreateNoWindow = false;
            start_ca.UseShellExecute = true;
            string exename = JohnDeere_Param.get_application1();
            Process p = Process.Start(exename, "OF " + csvImportPath);


        }

        


       

        private void groupBox1_Enter(object sender, EventArgs e)
        {

        }

        private void button12_Click(object sender, EventArgs e)
        {
            //appel de la lib d'export des besoins ici
           

        }

        private void cAToolStripMenuItem_Click(object sender, EventArgs e)
        {

        }

        private void sheetRequiToolStripMenuItem_Click(object sender, EventArgs e)
        {
            //appel de la lib d'export des besoins ici
           

        }

        private void donnéesTechniquesToolStripMenuItem_Click(object sender, EventArgs e)
        {

           

        }


        private void stockToolStripMenuItem_Click(object sender, EventArgs e)
        {
            //import_stock
            
        }

        

        

        

        private void AfterClose_Click(object sender, EventArgs e)
        {
            //IEntity TO_CUT_nesting;
            AF_JOHN_DEERE.JohnDeere_DoOnAction_After_Cutting_end doonaction = new JohnDeere_DoOnAction_After_Cutting_end();

             //string stage = "_CLOSED_NESTING";
            string stage = "_CUT_SHEET";
            //creation du fichier de sortie
            //recupere les path
            JohnDeere_Param.GetlistParam(_Context);
            IEntitySelector Entityselector = null;

            Entityselector = new EntitySelector();

            //entity type pointe sur la list d'objet du model
            Entityselector.Init(_Context, _Context.Kernel.GetEntityType(stage));
            Entityselector.MultiSelect = true;

            if (Entityselector.ShowDialog() == System.Windows.Forms.DialogResult.OK)
            {
              
                    doonaction.execute(Entityselector.SelectedEntity);

            }
            

            

        }


        private IEntity SelectEntity(string SelectionType)
        {
            IEntity selectedentity = null;
            IEntitySelector xselector = null;
            xselector = new EntitySelector();

            xselector.Init(_Context, _Context.Kernel.GetEntityType(SelectionType));//"_TO_CUT_NESTING"
            xselector.MultiSelect = true;

            if (xselector.ShowDialog() == System.Windows.Forms.DialogResult.OK)
            {
                foreach (IEntity xentity in xselector.SelectedEntity)
                {
                    selectedentity = xentity;

                }

            }

            return selectedentity;

        }
        private IEntity SelectEntity_If_Not_Exported(string entitystage)
        {
            IEntity selectedentity = null;
            IEntitySelector xselector = null;
            xselector = new EntitySelector();



            //GPAO_Exported
            //_Context.EntityManager.GetEntityList("_CLOSED_NESTING"entitystage, "GPAO_Exported", ConditionOperator.Equal, null);
            IEntityList GPAO_Exported_filter = _Context.EntityManager.GetEntityList(entitystage, "GPAO_Exported", ConditionOperator.Equal, null);
            GPAO_Exported_filter.Fill(false);

            _Context.EntityManager.GetExtendedEntityList(entitystage, GPAO_Exported_filter);
            IDynamicExtendedEntityList exported_nestings = _Context.EntityManager.GetDynamicExtendedEntityList(entitystage, GPAO_Exported_filter);
            exported_nestings.Fill(false);

            if (exported_nestings.Count == 0)
            {
                //_Context.EntityManager.GetEntityList("_CLOSED_NESTING"entitystage, "GPAO_Exported", ConditionOperator.Equal, false);
                GPAO_Exported_filter = _Context.EntityManager.GetEntityList(entitystage, "GPAO_Exported", ConditionOperator.Equal, false);
                GPAO_Exported_filter.Fill(false);
               // _Context.EntityManager.GetExtendedEntityList("_CLOSED_NESTING"entitystage, GPAO_Exported_filter);
                exported_nestings = _Context.EntityManager.GetDynamicExtendedEntityList(entitystage, GPAO_Exported_filter);
                exported_nestings.Fill(false);
            }

            xselector.Init(_Context, exported_nestings);//"_TO_CUT_NESTING"
            xselector.MultiSelect = true;

            if (xselector.ShowDialog() == System.Windows.Forms.DialogResult.OK)
            {
                foreach (IEntity xentity in xselector.SelectedEntity)
                {
                    selectedentity = xentity;

                }

            }

            return selectedentity;
        
    
              /*
            else
            {
                xselector.Init(_Context, _Context.Kernel.GetEntityType(entitytype));//"_TO_CUT_NESTING"
                xselector.MultiSelect = true;

            }*/

          

            

        }

        private void button13_Click(object sender, EventArgs e)
        {

        }

        private void button2_Click(object sender, EventArgs e)
        {

        }

        private void quitToolStripMenuItem_Click(object sender, EventArgs e)
        {
           
        }

       

        private void quitToolStripMenuItem1_Click(object sender, EventArgs e)
        {
            
        }

        private void quitToolStripMenuItem2_Click(object sender, EventArgs e)
        {
            this.Close();
        }


        /// <summary>
        /// recupere la reference du premier placement selectionné dans les placements a copuper
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void AfterSend_Click_1(object sender, EventArgs e)
        {


 



            AF_JOHN_DEERE.JohnDeere_DoOnAction_AfterSendToWorkshop doonaction = new JohnDeere_DoOnAction_AfterSendToWorkshop();
            string stage = "_TO_CUT_NESTING";
            JohnDeere_Param.GetlistParam(_Context);
            IEntitySelector nestingselector = null;
            nestingselector = new EntitySelector();

            //entity type pointe sur la list d'objet du model
            nestingselector.Init(_Context, _Context.Kernel.GetEntityType(stage));
            nestingselector.MultiSelect = true;

            if (nestingselector.ShowDialog() == System.Windows.Forms.DialogResult.OK)
            {


                foreach (IEntity selection in nestingselector.SelectedEntity)
                {
                    string filename = selection.GetFieldValueAsString("_REFERENCE");
                    string material = selection.GetFieldValueAsEntity("_MATERIAL").GetFieldValueAsString("_NAME");

                    IEntityList nestinglist = _Context.EntityManager.GetEntityList(stage, "_REFERENCE", ConditionOperator.Equal, filename);
                    nestinglist.Fill(false);
                    doonaction.execute(nestinglist);
                }
               
               

            }

               






            }

        
        }
    }















