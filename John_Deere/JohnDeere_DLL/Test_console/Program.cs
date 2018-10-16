//system
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Runtime.InteropServices;
using System.Diagnostics;
//psecific windows
using System.IO;
using Microsoft.Win32;

//dll almacam
using Wpm.Implement.Manager;
using Wpm.Implement.Processor;
using Wpm.Implement.ComponentEditor;  // ouverture de fenetres de selection almacam
using Wpm.Schema.Kernel;
using Actcut.ActcutModelManager;
using Actcut.NestingManager;
using Actcut.ResourceManager;
using Actcut.ResourceModel;
//dll personnalisées



using System.Windows.Forms;
using AF_JOHN_DEERE;

namespace AlmaCamUser1
{
    class Program
    {
               //initialisation des listes
        /// <summary>
        /// aucun log dans ce programme, seulement des messages informels
        /// Si pas de fichier detecté   , alors on annule l'import
        /// Si pas de type detecté      , alors on annule l'import
        /// </summary>
        /// <param name="args">arg  0 : type d'import, arg 1 chemin vers le fichier d'import
        /// il n'y a pas d'obligation a envoyer le chemin car ce dernier peut etre fourni par l'application almacam
        /// </param>
        static void Main(string[] args)
        {
               
            IContext _clipper_Context = null;
            
            string TypeImport = "";
            //string fulpathname = "";            
            string DbName = AF_JOHN_DEERE.Alma_RegitryInfos.GetLastDataBase();
            AF_JOHN_DEERE.Alma_Log.Create_Log();
            
            


            using (EventLog eventLog = new EventLog("Application"))
            {


                            

                ModelsRepository clipper_modelsRepository = new ModelsRepository();
                string csvImportPath = null;
                _clipper_Context = clipper_modelsRepository.GetModelContext(DbName);  //nom de la base;
                int i = _clipper_Context.ModelsRepository.ModelList.Count();
                JohnDeere_Param.GetlistParam(_clipper_Context);
                if (args.Length==0)  { 

                    /* dans ce cas on recupere les arguments de la base directement*/
                    /* on force l'import of*/
                    TypeImport = "OF";
                    /**/

                }
                else {//sinon on recupere le paramètre du type d'import
                    TypeImport = args[0].ToUpper().ToString();}
               

                 {
                    switch (TypeImport)
                    {
                        //fullpath name
                        case "STOCK":
                            //import stock
                         

                            break;

                        case "STOCK_PURGE":
                            //puge de tous les elements du stock

                           
                            break;

                        case "OF":
                         
                            clipper_modelsRepository = new ModelsRepository();
                            //import of                          
                            

                            if (args.Length==0 || args.Length == 1)
                            {
                                csvImportPath = JohnDeere_Param.GetPath("IMPORT_CDA");
                            }
                            else
                            {
                                csvImportPath = args[1].ToUpper().ToString();

                            }
                            
                            string of_dataModelstring = JohnDeere_Param.GetModelCA();

                           
                            using (JohnDeere_OF CahierAffaire = new JohnDeere_OF())
                            {
                                
                                CahierAffaire.Import(_clipper_Context, csvImportPath, of_dataModelstring);
                                
                                       
                            }


                         
                            break;

                    }



                }
            }
            


            }
        }
    }

