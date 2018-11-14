
//using AF_JOHN_DEERE_Dll;
//
//actcut
using Actcut.ActcutModelManager;
using Actcut.NestingManager;
using Actcut.ResourceManager;
using Actcut.ResourceModel;
//using AF_ImportTools_JD;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Drawing;
using System.Drawing.Imaging;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Resources;
using System.Security.AccessControl;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading;
using System.Threading.Tasks;
using System.Windows.Forms;
//almacam
using Wpm.Implement.ComponentEditor;  // ouverture de fenetres de selection almacam
using Wpm.Implement.Manager;
using Wpm.Implement.Processor;
using Wpm.Schema.Kernel;
using static AF_JOHN_DEERE.Util;

//



namespace AF_JOHN_DEERE
//namespace Import_GP
#region commande_processor

{
    /// <summary>
    /// automatisme BO : outils necessaire pour l'envoie d'infos au service windoxs
    /// </summary>
    public class Automation_Tools : IDisposable
    {
        public void Dispose()
        {
            //Dispose(true);
            GC.SuppressFinalize(this);
        }
        /// <summary>
        /// OBLIGATION D EXECUTER EN ADMIN
        /// recupere la liste des placements exportes (confition exported)
        /// recupere la liste des fichier dans les dossiers gpao
        /// cloture si le fichier n'est plus present
        /// ce code utilise les log windwos: pour supprimer les log windows
        /// Get-EventLog -List
        /// C:\> Remove-EventLog -LogName "MyLog"
        /// Remove-EventLog -Source "MyApp"
        /// </summary>
        /// <returns>true/false</returns>
        public bool Automatic_Nestings_Close(IContext Contextlocal, string Stage, WindowsLog log)
        {
            string message = ""; //message de suivi pour les log
            try
            {

                JohnDeere_Param.GetlistParam(Contextlocal);
                bool rst = false;
                //string stage = "_TO_CUT_NESTING";
                IEntityList nestings_list = null;
                IEntity current_nesting = null;
                System.Console.WriteLine("connecté à" + Contextlocal.Connection.DatabaseName + " Pour cloture");
                System.Console.WriteLine("fermeture des placements " + Stage);

                nestings_list = Contextlocal.EntityManager.GetEntityList(Stage, SimplifiedMethods.Get_Marqued_FieldName(Stage), ConditionOperator.Equal, true);
                nestings_list.Fill(false);


                if (nestings_list != null && nestings_list.Count() > 0)
                {
                    foreach (IEntity nesting in nestings_list)
                    {
                        string nesting_name;
                        string technology;
                        string path_to_file;
                        //string msgstart;
                        //On initialise le message
                        message = "";
                        //IEntityList nestings_to_close;
                        IEntity nesting_to_close;


                        nesting_to_close = nesting;

                        //get the nesting name 
                        nesting_name = current_nesting.GetFieldValueAsString("_NAME");
                        //get the technology--> get the folder
                        System.Console.WriteLine("----------------------------------------");
                        System.Console.WriteLine("Placement: " + nesting_name);
                        message += "Clôture de: " + nesting_name;
                        technology = Machine_Info.GetNestingTechnologyName(ref Contextlocal, ref current_nesting);
                        System.Console.WriteLine(nesting_name + ": Technologie: " + technology);
                        message += "techno detected:" + technology + "\n";
                        //technology = "";
                        //get the filename
                        @path_to_file = JohnDeere_Param.GetPath("Export_GPAO") + "\\" + technology + "\\" + nesting_name + ".txt";
                        message += "gpao file: " + @path_to_file + "\n";

                        if (!File.Exists(@path_to_file))
                        {//on cloture
                         //on reconstruit une liste des placaments                     
                         //nesting_to_close = current_nesting; //nestings_list.Where(x => x.GetFieldValueAsString("_NAME") == nesting_name).FirstOrDefault();
                            message += "fermeture du placement " + nesting_to_close.GetFieldValueAsString("_NAME") + "\n";
                            SimplifiedMethods.CloseNesting(Contextlocal, nesting_to_close);
                            log.WriteLogSuccess("Synthèse :\n " + message + "\n" + message);
                            System.Console.WriteLine("placement  " + nesting_to_close.GetFieldValueAsString("_NAME") + " fermé");

                        }
                        else
                        {

                            message += "le fichier de placement a été detecté -le placement ne sera pas cloturé " + "\n";
                        }
                        //suppresse the file
                        log.WriteLogEvent("Synthèse :\n " + message + "\n");
                    }
                }

                System.Console.Out.Close();
                return rst;

            }
            catch (Exception ie)
            {

                log.WriteLogWarningEvent("Probleme rencontré log de la cloture des placements :\n " + message);
                log.WriteLogWarningEvent("details :\n " + ie.Message);
                //System.Console.WriteLine("Erreur à la fermeture du placement " +ie.Message);
                //System.Console.ReadLine() ;
                return false;
            }
        }
        /// <summary>
        /// recupere la liste des placements exportes (confition exported)
        /// recupere la liste des fichier dans les dossiers gpao
        /// cloture si le fichier n'est plus present
        ///utilise les log standard alma 
        /// C:\> Remove-EventLog -LogName "MyLog"
        /// Remove-EventLog -Source "MyApp"
        /// </summary>
        /// <returns></returns>
        public bool Automatic_Nestings_Close(IContext Contextlocal, string Stage)
        {
            string message = ""; //message de suivi pour les log
            try
            {

                JohnDeere_Param.GetlistParam(Contextlocal);


                Alma_Log.Write_Log("Parametre recuperés");

                bool rst = false;
                //string stage = "_TO_CUT_NESTING";
                IEntityList nestings_list = null;
                IEntity current_nesting = null;

                Alma_Log.Write_Log("connecté à" + Contextlocal.Connection.DatabaseName + " Pour cloture");
                System.Console.WriteLine("connecté à" + Contextlocal.Connection.DatabaseName + " Pour cloture");
                Alma_Log.Write_Log("fermeture des placement " + Stage);
                System.Console.WriteLine("fermeture des placement " + Stage);

                nestings_list = Contextlocal.EntityManager.GetEntityList(Stage, Stage + "_GPAO_Exported", ConditionOperator.Equal, true);
                nestings_list.Fill(false);

                if (nestings_list != null && nestings_list.Count() > 0)
                {
                    foreach (IEntity nesting in nestings_list)
                    {
                        string nesting_name;
                        string technology;
                        string path_to_file;
                        //string msgstart;
                        //On initialise le message
                        message = "";
                        //IEntityList nestings_to_close;
                        IEntity nesting_to_close;

                        //nesting_to_close= nestings_list.Where(x=>x.GetFieldValueAsString("_NAME")== nesting.GetFieldValueAsString("_NAME"));
                        current_nesting = nesting;

                        //get the nesting name 
                        nesting_name = current_nesting.GetFieldValueAsString("_NAME");
                        //get the technology--> get the folder
                        message += "closing :" + nesting_name;
                        technology = Machine_Info.GetNestingTechnologyName(ref Contextlocal, ref current_nesting);
                        message += "techno detected:" + technology;
                        //technology = "";
                        //get the filename
                        @path_to_file = JohnDeere_Param.GetPath("Export_GPAO") + "\\" + technology + "\\" + technology + "\\" + nesting_name + ".txt";
                        message += "gapo file: " + @path_to_file;
                        if (!File.Exists(@path_to_file))
                        {//on cloture
                         //on reconstruit une liste des placaments                     
                            nesting_to_close = nestings_list.Where(x => x.GetFieldValueAsString("_NAME") == nesting_name).FirstOrDefault();
                            Alma_Log.Write_Log("fermeture du placement " + nesting_to_close.GetFieldValueAsString("_NAME"));
                            System.Console.WriteLine("fermeture du placement " + nesting_to_close.GetFieldValueAsString("_NAME"));

                        }
                        //suppresse the file

                    }
                }

                System.Console.Out.Close();
                return rst;

            }
            catch (Exception ie)
            {
                Alma_Log.Write_Log("Probleme rencontré log de la cloture des placements :\n " + message);
                Alma_Log.Write_Log("details :\n " + ie.Message);

                return false;
            }
        }
    }

    /// <summary>
    /// les commandes processor designent les boutons d'action integerés dans l'interface almaquote 
    /// cette classe lance l'application deporter d'import du stock
    /// pour le moment cette application est l'executable suivant
    /// C:\AlmaCAM\Bin\AlmaCamTrainingTest.exe
    /// simplecommandeprocessor (commande globale)
    /// </summary>
    public class JohnDeereIE : SimpleCommandProcessor
    {
        //public IContext contextlocal = null;
        public override bool Execute()
        {

            //MessageBox.Show("vincent");
            ProcessStartInfo start_test = new ProcessStartInfo
            {
                Arguments = "",

                FileName = @"C:\AlmaCAM\Bin\AlmaCamTrainingTest.exe"
            };

            Process.Start(start_test);

            return base.Execute();
        }
    }


    public static class JohnDeereExit
    {

        public static void Close()
        {

            Alma_Log.Final_Open_Log();
            Application.Exit();


        }


    }


    //PARAMETRES
    ///////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
    ///////////////////////////////////////////////////////IMPORT///////////////OF////////////////////////////////////////////////////////////////////////////////////////////////
    ///////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
    ///////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////

    /// <summary>
    /// la classe JohnDeere_param recupere les paramètres de almacam dans les options de la passerelle (section piece a produire) 
    /// comme la lecture des paramètres est indispensable, cette classe verifie aussi la presence sdes dossier clipper , le nom de la base ainsi
    /// que la compatibilité  de almacam 
    /// </summary>

    public static class JohnDeere_Param
    {
        static Dictionary<string, object> Parameters_Dictionnary; // liste des path et des types pour le format du fichier de stock et des of

        // public static IContext context;
        static JohnDeere_Param()
        {
            Parameters_Dictionnary = new Dictionary<string, object>();


        }

        /// <summary>
        /// recuperation des path clipper+fichier.csv (echange csv) ou autre
        /// </summary>
        /// //exemple "H:\tutu\toto\cahieraffaire.csv"
        /// <param name="context"> contexte </param>
        public static Boolean GetlistParam(IContext context)
        {

            try
            {
                string parametre_name;

                Parameters_Dictionnary.Clear();
                /**/
                //ListPath.Add("EXPORT_GP", context.ParameterSetManager.GetParameterValue("EXPORT_GP", "DIR").GetValueAsString());
                //chemin import cahier affaire
                //chemin import
                parametre_name = "IMPORT_CDA";
                Alma_Log.Info("recuperation du parametre " + parametre_name, "GetlistParam");
                Parameters_Dictionnary.Add(parametre_name, context.ParameterSetManager.GetParameterValue("AF_JOHN_DEERE_DLL", "IMPORT_CDA").GetValueAsString());
                parametre_name = "Export_GPAO";
                Alma_Log.Info("recuperation du parametre " + parametre_name, "GetlistParam");
                Parameters_Dictionnary.Add(parametre_name, context.ParameterSetManager.GetParameterValue("AF_JOHN_DEERE_DLL", "EXPORT_Rp").GetValueAsString());
                //description import
                parametre_name = "IMPORT_AUTO";
                Alma_Log.Info("recuperation du parametre " + parametre_name, "GetlistParam");
                Parameters_Dictionnary.Add(parametre_name, context.ParameterSetManager.GetParameterValue("AF_JOHN_DEERE_DLL", "IMPORT_AUTO").GetValueAsBoolean());                /**/
                                                                                                                                                                                  /**/
                parametre_name = "MODEL_CA";
                Alma_Log.Info("recuperation du parametre " + parametre_name, "GetlistParam");
                Parameters_Dictionnary.Add(parametre_name, context.ParameterSetManager.GetParameterValue("AF_JOHN_DEERE_DLL", "MODEL_CA").GetValueAsString());
                parametre_name = "MODEL_PATH";
                Alma_Log.Info("recuperation du parametre " + parametre_name, "GetlistParam");
                Parameters_Dictionnary.Add(parametre_name, context.ParameterSetManager.GetParameterValue("AF_JOHN_DEERE_DLL", "MODEL_PATH").GetValueAsString());
                /*
                parametre_name = "APPLICATION1";
                Alma_Log.Info("recuperation du parametre " + parametre_name, "GetlistParam");
                Parameters_Dictionnary.Add(parametre_name, context.ParameterSetManager.GetParameterValue("JohnDeere_DLL", "APPLICATION1").GetValueAsString());
                */
                //log
                parametre_name = "VERBOSE_LOG";
                Alma_Log.Info("recuperation du parametre " + parametre_name, "GetlistParam");
                Parameters_Dictionnary.Add(parametre_name, context.ParameterSetManager.GetParameterValue("AF_JOHN_DEERE_DLL", "VERBOSE_LOG").GetValueAsBoolean());
                //nom mahine clipper
                /*
                parametre_name = "JohnDeere_MACHINE_CF";
                Alma_Log.Info("recuperation du parametre " + parametre_name, "GetlistParam");
                Parameters_Dictionnary.Add(parametre_name, context.ParameterSetManager.GetParameterValue("JohnDeere_DLL", "JohnDeere_MACHINE_CF").GetValueAsString());
                */
                /*parametres de sorties*/
                parametre_name = "STRING_FORMAT_DOUBLE";
                Alma_Log.Info("recuperation du parametre " + parametre_name, "GetlistParam");
                Parameters_Dictionnary.Add(parametre_name, "{0:0.00###}");

                //verification des chemins


                //verification des path
                Alma_Log.Info("verification de l'existance des path ", "GetlistParam");
                if (checkClipperFolderExists() == false) { throw new System.ApplicationException("Certains chemin d'echanges de la Passerelle AlmaCam-SAGE ne sont pas accessibles"); };
                ///string AlmaCAmVersion = Directory.GetCurrentDirectory().ToString() + Parameters_Dictionnary.Values("");///
                if (ckeckCompatibilityVersion() == false) { throw new System.ApplicationException("Version de la Dll SAGE n'est pas validée pour cette version d'AlmaCam"); }

                return true;
            }
            catch (KeyNotFoundException ex)
            {
                Alma_Log.Error(ex, "CETTE BASE NE SEMBLE PAS ETRE CONFIGUREE POUR SAGE !!! ");
                Alma_Log.Error(ex, "Veuiller verifier la configuration des paramètres de l'import SAGE (nom et id des champs....)");
                MessageBox.Show(Alma_RegitryInfos.GetLastDataBase() + " :CETTE BASE NE SEMBLE PAS ETRE CONFIGUREE POUR SAGE !!! \r\n " +
                "Veuillez verifier la base selectionnées pour l'ouverture d'AlmaCam");
                //on sort
                //ClipperExit.Close();
                return false;
            }
            catch (System.ApplicationException exVersion)
            {
                MessageBox.Show(exVersion.Message);
                Alma_Log.Error(exVersion, "Version incompatible ou mauvaise configuration de la base almacam");
                //ClipperExit.Close();
                return false;
            }

            catch (System.IO.DirectoryNotFoundException exFolder)
            {
                MessageBox.Show(exFolder.Message);
                Alma_Log.Error(exFolder, "L'un des dossier de d'echange n'existe pas");
                return false;
                //ClipperExit.Close();
            }

            //finally { return false; }


            //MessageBox.Show(e.Message+"/r/n Version incompatible ou mauvaise configuration de la base almacam"); }
        }

        /// <summary>
        /// retourn la valeur de la clé recherché dans les paramètres
        /// </summary>
        /// <typeparam name="T">type générique </typeparam>
        /// <param name="context">contexte </param>
        /// <param name="PathVariable">clé a rechercher</param>
        /// <returns></returns>
        public static T GetParam<T>(this IDictionary<string, object> dic, string key)
        {
            if (Parameters_Dictionnary.ContainsKey(key))
            {
                return (T)dic[key];
            }
            else { return default(T); }
        }

        /// <summary>
        /// retourne un chemin windows type string
        /// </summary>
        /// <param name="key">nom de la clé dans le dictionnaire</param>
        /// <returns>chemin windows : type c:\actcut...</returns>
        public static string GetPath(string key)
        {

            //GetlistParam(context);//    Alma_Log.Info("verification de la clé " + key, "GetPath");
            try
            {
                // if (Parameters_Dictionnary.ContainsKey(key))
                {
                    object path;
                    Parameters_Dictionnary.TryGetValue(key, out path);
                    return path.ToString();
                }
            }
            catch (Exception ie)
            {
                Alma_Log.Info("impossible de trouver la clé  " + key, "GetPath");
                Alma_Log.Info("impossible de trouver la clé  " + ie.Message, "GetPath");
                return "Undef";
            }


        }
        /// <summary>
        /// recupere la case a cocher d'automation
        /// </summary>
        /// <returns>boolean true/false</returns>
        public static bool IsAutomatiqueImport()
        {
            string key = "IMPORT_AUTO";
            //GetlistParam(context);//
            if (Parameters_Dictionnary.ContainsKey(key)) { return (bool)Parameters_Dictionnary[key]; } else { return false; }

        }

        /// <summary>
        /// recupere
        /// </summary>
        /// <returns>un model de fichier csv sous la forme d'une liste de champs 
        /// champs : numero du champ dans #nom du champs  dans almacam#Type  -> plus tard si besoin numero du champ dans #nom du champs  dans almacam#Type  # taille max
        /// 0#NAME#string;1#AFFAIRE#string;2#THICKNESS#string;3#MATERIAL_CLIPPER#string;4#CENTREFRAIS#string;5#TECHNOLOGIE#string;6#FAMILY#string;7#IDLNROUT#string;8#CENTREFRAISSUIV#string;9#CUSTOMER#string;10#PART_INITIAL_QUANTITY#double;11#QUANTITY#double;12#ECOQTY#double;13#STARTDATE#date;14#ENDDATE#date;15#PLAN#string;....
        /// 
        /// </returns>
        public static string GetModelCA()
        {
            string key = "MODEL_CA";
            //GetlistParam(context);//
            if (Parameters_Dictionnary.ContainsKey(key)) { return (string)Parameters_Dictionnary[key]; } else { return "Undef  model CA"; }

        }


        public static string GetModelDM()
        {
            string key = "MODEL_DM";
            //GetlistParam(context);//
            if (Parameters_Dictionnary.ContainsKey(key)) { return (string)Parameters_Dictionnary[key]; } else { return "Undef  model  DM"; }

        }


        public static string GetModelPATH()
        {
            string key = "MODEL_PATH";
            //GetlistParam(context);//
            if (Parameters_Dictionnary.ContainsKey(key)) { return (string)Parameters_Dictionnary[key]; } else { return "Undef model PATH"; }

        }


        public static string get_string_format_double()
        {

            string key = "STRING_FORMAT_DOUBLE";
            //GetlistParam(context);//
            if (Parameters_Dictionnary.ContainsKey(key)) { return (string)Parameters_Dictionnary[key]; } else { return "Undef STRING_FORMAT_DOUBLE"; }

        }

        public static string get_application1()
        {
            string key = "APPLICATION1";
            //GetlistParam(context);//
            if (Parameters_Dictionnary.ContainsKey(key)) { return (string)@Parameters_Dictionnary[key]; } else { return "Undef model PATH"; }
        }

        public static bool getVerbose_Log()
        {
            string key = "VERBOSE_LOG"; //log verbeux
            if (Parameters_Dictionnary.ContainsKey(key)) { return (bool)Parameters_Dictionnary[key]; } else { return true; }
        }

        public static string get_JohnDeere_Machine_Cf()
        {
            string key = "JohnDeere_MACHINE_CF"; //log verbeux
            //GetlistParam(context);//
            if (Parameters_Dictionnary.ContainsKey(key)) { return (string)Parameters_Dictionnary[key]; } else { return "Undef SAGE machine"; }

        }


        public static string get_AlmaCamEditorName()
        {
            string key = "ALMACAM_EDITOR_NAME"; //log verbeux
            //GetlistParam(context);//
            if (Parameters_Dictionnary.ContainsKey(key)) { return (string)Parameters_Dictionnary[key]; } else { return "Wpm.Implement.Editor.exe"; }

        }
        /// <summary>
        /// verification de l'existance des dossiers d'echange
        /// </summary>
        /// <returns>true si tous les dossier existent, false si ils n'existent pas</returns>
        public static Boolean checkClipperFolderExists()
        {
            try
            {
                bool rst;
                Alma_Log.Info("checking IMPORT_CDA", "checkSAGEFolderExists");
                string dossier = "IMPORT_CDA";
                //Directory.GetDirectories(Path.GetDirectoryName(JohnDeere_Param.GetPath("IMPORT_CDA")));
                rst=Directory.Exists(Path.GetDirectoryName(JohnDeere_Param.GetPath(dossier)));
                Alma_Log.Info("checking IMPORT_DM", "checkSAGEFolderExists");
                //Directory.GetDirectories(Path.GetDirectoryName(JohnDeere_Param.GetPath("Export_GPAO")));
                dossier = "Export_GPAO";
                rst = rst && Directory.Exists(Path.GetDirectoryName(JohnDeere_Param.GetPath(dossier)));
                Alma_Log.Info("checking EXPORT_DT", "checkClipperFolderExists");
               

                return rst;
            }
            catch (System.IO.IOException ie)   //.DirectoryNotFoundException ie)
            {
                Alma_Log.Error(ie, " les dossiers d'echange sont mal definis verifier : IMPORT_CDA, IMPORT_DM ,Export_GPAO, EXPORT_Dt ,_EXPORT_GP_DIRECTORY");
                return false;
                throw;
            }


        }


        /// <summary>
        /// verification sur le nom de la base de données
        /// car clipper a écrit en dur le nom de la base de données
        /// </summary>
        /// <returns>true si le nom de la base est correcte</returns>
        public static Boolean checkDatabasename(string actualdatabasename)
        {
            try
            {
                bool res = true;

                string lastopenDbname = Alma_RegitryInfos.GetLastDataBase();
                if (Alma_RegitryInfos.GetLastDataBase().Count() == 0)
                {

                    new Exception("Impossible de lire le nom de la dernière base ouvert dans le registre");
                    {
                        throw new Exception(AF_JOHN_DEERE_Dll.Properties.Resources.JohnDeere_Almacam_Database_Name.ToString() + ", la base :" + AF_JOHN_DEERE_Dll.Properties.Resources.JohnDeere_Almacam_Database_Name + " est introuvable");

                    }

                }

                else
                {

                    string working_db = AF_JOHN_DEERE_Dll.Properties.Resources.JohnDeere_Almacam_Database_Name.ToString();
                    Alma_Log.Write_Log("derniere base ouverte: " + working_db);
                    if (JohnDeere_Param.checkDatabasename(working_db) == false) { throw new Exception(Alma_RegitryInfos.GetLastDataBase() + ", le Nom de la base est incorrecte,  elle doit se nommer :" + AF_JOHN_DEERE_Dll.Properties.Resources.JohnDeere_Almacam_Database_Name); }


                }
                return res;
            }
            //catch (Exception e)
            catch (Exception ex)
            {
                Alma_Log.Write_Log(ex.Message);
                MessageBox.Show(ex.Message);
                return false;
                throw;
            }




        }


        /// <summary>
        /// retourne la verison de la dll clipperDll.dll
        /// </summary>
        /// <returns></returns>
        public static string getClipperDllVersion()
        {
            return Application.ProductVersion.ToString().Substring(0, 3);
        }

        /// <summary>
        /// retourn la version compatible almacam indiquee dans les ressources almacam
        /// </summary>
        /// <returns></returns>
        public static string getAlmaCAMCompatibleVerion()
        {
            return AF_JOHN_DEERE_Dll.Properties.Resources.Almcam_Version.ToString();

        }
        /// <summary>
        /// recupere le numero de version compatible. et le compare a la version de l'executable almacam
        /// la version est bloquée par une infos des ressources de la dll JohnDeere_dll
        /// </summary>
        /// <returns>true si le test est accepté</returns>
        public static Boolean ckeckCompatibilityVersion()
        {
            bool res = false;

            try
            {

                bool compatible = false;
                string versionalmacam;
                string almacameditorfullpath = Directory.GetCurrentDirectory().ToString() + "\\" + get_AlmaCamEditorName();
                string almacamCompatibleversion = AF_JOHN_DEERE_Dll.Properties.Resources.Almcam_Version.ToString();
                //get version//
                versionalmacam = FileVersionInfo.GetVersionInfo(almacameditorfullpath).ProductVersion.ToString();
                Alma_Log.Write_Log("version almacam detectée: " + versionalmacam);
                Alma_Log.Write_Log_Important("version almacam detectée: " + versionalmacam);
                foreach (string v in almacamCompatibleversion.Split(';'))
                {
                    if (versionalmacam.StartsWith(v))
                    {
                        compatible = true;
                        break;
                    };
                }


                ///

                if (compatible == true)
                {
                    Alma_Log.Write_Log("verion wpm.exe :" + versionalmacam + " version autoriséee pour cette dll " + almacamCompatibleversion);
                    Alma_Log.Write_Log("test ok");
                    res = true;
                }
                else
                {
                    Alma_Log.Write_Log("la librairie JohnDeeredll.dll version " + almacamCompatibleversion + " est  incompatible Almacam  " + versionalmacam);
                }

                return res;

            }
            catch
            {

                Alma_Log.Write_Log(": time tag:  ");
                return res;

            }

        }


    }


    /// <summary>
    /// recupere les of exportes de clipper
    /// </summary>
    public class JohnDeere_OF : IDisposable, IImport
    {
        string CsvImportPath = null;


        public void Dispose()
        {
            //Dispose(true);
            GC.SuppressFinalize(this);
        }






        /// <summary>
        /// 
        /// </summary>
        /// <param name="contextlocal"></param>
        /// <param name="line_dictionnary"></param>
        /// <param name="CentreFrais_Dictionnary"></param>
        /// <param name="reference_to_produce"></param>
        /// <param name="reference"></param>
        /// <param name="timetag"></param>
        /// <returns></returns>
        public bool CreateNewPartToProduce(IContext contextlocal, Dictionary<string, object> line_dictionnary, Dictionary<string, string> CentreFrais_Dictionnary, ref IEntity reference_to_produce, ref IEntity reference, string timetag)
        {
            bool result = false;

            try
            {
                //la piece ne contient pas de gamme
                //cas des pieces oranges : Pas de cf, pas de id_piece_cfao, on considere que c'est une piece orange--> on ne creer que la reference. 
                /*
                if ((Data_Model.ExistsInDictionnary("CENTREFRAIS") == false) && (Data_Model.ExistsInDictionnary("CENTREFRAIS") == false))
            {
                return false;
            }*/
                //string referenceName = null;
                //Boolean need_prep = false;
                //int index_extension = 0;  //> 0 si ;emf;dpr detectée
                PartInfo machinable_part_infos = null; //infos de machinabe part

                //creation de la nouvelle reference
                reference_to_produce = contextlocal.EntityManager.CreateEntity("_TO_PRODUCE_REFERENCE");
                //recuperation et assignaton de la machine si elle existe
                //string machine_name = "";
                //Data_Model.ExistsInDictionnary(line_dictionnary["CENTREFRAIS"].ToString(), ref CentreFrais_Dictionnary);
                //lecture des part infos (optionnel) car le get reference fait le travail                 
                machinable_part_infos = new PartInfo();
                //bool fastmode = true;
                //bool result = false;
                //ecriture du time tag
                reference_to_produce.SetFieldValue("AF_TIME_STAMP", timetag.Replace("_", ""));
                reference_to_produce.SetFieldValue("_REFERENCE", reference.Id32);
                //reference_to_produce.SetFieldValue("NEED_PREP", need_prep);


                Update_Part_Item(contextlocal, ref reference_to_produce, ref line_dictionnary);
                Alma_Log.Write_Log(MethodBase.GetCurrentMethod().Name + "update infos succeed");
                reference_to_produce.Save();
                line_dictionnary.Clear();

                return result;
            }
            catch { return result; }
        }



        /// <summary>
        /// Recupere la liste de toutes les machines sous la forme litterale "nom machine" "centre de frais"
        /// </summary>
        /// <param name="contextlocal">context local</param>
        /// <param name="JohnDeere_Machine">entité machine clipper</param>
        /// <param name="JohnDeere_Centre_Frais">entité centre de frais clipper</param>
        /// <returns></returns>
        public Boolean Get_JohnDeere_Machine(IContext contextlocal, out IEntity JohnDeere_Machine, out IEntity JohnDeere_Centre_Frais, out Dictionary<string, string> CentreFrais_Dictionnary)
        {



            CentreFrais_Dictionnary = new Dictionary<string, string>();
            IEntityList machine_liste = null;
            //recuperation de la machine clipper et initialisation des listes
            //CentreFrais_Dictionnary = null;
            JohnDeere_Machine = null;
            JohnDeere_Centre_Frais = null;
            //CentreFrais_Dictionnary.Clear();
            //verification que toutes les machineS sont conformes pour une intégration clipper
            ///remplissage des machines et verification de la presence du centre de frais demandé par clipper
            machine_liste = contextlocal.EntityManager.GetEntityList("_CUT_MACHINE_TYPE");
            machine_liste.Fill(false);


            foreach (IEntity machine in machine_liste)

            {
                IEntity cf;
                cf = machine.GetFieldValueAsEntity("CENTREFRAIS_MACHINE");

                if (!object.Equals(machine.GetFieldValueAsEntity("CENTREFRAIS_MACHINE"), null))
                {
                    cf = machine.GetFieldValueAsEntity("CENTREFRAIS_MACHINE");
                }
                else
                {
                    cf = null;
                    Alma_Log.Write_Log(MethodBase.GetCurrentMethod().Name + ": Missing  cost center on : " + machine.DefaultValue);
                    Alma_Log.Error("centre de frais non defini sur la machine  !!!" + machine.DefaultValue, MethodBase.GetCurrentMethod().Name);

                }

                ///creation du dictionnaire des machines installées   
                if (cf.DefaultValue != "" && machine.DefaultValue != "" && JohnDeere_Param.get_JohnDeere_Machine_Cf() != null
                    )
                {
                    if (CentreFrais_Dictionnary.ContainsKey(cf.DefaultValue) == false) { CentreFrais_Dictionnary.Add(cf.DefaultValue, machine.DefaultValue); }

                    if (cf.DefaultValue == JohnDeere_Param.get_JohnDeere_Machine_Cf())
                    {
                        if (JohnDeere_Param.get_JohnDeere_Machine_Cf() != "Undef SAGE machine")
                        {
                            JohnDeere_Centre_Frais = cf;
                            JohnDeere_Machine = machine;
                        }
                        else
                        {
                            Alma_Log.Write_Log(MethodBase.GetCurrentMethod().Name + ": Missing  SAGE machine !!! ");
                            Alma_Log.Error("IL MANQUE LA MACHINE SAGE !!!", MethodBase.GetCurrentMethod().Name);
                            return false;//throw new Exception(machine.DefaultValue + " : Missing  cost center definition"); 
                        }

                    }

                }

                else
                { /*on log on arrete tout */
                    Alma_Log.Write_Log(MethodBase.GetCurrentMethod().Name + ": Missing  cost center definition on a machine !!! ");
                    Alma_Log.Error("IL MANQUE LE CENTRE DE FRAIS SUR L UNE DES MACHINES INSTALLEE !!!", MethodBase.GetCurrentMethod().Name);
                    return false;//throw new Exception(machine.DefaultValue + " : Missing  cost center definition"); 
                }
            }
            return true;


        }

        /// <summary>
        /// creation d'une references vide pour creation
        /// </summary>
        /// <param name="contextlocal">context</param>
        /// <param name="line_dictionnary">dictionnaire de ligne</param>
        public IEntity CreateNewReference(IContext contextlocal, Dictionary<string, object> line_dictionnary, ref string NewReferenceName)
        {
            try
            {



                IEntity newreference = null;
                IEntity material = null;
                //IEntity machine = null;

                Alma_Log.Write_Log(MethodBase.GetCurrentMethod().Name + ": creation d'une nouvelle piece !! ");
                // string referenceName = null;
                int index_extension = 0;
                //si la machine clipper n'est pas nulle
                //on initialise la machine a la machine clipper
                //if (JohnDeere_machine.Id32 != 0) { machine = JohnDeere_machine; }

                if (line_dictionnary.ContainsKey("_MATERIAL") && line_dictionnary.ContainsKey("THICKNESS") && line_dictionnary.ContainsKey("_NAME"))
                {
                    //recuperation de la matiere 
                    material = GetMaterialEntity(contextlocal, ref line_dictionnary);
                    //recupe du nom de la geométrie                 
                    //string referenceName = "undef";    just in case mais inutiel
                    if (NewReferenceName == null)
                    {

                        Alma_Log.Write_Log(MethodBase.GetCurrentMethod().Name + ": Unfortunate error: NewreferenceName does not existes and new reference has been created !! ");
                        if (Data_Model.ExistsInDictionnary("FILENAME", ref line_dictionnary))
                        {

                            NewReferenceName = line_dictionnary["FILENAME"].ToString();
                            if (NewReferenceName.ToUpper().IndexOf(".DPR.EMF") > 0) { index_extension = 7; }
                            if (NewReferenceName.ToUpper().IndexOf(".DPR") > 0) { index_extension = 4; }
                        }
                        else
                        {
                            NewReferenceName = line_dictionnary["_NAME"].ToString();
                        }

                        NewReferenceName = Path.GetFileNameWithoutExtension(@NewReferenceName.Substring(0, (NewReferenceName.Length) - index_extension));

                    }




                    //

                    //creation des infos complementaires de reference notamment les données sans dt
                    //creation de l'entité
                    newreference = contextlocal.EntityManager.CreateEntity("_REFERENCE");
                    //remplacement par la machine clipper dont le cf est clip7
                    //avant tou on let la machine clipper par defaut
                    //champs standards

                    newreference.SetFieldValue("_DEFAULT_CUT_MACHINE_TYPE", SimplifiedMethods.GetDefaultAvailableMachine(contextlocal, material));

                    newreference.SetFieldValue("_NAME", NewReferenceName);
                    newreference.SetFieldValue("_MATERIAL", material.Id32);

                    if (contextlocal.UserId != -1)
                    {
                        newreference.SetFieldValue("_AUTHOR", contextlocal.UserId);
                    }
                    //newreference.SetFieldValue("_AUTHOR", contextlocal.UserId); 


                    newreference.Save();
                    //creation de la prepâration associée
                    //AF_ImportTools.SimplifiedMethods.CreateMachinablePartFromReference(contextlocal, newreference, machine);
                    SimplifiedMethods.CreateMachinableParts(contextlocal, newreference, material);

                }



                return newreference;

            }
            catch (Exception ie)
            {
                Alma_Log.Write_Log(MethodBase.GetCurrentMethod().Name + " : Fails");
                System.Windows.Forms.MessageBox.Show(ie.Message);
                return null;
            }






        }
        /// <summary>
        /// controle de l'integerite des données du fichier texte
        /// on controle les champs obligatoires pour l'import, et l'existantce du centre de frais avant de continue l'import
        /// </summary>
        /// <param name="line_dictionnary">dictionnaire de ligne interprété par le datamodel</param>
        /// <returns>false ou tuue si integre</returns>
        public Boolean  CheckDataIntegerity(IContext contextlocal, Dictionary<string, object> line_dictionnary)
        {
            //
            try
            {
                ///////////////////////////////////////////////////////////////////////////
                ///condition cumulées sur result?                
                Boolean result = true;
                ///////////////////////////////////////////////////////////////////////////
                string currenfieldsname;
                ///matiere
                ///

                //la geomerie est elle spécifiée ?
                currenfieldsname = "_REFERENCE";
                if (line_dictionnary.ContainsKey(currenfieldsname))
                {
                    result = result & true;
                }
                else
                {
                    Alma_Log.Write_Log(MethodBase.GetCurrentMethod().Name + currenfieldsname + " : missing ");
                    Alma_Log.Write_Log_Important(MethodBase.GetCurrentMethod().Name + ":" + line_dictionnary["_REFERENCE"] + ": champs obligatoire Affaire non detectée sur la ligne a importée, line ignored"); result = result & false;
                    result = result & false;
                }


                ///////////////////////////////////////////////////////////////////////////
                //les quantités negatives sont interdites
                currenfieldsname = "_QUANTITY";
                if (line_dictionnary.ContainsKey(currenfieldsname))
                {
                    if (int.Parse(line_dictionnary["_QUANTITY"].ToString().Trim()) < 0 || int.Parse(line_dictionnary["_QUANTITY"].ToString().Trim()) == 0)
                    {
                        Alma_Log.Error(line_dictionnary["_NAME"] + ":_QUANTITY negative ou null detecté sur la ligne a importée, line ignored", MethodBase.GetCurrentMethod().Name);
                        Alma_Log.Write_Log_Important(MethodBase.GetCurrentMethod().Name + ":" + line_dictionnary["_NAME"] + ": champs obligatoire :_QUANTITY non detecté sur la ligne a importée, line ignored"); result = false;
                    }
                }
                else
                {
                    Alma_Log.Write_Log(MethodBase.GetCurrentMethod().Name + currenfieldsname + " : missing ");
                    result = result & false;
                }

                ///////////////////////////////////////////////////////////////////////////
                //le nom de la piece à produire doit exister
                currenfieldsname = "_NAME";
                if (line_dictionnary.ContainsKey(currenfieldsname) != true)
                {
                    Alma_Log.Write_Log(MethodBase.GetCurrentMethod().Name + currenfieldsname + " : missing ");
                    Alma_Log.Write_Log(MethodBase.GetCurrentMethod().Name + ": champs obligatoire:  pas de nom de reference trouvée"); result = result & false;
                    Alma_Log.Write_Log_Important(MethodBase.GetCurrentMethod().Name + ":" + line_dictionnary["_NAME"] + ": champs obligatoire: pas de non de piece detecté sur la ligne a importée, line ignored"); result = result & false;
                }
                else
                {
                    result = result & true;
                }


                ///////////////////////////////////////////////////////////////////////////
                //les matieres sont désormais obligatoires
                //string nuance_name = line_dictionnary["_MATERIAL"].ToString().Replace('§', '*');
                //string nuance = null;
                //string material_name = null;
                //double thickness = 0;




                return result;
            }


            catch (Exception ie)
            {
                Alma_Log.Write_Log(MethodBase.GetCurrentMethod().Name + ": erreur "+ie.Message);
                // MessageBox.Show(ie.Message);
                return false;
            }



        }
        /// <summary>
        /// renvoie l'entite matiere a partie de la nuance et de l'epaisseur contenu dans le line dictionnary
        /// </summary>
        /// <param name="contextlocal">ientity context</param>
        /// <param name="material_name">ientity  material</param>
        /// <param name="line_dictionnary">dictionnary <string,object> line_dictionnary</param>
        /// <returns></returns>
        public IEntity GetMaterialEntity(IContext contextlocal, ref Dictionary<string, object> line_dictionnary)
        {
            IEntity material = null;

            try
            {

                //IEntityList materials = null;
                //verification simple par nom nuance*etat epaisseur en rgardnat une structure comme ceci
                //"SPC*BRUT 1.00" //attention pas de control de l'obsolecence pour le moment
                if (line_dictionnary.ContainsKey("_MATERIAL") && line_dictionnary.ContainsKey("THICKNESS"))
                {
                    material = Material.getMaterial_Entity(contextlocal, line_dictionnary["_MATERIAL"].ToString(), Convert.ToDouble(line_dictionnary["THICKNESS"]));
                }/*
                materials = contextlocal.EntityManager.GetEntityList("_MATERIAL", "_NAME", ConditionOperator.Equal, material_name);
                materials.Fill(false);

                if (materials.Count() > 0 && materials.FirstOrDefault().Status.ToString() == "Normal")
                { material = materials.FirstOrDefault(); }
                else { material = null; }*/


                return material;
            }
            catch (Exception ie)
            {
                Alma_Log.Write_Log(MethodBase.GetCurrentMethod().Name + ": erreur :");
                MessageBox.Show(ie.Message);
                return material;
            }

        }

        /// <summary>
        /// recupere une reference en fonction d'un 
        /// numero d'indice //(remplacé par un indice d'identification piece en position 27)
        /// ce numero d'indice est egale a l'id de la piece dans la table reference sauf si l'indice est negatif.
        /// Si l'indice est negatif alors l'indice vient d'une piece cotée.
        /// 
        /// </summary>
        /// <param name="contextlocal">contexte local</param>
        /// <param name="reference">entite reference</param>
        /// <param name="line_dictionnary">dictionnaire de ligne</param>
        /// <returns>true si la reference est detectee en fonction du numero de plan</returns>
        public bool GetReference(IContext contextlocal, ref IEntity reference, ref Dictionary<string, object> line_dictionnary, ref string NewReferenceName)
        {
            reference = null;
            //IEntityList references = null;
            //Int32 new_reference_id = 0;
            IEntityList reference_list = null;
            IEntity reference_part = null;
            //IEntity material = null;
            bool result = false;


            try
            {
                //int index_extension = 7;
                if (Data_Model.ExistsInDictionnary("_REFERENCE", ref line_dictionnary))
                {

                    //IEntityList reference_partlist;
                    string reference_name = null;
                    reference_name = line_dictionnary["_REFERENCE"].ToString();
                    bool LockedReference = false; //reference en cours de dessin si true

                    if (reference_name != "")
                    {
                        Alma_Log.Write_Log(MethodBase.GetCurrentMethod().Name + ":Pièce" + reference_name + "  identifiée. ");
                        //depuis le sp3 on recherche dans quote part
                        reference_list = contextlocal.EntityManager.GetEntityList("_REFERENCE", "_NAME", ConditionOperator.Equal, reference_name);
                        //reference_list = contextlocal.EntityManager.GetEntityList("_REFERENCE",LogicOperator.And, "_NAME", ConditionOperator.Equal, reference_name, "AF_LOCKED_REFERENCE", ConditionOperator.NotEqual, true);
                        reference_list.Fill(false);


                        if (reference_list.Count == 1)
                        {
                            reference_part = SimplifiedMethods.GetFirtOfList(reference_list);
                            LockedReference = reference_part.GetFieldValueAsBoolean("AF_LOCKED_REFERENCE");

                            if (reference_part != null && LockedReference != true)
                            {
                                NewReferenceName = reference_part.GetFieldValueAsString("_NAME");
                                reference = reference_part;
                                result = true;
                            }
                            else
                            {
                                Alma_Log.Error("REFERENCE EN REVISION : ref " + reference_name, MethodBase.GetCurrentMethod().Name + "Reference non recuperable : import impossible");
                                Alma_Log.Write_Log(MethodBase.GetCurrentMethod().Name + ":" + reference_name + ": not found :");
                                NewReferenceName = null;
                                result = false;
                            }



                        }
                        else if (reference_list.Count > 1)
                        {
                            // sinon erreur et on creer une nouvelle piece
                            Alma_Log.Error("Doublon detecté : ref " + reference_name, MethodBase.GetCurrentMethod().Name + " import abandonné");
                            Alma_Log.Write_Log(MethodBase.GetCurrentMethod().Name + ":" + reference_name + ": called twice :");
                            result = false;
                        }

                        else
                        {
                            // sinon erreur et on creer une nouvelle piece
                            Alma_Log.Error("Part NON TROUVEE  : ref " + reference_name, MethodBase.GetCurrentMethod().Name + " import abandonné");
                            Alma_Log.Write_Log(MethodBase.GetCurrentMethod().Name + ":" + reference_name + ": not found :");
                            result = false;
                        }

                    }


                }

                else
                {

                    //reference non indiqué on cree  une nouvelle piece 
                    Alma_Log.Error("AUCUNE REFERENCE TRANSMISE : ", MethodBase.GetCurrentMethod().Name + "Reference non trouvée : crreation d'une nouvelle piece");
                    Alma_Log.Write_Log(MethodBase.GetCurrentMethod().Name + ":Part id_cfao   : not found :");
                    NewReferenceName = line_dictionnary["_REFERENCE"].ToString(); ;
                    result = false;
                    //result = true;


                }

                return result;


            }

            catch (Exception ie)
            {
                //on log
                Alma_Log.Write_Log(MethodBase.GetCurrentMethod().Name + ": erreur :");
                MessageBox.Show(ie.Message);
                return result;
            }



        }
        /// met a jour les valeurs   dans les pieces a produires almacam
        /// </summary>
        /// <param name="contextlocal">contexte context</param>
        /// <param name="sheet">ientity sheet  </param>
        /// <param name="stock">inentity stock</param>
        /// <param name="line_dictionnary">dictionnary linedisctionary</param>
        /// <param name="type_tole">type tole  ou chute</param>
        /// 
        public void Update_Part_Item(IContext contextlocal, ref IEntity reference_to_produce, ref Dictionary<string, object> line_dictionnary)
        {
            try
            {
                foreach (var field in line_dictionnary)
                {
                    //on recupere la reference a usiner
                    //rien pour le moment
                    //on verifie que le champs existe bien avant de l'ecrire
                    if (contextlocal.Kernel.GetEntityType("_TO_PRODUCE_REFERENCE").FieldList.ContainsKey(field.Key))
                    {
                        //traitement specifique
                        switch (field.Key)
                        {
                            case "_REFERENCE":
                                //on ignore car a reference est settée avant
                                break;

                            case "_NAME":
                                //on ignore car a reference est settée avant
                                reference_to_produce.SetFieldValue("AF_ORDRE", field.Value);
                                reference_to_produce.SetFieldValue("_NAME", field.Value);
                                break;

                            default:
                                reference_to_produce.SetFieldValue(field.Key, field.Value);
                                break;
                        }
                    }


                }

            }
            catch (Exception ie)
            {
                Alma_Log.Write_Log(MethodBase.GetCurrentMethod().Name + ": erreur :");
                MessageBox.Show(ie.Message);
            }
        }
        /// <summary>
        /// en standard
        /// import un of 
        /// </summary>
        /// <param name="contextlocal">contexte alma cam</param>
        /// <param name="pathToFile">chemin vers le fichier csv separateur ";"</param>
        /// <param name="sans_donnees_technique">true si import sans données techniques</param>
        /// <param name="DataModelString">string de description d'une ligne csv sous la forme 
        /// numeroIndex#NomChampAlmaCam#Type  exemple : 0#AFFAIRE#STRING</param>
        /// <summary>
        public void Import(IContext contextlocal, string pathToFile, string DataModelString)
        {

            //recuperation des path
            CsvImportPath = pathToFile;


            try

            {
                //verification standards
                //creation du timetag d'import
                string timetag = string.Format("{0:d_M_yyyy_HH_mm_ss}", DateTime.Now);
                Alma_Log.Create_Log(JohnDeere_Param.getVerbose_Log());
                //bool import_sans_donnee_technique = false;
                Alma_Log.Write_Log(MethodBase.GetCurrentMethod().Name + ": importe tag :" + timetag);
                //ouverture du fichier csv lancement du curseur
                // Set cursor as hourglass
                Cursor.Current = Cursors.WaitCursor;



                if (File.Exists(CsvImportPath) == false)
                {
                    Alma_Log.Error("Fichier Non Trouvé:" + CsvImportPath, MethodBase.GetCurrentMethod().Name);
                    throw new Exception("csv File Note Found:\r\n" + CsvImportPath);
                }
                //avec ou sans dt
                // if (Sans_Donnees_Technique) { Alma_Log.Write_Log(MethodBase.GetCurrentMethod().Name + ": import  sans dt !! " + CsvImportPath); } else { Alma_Log.Write_Log(MethodBase.GetCurrentMethod().Name + ": import standard !! " + CsvImportPath); }


                using (StreamReader csvfile = new StreamReader(CsvImportPath, Encoding.Default))
                {
                    //recuperation des elements de la base almacam
                    //declaration des dictionaires
                    Dictionary<string, object> line_Dictionnary = new Dictionary<string, object>();
                    //on lit les centres de frais 
                    //Dictionary<string, string> CentreFrais_Dictionnary = null;
                    Data_Model.setFieldDictionnary(DataModelString);
                    Alma_Log.Write_Log(MethodBase.GetCurrentMethod().Name + ": reading data model :success !! ");
                    ///remplissage des machines et verification de la presence du centre de frais demandé par clipper
                    ///plus utile
                    //IEntityList machine_liste = contextlocal.EntityManager.GetEntityList("_CUT_MACHINE_TYPE");
                    //machine_liste.Fill(false);

                    //recuperation de la machine clipper
                    //IEntity JohnDeere_Machine = null;
                    //IEntity JohnDeere_Centre_Frais = null;
                    //recuperation de la machine clipper et construction de la liste machine
                    //Get_JohnDeere_Machine(contextlocal, out JohnDeere_Machine, out JohnDeere_Centre_Frais, out CentreFrais_Dictionnary);

                    //verification que toutes les machines sont conformes pour une intégration clipper


                    int ligneNumber = 0;
                    //lecture à la ligne
                    string line;
                    line = null;

                    while (((line = csvfile.ReadLine()) != null))
                    {

                        //on ne traite pas les lignes vides
                        ligneNumber++;

                        //verification des headers
                        if (ligneNumber == 1) { continue; }
                        //verificziton des lignes vides
                        if ((line.Trim()) == "")
                        {
                            Alma_Log.Write_Log(MethodBase.GetCurrentMethod().Name + " : " + ligneNumber + ": empty line detected !! ");
                            continue;
                        }

                        //lecture des donnees
                        //ignore first line


                        line_Dictionnary = Data_Model.ReadCsvLine_With_Dictionnary(line);
                        Alma_Log.Write_Log(MethodBase.GetCurrentMethod().Name + " : " + ligneNumber + ": line disctionnary interpreter succeeded !! ");

                        //control des données    //verification des donnée importées
                        if (CheckDataIntegerity(contextlocal, line_Dictionnary) != true)
                        {
                            /*on log et on continue(on passe a la ligne suivante*/
                            Alma_Log.Write_Log(MethodBase.GetCurrentMethod().Name + " : " + ligneNumber + ": data integrity fails, line ignored !!! ");
                            continue;
                        }





                        IEntity reference_to_produce = null;
                        IEntity reference = null;
                        string NewReferenceName = null;
                        //
                        //on recherche la refeence avec la bonne matiere /epaisseur si elles n'existe pas on la creer 
                        if (GetReference(contextlocal, ref reference, ref line_Dictionnary, ref NewReferenceName) == false)
                        {
                            //on ne creer jamais de pieces
                            continue;


                        }

                        //////on met a jour les données sur les piece 2d  : CUSTOMER_REFERENCE_INFOS
                        /*
                        if (reference != null)
                        {///champs spécifique piece 2d
                            //infos liées a l'import cfao
                            //CUSTOMER_REFERENCE_INFOS;
                            JohnDeere_Param.GetlistParam(contextlocal);
                            string Field_value = JohnDeere_Param.GetPath("CUSTOMER_REFERENCE_INFOS");
                            Field_value.Split('|');//"CUSTOMER"
                            reference.SetFieldValue(Field_value.Split('|')[0], line_Dictionnary[Field_value.Split('|')[1]]);
                            reference.Save();
                            
                        }*/




                        //creation de la nouvelle piece a produire associée
                        // GetMaterialEntity(contextlocal, ref Dictionary<string, object> line_dictionnary);
                        //if (Sans_Donnees_Technique == false) {
                        CreateNewPartToProduce(contextlocal, line_Dictionnary, null, ref reference_to_produce, ref reference, timetag);
                        //on ne creer jamais de references elles doivent etre créées avant
                        //AF_ImportTools.SimplifiedMethods.CreateMachinableParts(contextlocal, reference,  GetMaterialEntity(contextlocal, ref line_Dictionnary));
                        //}
                    }



                }

                // Set cursor as default arrow
                Cursor.Current = Cursors.Default;
                //File_Tools.Rename_Csv(CsvImportPath, timetag);
                File_Tools.Rename_And_Move_Csv(CsvImportPath, "Imported");
                Alma_Log.Final_Open_Log();
                //File_tools
            }

            catch (Exception e)
            {
                Alma_Log.Write_Log(e.Message);
                Alma_Log.Final_Open_Log();

            }
        }

    }




    /// <summary>
    /// retour gp la cloture
    /// </summary>
    ///
    public class JohnDeere_DoOnAction_After_Cutting_end : AfterEndCuttingEvent
    {
        //JohnDeere_DoOnAction_After_Cutting_end
        public override void OnAfterEndCuttingEvent(IContext context, AfterEndCuttingArgs args)
        {
            {
                //this.execute(contextlocal, args.NestingEntity);
                execute(args.ToCutSheetEntity);
            }
        }

        /// <summary>
        /// creation auto du fichier texte à  la cloture
        /// </summary>
        /// <param name="args"></param>
        public void execute(IEntity entity)
        {
            //recuperation des path
            JohnDeere_Param.GetlistParam(entity.Context);
            string export_gpao_path = JohnDeere_Param.GetPath("Export_GPAO");

           
            {
                JohnDeere_Infos current_JohnDeere_nestinfos = new JohnDeere_Infos();
                current_JohnDeere_nestinfos.GetNestInfosBySheet(entity);
                current_JohnDeere_nestinfos.Export_NestInfosToFile(export_gpao_path);
            
                current_JohnDeere_nestinfos = null;
            }


        }
        /// <summary>
        /// retourne le fichier clipper a partie de la boite de dialogue de la ploateforme
        /// 
        /// </summary>
        /// <param name="SelectedEntities"> IEntitySelector Entityselector </param>
        /// 
        public void execute(IEnumerable<IEntity> SelectedEntities)
        {
            string stage = SelectedEntities.First().EntityType.Key;
            //creation du fichier de sortie
            JohnDeere_Param.GetlistParam(SelectedEntities.First().Context);
            if (SelectedEntities.Count() > 0)
            {
                foreach (IEntity entity in SelectedEntities)
                {
                    string export_gpao_path = JohnDeere_Param.GetPath("Export_GPAO");

                    //using (JohnDeere_Infos current_JohnDeere_nestinfos = new JohnDeere_Infos())
                    {
                        JohnDeere_Infos current_JohnDeere_nestinfos = new JohnDeere_Infos();
                        current_JohnDeere_nestinfos.GetNestInfosBySheet(entity);
                        current_JohnDeere_nestinfos.Export_NestInfosToFile(export_gpao_path);
                        //validation du stock

                        //
                        current_JohnDeere_nestinfos = null;
                    }


                }

            }



        }
    }


    public class Util
    {
        public enum MessageType
        {
            Information,
            Avertissement,
            Erreur
        }

        public enum ModuleType
        {
            ExportCsv,
            ExportTxt,
            ImportOF,
            ImportPièces,
            ImportStock,
            Parametres
        }

        public void WriteTraceLine(string message, MessageType type, Util.ModuleType module)
        {
            Trace.WriteLine(
                    string.Format("{0} {1}\t{2}\t{3}",
                                  DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss"),
                                  module,
                                  type,
                                  message));
        }

        public TextWriterTraceListener OpenTrace(string fileName)
        {
            TextWriterTraceListener textWriterTraceListener = new TextWriterTraceListener(fileName);
            Trace.Listeners.RemoveAt(0);
            Trace.Listeners.Add(textWriterTraceListener);
            Trace.AutoFlush = true;

            return textWriterTraceListener;
        }

        public void CloseTrace(TextWriterTraceListener traceFile)
        {
            traceFile.Close();
            traceFile = null;
        }

        public void ViewLog(string fileName)
        {
            Process.Start("notepad.exe", fileName);
        }

        public void EmfToJpg(string emfFileName)
        {
            string jpgFileName = emfFileName.ToLower().Replace(".emf", ".jpg");

            Image image = Image.FromFile(emfFileName);
            Bitmap bitmap = new Bitmap(image.Width, image.Height);

            Graphics graphics = Graphics.FromImage(bitmap);
            graphics.Clear(Color.White);
            graphics.DrawImage(image, 0, 0, image.Width, image.Height);

            Bitmap tempImage = new Bitmap(bitmap);
            bitmap.Dispose();
            image.Dispose();

            if (File.Exists(jpgFileName))
            {
                File.Delete(jpgFileName);
            }

            tempImage.Save(jpgFileName, ImageFormat.Jpeg);
            tempImage.Dispose();
            File.Delete(emfFileName);
        }

        public double ToDouble(string strIn)
        {
            string decimalSeparator = Thread.CurrentThread.CurrentCulture.NumberFormat.NumberDecimalSeparator;
            Regex pattern = new Regex("[.,]");

            return double.Parse(pattern.Replace(strIn, decimalSeparator));
        }

        public string CheckFolder(string folder)
        {
            if (folder == string.Empty)
            {
                folder = "-1;Répertoire non défini";
            }

            if (Directory.Exists(folder) == false)
            {
                folder = "-2;Répertoire inexistant " + folder;
            }

            if (folder[folder.Length - 1] != '\\')
            {
                folder += "\\";
            }

            if (!folder.StartsWith("-"))
            {
                try
                {
                    DirectorySecurity ds = Directory.GetAccessControl(folder);
                }
                catch
                {
                    folder = "-3;Répertoire en lecture seule " + folder;
                }
            }
            return folder;
        }

        public List<Tuple<string, string>> Folders(IContext context)
        {
            Util util = new Util();

            string tempPath = util.CheckFolder(Path.GetTempPath()) + "log.txt";

            TextWriterTraceListener logTextWriterTraceListener = util.OpenTrace(tempPath);
            bool erreur = false;

            IEntityList folders = null;
            try
            {
                folders = context.EntityManager.GetEntityList("AF_FOLDERS");
            }
            catch
            {
                util.WriteTraceLine("Key AF_FOLDERS missing", MessageType.Erreur, ModuleType.Parametres);
                erreur = true;
            }

            List<Tuple<string, string>> retour = new List<Tuple<string, string>>();

            if (folders != null)
            {
                folders.Fill(false);
                if (folders.Count != 0)
                {
                    foreach (IEntity folder in folders)
                    {
                        string tmp = string.Empty;
                        try
                        {
                            tmp = (string)folder.GetFieldValue("AF_FOLDER_NAME");

                            string tmpFolder = (string)folder.GetFieldValue("AF_FOLDER_LOCATION");
                            if (!tmpFolder.ToLower().Contains(".csv"))
                            {
                                tmpFolder = util.CheckFolder(tmpFolder);
                            }

                            if (tmpFolder.StartsWith("-"))
                            {
                                util.WriteTraceLine(tmpFolder.Split(';')[1], MessageType.Erreur, ModuleType.Parametres);
                                erreur = true;
                            }
                            else
                            {
                                retour.Add(new Tuple<string, string>(tmp, tmpFolder));
                            }
                        }
                        catch
                        {
                        }
                    }
                }
            }
            if (erreur)
            {
                util.ViewLog(tempPath);
                util.CloseTrace(logTextWriterTraceListener);
                util = null;
            }

            return retour;
        }
    }

    /// <summary>
    /// retour gp à L ENVOIE A L ATELIER
    /// </summary>
    /// 

    public class JohnDeere_DoOnAction_AfterSendToWorkshop : AfterSendToWorkshopEvent
    {
        
        //cette fonction est lancée autant de fois qu'il y a de selection
        //la multiselection n'est pas controlée
        public override void OnAfterSendToWorkshopEvent(IContext contextlocal, AfterSendToWorkshopArgs args)
        {
            try
            {
                string stage = "_TO_CUT_NESTING";
                IEntityList nestinglist = contextlocal.EntityManager.GetEntityList(stage, "_REFERENCE", ConditionOperator.Equal, args.NestingEntity.GetFieldValueAsString("_REFERENCE"));
                nestinglist.Fill(false);
                //ecriture du fichier de sortie

                planification(contextlocal, nestinglist);
                this.execute(nestinglist);
            }
            catch (Exception ie)
            {
                string msg = ie.Message;
                Alma_Log.Write_Log(" Erreur sur la cloture.. ");
                Alma_Log.Close_Log();
            }
        }

        private void planification(IContext contextlocal, IEntityList nestingsList)
        {
            Util util = new Util();

            List<Tuple<string, string>> repertoires = util.Folders(contextlocal);

            var index = repertoires.FindIndex(s => s.Item1 == "Export_Plannification");
            string repertoireExport = string.Empty;
            string fichierExport = string.Empty;

            if (index != -1)
            {
                repertoireExport = repertoires[index].Item2;
            }
            else
            {
                MessageBox.Show("Folders : Clé Export_Plannification introuvable");
                return;
            }

            List<string> csvToExport = new List<string>
            {
                "NomProgramme; RefTole; NomPlanDeCoupe; NOPGM; MULTIPLICITY; TOTALTIME; NomPiece; QuantitePiecePourLePlanDeCoupe"
            };

            foreach (IEntity nesting in nestingsList)
            {
                IEntityList nestedReferenceList = contextlocal.EntityManager.GetEntityList("_NESTED_REFERENCE", "_NESTING", ConditionOperator.Equal, nesting.Id);
                nestedReferenceList.Fill(false);

                IEntityList cnFileEntity = contextlocal.EntityManager.GetEntityList("_CN_FILE", "_SEQUENCED_NESTING", ConditionOperator.Equal, nesting.Id32);
                cnFileEntity.Fill(false);
                IEntity firstCnFile = cnFileEntity.First();

                string numeroPlacement = "-" + nesting.GetFieldValueAsString("_NAME").Split('-')[2];
                string nomProgramme = nesting.GetFieldValueAsString("_REFERENCE");
                string refTole = nesting.GetFieldValueAsEntity("_MATERIAL").GetFieldValueAsEntity("_QUALITY").GetFieldValueAsString("_NAME");   //  matière
                string nomPlanDeCoupe = nesting.GetFieldValueAsString("_REFERENCE") + numeroPlacement;
                string noPgm = firstCnFile.GetFieldValueAsInt("_NOPGM").ToString();
                string totalTime = (nesting.GetFieldValueAsDouble("_TOTALTIME") / 60).ToString("#0.00");

                string multiplicite = nesting.GetFieldValueAsInt("_QUANTITY").ToString();
                foreach (IEntity nestedReference in nestedReferenceList)
                {
                    IEntity reference = nestedReference.GetFieldValueAsEntity("_REFERENCE");
                    string partName = reference.GetFieldValueAsString("_NAME");
                    string quantitePiecePourLePlanDeCoupe = (nestedReference.GetFieldValueAsInt("_QUANTITY") / nesting.GetFieldValueAsInt("_QUANTITY")).ToString();
                    csvToExport.Add(nomProgramme + ";" + refTole + ";" + nomPlanDeCoupe + ";" + noPgm + ";" + multiplicite + ";" + totalTime + ";" + partName + ";" + quantitePiecePourLePlanDeCoupe);
                    fichierExport = nomProgramme + ".csv";
                }
            }

            if (File.Exists(fichierExport))
            {
                using (StreamWriter sw = new StreamWriter(repertoireExport + "\\" + fichierExport, true))
                {
                    bool firstLine = true;
                    foreach (string line in csvToExport)
                    {
                        if (firstLine == false)
                        {
                            sw.WriteLine(line);
                        }
                        else
                        {
                            firstLine = false;
                        }
                    }

                    sw.WriteLine("xxxxxxxxxxxxxxxxxxxxxx");

                }
            }
            else
            {
                using (StreamWriter sw = new StreamWriter(repertoireExport + "\\" + fichierExport))
                {
                    foreach (string line in csvToExport)
                    {
                        sw.WriteLine(line);
                    }
                }
            }
        }

        private List<string>  GetRetourGp(IEntityList nestinglist, string material, string filename)
        {
            try
            {
                List<string> RetourGp = new List<string>();
                //contexte du premier nesting
                if (nestinglist != null)
                {
                    //creation des listes
                    List<JohnDeere_Infos> JohnDeere_Infoslist = new List<JohnDeere_Infos>();
                    List<Nested_PartInfo> parts = new List<Nested_PartInfo>();
                    List<Tole> toles = new List<Tole>();
                    double partstotalsurface = 0;
                    double sheetTotalSurface = 0;
                    //sommes des surfaces des toles
                    //calcul du ratio
                   
                   
                    string separator = "\t";
                    //double checksom = 0;
                    string format_double = "{0:0.000}";
                    int accuracy = JohnDeere_Param.get_string_format_double().Split('.')[1].Count() - 1;
                    List<string> ofdone = new List<string>();
                    Int64 toleNumber = 0;
                        //// recuperation des informations a regrouper
                        foreach (IEntity nesting in nestinglist)
                        {

                                        toleNumber += nesting.GetFieldValueAsInt("_QUANTITY");
                                        JohnDeere_Infos JohnDeere_Infos = new JohnDeere_Infos();
                                        //meme matiere uniquement
                                        if (nesting.GetFieldValueAsEntity("_MATERIAL").GetFieldValueAsString("_NAME") == material)
                                        {
                                            JohnDeere_Infos = this.execute(nesting);
                                            
                                            foreach (Nest_Infos_2 currentnestinfos in JohnDeere_Infos.nestinfoslist)
                                            {
                                                parts = parts.Concat(currentnestinfos.Nested_Part_Infos_List).ToList();
                                                partstotalsurface += currentnestinfos.Calculus_Parts_Total_Surface;
                                                sheetTotalSurface += currentnestinfos.Tole_Nesting.Sheet_Surface;
                                              
                                            }

                         
                            
                                        }
                                        else
                                        {

                                            Alma_Log.Write_Log(" Des matieres differentes ont ete detectées dans les  regroupements par reference ");
                                            Alma_Log.Close_Log();
                                            return null;

                                        }
                        }

                        //on recupere le ratio global
                        double ratio = sheetTotalSurface / partstotalsurface;
                    
                    //on boucle sur les of
                    foreach (Nested_PartInfo p in parts)
                        {
                            if (!ofdone.Contains(p.Part_Name))
                            {
                                List<Nested_PartInfo> of = new List<Nested_PartInfo>();
                                of = parts.Where(x => x.Part_Name == p.Part_Name).ToList();
                                //ecriture des lignes//
                                double partquantity;
                                
                                double totalpartsurface;
                            ///
                                double s =sheetTotalSurface;
                                partquantity = of.Sum(c => c.Nested_Quantity);
                                //totalpartsurface = of.Sum(c => c.Surface * ratio) * partquantity;
                                totalpartsurface = p.Surface * partquantity* ratio;
                                //ajout de la ligne dans le fichier de retour
                                  string line = p.Part_Name + separator + partquantity + separator + String.Format(format_double, totalpartsurface/ sheetTotalSurface* toleNumber) + separator + filename;

                            ////
                            RetourGp.Add(line);
                            
                            ofdone.Add(p.Part_Name);





                        }


                        




                    }


                    JohnDeere_Infoslist = null;
                    toles = null;                   
                    ofdone = null;
                    //lapurge
                        parts = null;

                    return RetourGp;


                }

                else { return null; }
               


            }
            catch (Exception ) { return null; }
        }



        /// <summary>
        /// creation auto du fichier texte à  la cloture
        /// </summary>
        /// <param name="args"></param>
        public JohnDeere_Infos execute(IEntity entity)
        {
            //recuperation des path

            JohnDeere_Infos current_johndeere_nestinfos = new JohnDeere_Infos();
            current_johndeere_nestinfos.GetNestInfosByNesting(entity.Context, entity, "_TO_CUT_NESTING");
            return current_johndeere_nestinfos;


        }

        /// <summary>
        /// creation auto du fichier texte à la cloture
        /// ATTENTION  :cette methode executer sur chaque placmeents sans disctinction
        /// </summary>
        /// <param name="args">liste de placement</param>
        public void execute(IEntityList nestinglist)
        {
            try
            {
                //recuperation des paramtres d'export
                JohnDeere_Param.GetlistParam(nestinglist.First().Context);
                string export_gpao_path = JohnDeere_Param.GetPath("Export_GPAO");
                List<string> retourGp=new List<string>();     

                string filename = nestinglist.First().GetFieldValueAsString("_REFERENCE");
                string material = nestinglist.First().GetFieldValueAsEntity("_MATERIAL").GetFieldValueAsString("_NAME");
                //ecriture du fichier de sortie
                foreach (string s in this.GetRetourGp(nestinglist, material, filename))
                {//renvoie une liste contenant les données gp pour chaque fichier
                    retourGp.Add(s);
                }
                
                using (StreamWriter export_gpao_file = new StreamWriter(@JohnDeere_Param.GetPath("Export_GPAO") + "\\" + filename + ".txt"))
                {
                    foreach (string line in retourGp)
                            {//ecriture du contenu de retourgp
                                               export_gpao_file.WriteLine(line.Replace(".", ","));
                            }

                    export_gpao_file.WriteLine("xxxxxxxxxxxxxxxxxxxxxx");
                }


                retourGp = null;

            }
            catch (Exception )
            {

            }
        }

    }


    /// <summary>
    /// retour gp AVANT L ENVOIE A L ATELIER
    /// </summary>
    /// MyBeforeSendToWorkshopEvent : BeforeSendToWorkshopEvent
    /// 
    /// 
    /// public class JohnDeere_DoOnAction_BeforeSendToWorkshop : BeforeSendToWorkshopEvent


    //creation des clipper infos issue des generic gp infos
    public class JohnDeere_Infos : Generic_GP_Infos
    {
        public override void Export_NestInfosToFile(string export_gpao_path)
        {
            base.Export_NestInfosToFile(export_gpao_path);
            /***/

            ///recuperation des placements selectionnés

            foreach (Nest_Infos_2 currentnestinfos in this.nestinfoslist)
            {
                //string Separator = ";";
                string filename = "";
                //nom du placement si refrence placement vide
                //filename = currentnestinfos.Tole_Nesting.Sheet_Reference; //?? currentnestinfos.Tole_Nesting.To_Cut_Sheet_Name;
                filename = currentnestinfos.Nesting_Reference ?? currentnestinfos.Tole_Nesting.To_Cut_Sheet_Name;

                using (StreamWriter export_gpao_file = new StreamWriter(@export_gpao_path + "\\" + filename + ".txt"))
                {
                    //recuperation du nom de placement



                    double checksum = 0;
                    //ecriture des details pieces
                    //recue
                    foreach (Nested_PartInfo JohnDeerpartSet in currentnestinfos.Nested_Part_Infos_List)//.Nested_PartSet_Infos_List)
                    {
                        string separator = "\t";
                        string ofname = JohnDeerpartSet.Part_Name;
                        //on compte le nombre de caracter apres la virgule de la chaine (JohnDeere_Param.get_string_format_double())-1
                        int accuracy = JohnDeere_Param.get_string_format_double().Split('.')[1].Count() - 1;
                        string format_double = "{0:0.000}"; ////JohnDeere_Param.get_string_format_double();
                        double partNestedQuantity = JohnDeerpartSet.Nested_Quantity;
                        double partRatio = Math.Round(JohnDeerpartSet.Part_Total_Ratio, accuracy);
                        checksum += partRatio;
                        //scurrentnestinfos.Nested_Part_Infos_List
                        //Offcut_infos_List.Sum(o => o.Sheet_Surface);

                        //ajustement
                        if (JohnDeerpartSet == currentnestinfos.Nested_Part_Infos_List.Last<Nested_PartInfo>())
                        {
                            //JohnDeerpartSet.Part_Total_Ratio = JohnDeerpartSet.Part_Total_Ratio + (1 - checksum);
                            partRatio = partRatio + (1 - checksum);

                        }


                        string line = ofname + separator + partNestedQuantity + separator + String.Format(format_double, partRatio) + separator + filename;
                        export_gpao_file.WriteLine(line);




                    }


                 

                }






            }



        }

        #endregion

        //public string Numlot { get; set; }
        //public string NumMatlot { get; set; }


        /// <summary>
        /// inforlmation specifique a recuperer
        /// </summary>
        /// <param name="Tole"></param>
        /// 

        //infos spec des toles
        public override void SetSpecific_Tole_Infos(Tole Tole)
        {
            base.SetSpecific_Tole_Infos(Tole);
            Tole.Specific_Tole_Fields.Add<string>("NUMMATLOT", Tole.StockEntity.GetFieldValueAsString("NUMMATLOT"));
            Tole.Specific_Tole_Fields.Add<string>("NUMLOT", Tole.StockEntity.GetFieldValueAsString("NUMLOT"));

        }


        //inofs specifiques des parts
        public override void SetSpecific_Part_Infos(List<Nested_PartInfo> Nested_Part_Infos_List)
        {
            base.SetSpecific_Part_Infos(Nested_Part_Infos_List);

            foreach (Nested_PartInfo part in Nested_Part_Infos_List)
            {
                part.Nested_PartInfo_specificFields.Add<string>("AF_ORDRE", part.Part_To_Produce_IEntity.GetFieldValueAsString("AF_ORDRE"));
                part.Nested_PartInfo_specificFields.Add<string>("AF_COMPOSANT", part.Part_To_Produce_IEntity.GetFieldValueAsString("AF_COMPOSANT"));
                part.Nested_PartInfo_specificFields.Add<string>("AF_COMPOSANT_TH", part.Part_To_Produce_IEntity.GetFieldValueAsString("AF_COMPOSANT_TH"));
                part.Nested_PartInfo_specificFields.Add<string>("AF_PT", part.Part_To_Produce_IEntity.GetFieldValueAsString("AF_PT"));
                //on recherche les pieces fantomes
                if (part.Part_To_Produce_IEntity.GetFieldValueAsString("AF_ORDRE") == string.Empty)
                {
                    part.Part_IsGpao = false;
                }

                //part.Nested_PartInfo_specificFields.Add<string>("IDLNROUT", part.Part_To_Produce_IEntity.GetFieldValueAsString("IDLNROUT"));

                //clipperpart.Nested_PartInfo_specificFields.Get<string>("IDLBOM", out idlnbom);


            }




        }

        //infos specifique des chutes
        public override void SetSpecific_Offcut_Infos(List<Tole> Offcut_infos_List)
        {
            base.SetSpecific_Offcut_Infos(Offcut_infos_List);


            foreach (Tole offcut in Offcut_infos_List)
            {

                offcut.Specific_Tole_Fields.Add<string>("NUMLOT", offcut.StockEntity.GetFieldValueAsString("NUMLOT"));
                offcut.Specific_Tole_Fields.Add<string>("NUMMATLOT", offcut.StockEntity.GetFieldValueAsString("NUMMATLOT"));


            }




        }


    }



}

