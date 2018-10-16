using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

using System.Runtime.InteropServices;
using System.Diagnostics;



using Wpm.Implement.Manager;
using Wpm.Implement.Processor;
using Wpm.Implement.ComponentEditor;  // ouverture de fenetres de selection almacam
using Wpm.Schema.Kernel;
using Actcut.ActcutModelManager;
using Actcut.NestingManager;
using Actcut.ResourceManager;
using Actcut.ResourceModel;


//dll personnalisées
using Clipper_Dll;
using ImportTools;
using System.IO;
using Microsoft.Win32;

///executable clipper





namespace Clipper_Import
    {
        class Program
        {
        //initialisation des listes
       // [DllImport("user32.dll")]
       // static extern bool ShowWindow(IntPtr hWnd, int nCmdShow);

       // [DllImport("Kernel32")]
       // private static extern IntPtr GetConsoleWindow();

       // const int SW_HIDE=0;
       // const int SW_SHOW=5;

            static void Main(string[] args)
            {

                string DbName;
                IContext _clipper_Context = null;


                Console.WriteLine("hello");
                Console.WriteLine(args[1].ToString());
                Console.WriteLine(args[0].ToString());
                Console.ReadLine();

                //EventLog.WriteEntry("Clipper_import", "arguments " + args[0] + " ; " + args[1], EventLogEntryType.Information, 255);
               // RegistryKey registryKey = Registry.CurrentUser.OpenSubKey(@"HKEY_CURRENT_USER\Software\Alma\Wpm");
               // DbName = (string)registryKey.GetValue("LastModelDatabaseName");
                DbName = "AlmaCAM_Test";
                ModelsRepository clipper_modelsRepository = new ModelsRepository();
                clipper_modelsRepository = null;
               
              
                _clipper_Context = clipper_modelsRepository.GetModelContext(DbName);  //nom de la base;
                int i = _clipper_Context.ModelsRepository.ModelList.Count();
                Clipper_Param.GetlistParam(_clipper_Context);
         

             
             //string DbName = "AlmaCAM_Clipper_5";
             //creation du model repository
            

             //IntPtr hwnd;
            // hwnd=GetConsoleWindow();
             //ShowWindow(hwnd,SW_HIDE);
             //ShowWindow(hwnd,SW_SHOW);
               
             using (EventLog eventLog = new EventLog("Application"))
             {
                 eventLog.Source = "Clipper_import";
                 //EventLog.WriteEntry("Clipper_import", "Found " + (string)registryKey.GetValue("LastModelDatabaseName"), EventLogEntryType.Information, 255);
                 EventLog.WriteEntry("Clipper_import", "Found " + DbName, EventLogEntryType.Information, 255);
                                 if (args.Length != 0)
                                 {
                                     string fulpathname = args[1];
                                     switch (args[0].ToUpper()) {
                                         //fullpath name
                                         case "STOCK":
                                             //import stock
                                             string csvImportPath = Clipper_Param.GetPath("IMPORT_DM");
                                              string dataModelstring = Clipper_Param.GetModelDM();

                            
                                             using (Clipper_Stock Stock = new Clipper_Stock())
                                             {
                                                 Stock.Import(_clipper_Context, csvImportPath, dataModelstring);
                                             }
                                             clipper_modelsRepository = null;
                             
                                             EventLog.WriteEntry("Clipper_import", "Import du stock terminé", EventLogEntryType.Information, 255);
                                             break;
                             
                                         case "STOCK_PURGE":
                                             //puge de tous les elements du stock
                                             // clipper_modelsRepository =  new ModelsRepository();
                                             //_clipper_Context = clipper_modelsRepository.GetModelContext(DbName);  //nom de la base;
                                              //int i = _clipper_Context.ModelsRepository.ModelList.Count();

                                             IEntityList stocks = _clipper_Context.EntityManager.GetEntityList("_STOCK");
                                                stocks.Fill(false);
            
                                                foreach ( IEntity stock in stocks)
                                                {
                                                stock.Delete();
                                                }


                                                IEntityList formats = _clipper_Context.EntityManager.GetEntityList("_SHEET");
                                                formats.Fill(false);

                                                foreach (IEntity format in formats)
                                                {
                                                format.Delete();
                                                }
                                                clipper_modelsRepository = null;
                                             break;

                                         case "OF":
                                             //import stock
                                             clipper_modelsRepository = new ModelsRepository();
                                             //import of
                                             //chargement de sparamètres
                                             //Clipper_Param.GetlistParam(_clipper_Context);
                                             string csvImport_of_Path = Clipper_Param.GetPath("IMPORT_CDA");
                                             string of_dataModelstring = Clipper_Param.GetModelCA();

                                             using (Clipper_OF CahierAffaire = new Clipper_OF())
                                             {

                                                 CahierAffaire.Import(_clipper_Context, csvImport_of_Path, of_dataModelstring, false);
                                             }



                                             //
                                             //csv
                                             string csvfile = System.IO.Path.GetFileNameWithoutExtension(csvImport_of_Path);
                                             dataModelstring = Clipper_Param.GetModelCA();
                                             string csvImportPath_sans_dt = csvImport_of_Path.Substring(0, (csvImport_of_Path.Length - (csvfile.Length + 4))) + csvfile + "_SANSDT.csv";
                                             using (Clipper_OF CahierAffaire_sans_Dt = new Clipper_OF())
                                             {


                                                 CahierAffaire_sans_Dt.Import(_clipper_Context, csvImportPath_sans_dt, dataModelstring, true);

                                             }

                                             EventLog.WriteEntry("Clipper_import", "Import du ca terminé" + csvImportPath_sans_dt, EventLogEntryType.Information, 255);
                                             clipper_modelsRepository = null;
                                             break;



                                     }

                     
                     
                                 }
                                 else
                                 {


                                     EventLog.WriteEntry("Clipper_import", "impossible de trouver les arguments demandés <type import> <nomfichier>", EventLogEntryType.Warning, 255);
                                 }


             

                             } 


            }
    }

  }


