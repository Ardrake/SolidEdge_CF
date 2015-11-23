using  System;
using  System.Reflection;
using System.Collections.Generic;
using  System.Runtime.InteropServices;
using SolidEdgeDraft;

namespace SE_interface
{
    class Program
    {
        static SolidEdgeFramework.Application application = null;
        static SolidEdgeFramework.Documents documents = null;
        static SolidEdgeAssembly.AssemblyDocument assembly = null;
        static SolidEdgeDraft.DraftDocument draft = null;
        static SolidEdgePart.PartDocument part = null;
        static Type type = null;

        static SolidEdgeDraft.Sections sections = null;
        static SolidEdgeDraft.Section section = null;
        static SolidEdgeDraft.SectionSheets sectionSheets = null;
        static SolidEdgeDraft.Sheets sheets = null;
        static SolidEdgeDraft.Sheet sheet = null;
        static string format1 = "Section = {0}";
        static string format2 = "Sheet = {0}";

        static List<string> ListeDesDraft =  new List<string> { "MUR A", "MUR B", "MUR C", "MUR D", "CHEVRONS" };

        

        static void Main(string[] args)
        {
            
            try
            {
                // Get the type from the Solid Edge ProgID
                type = Type.GetTypeFromProgID("SolidEdge.Application");
                // Connect to a running instance of Solid Edge
                application = (SolidEdgeFramework.Application)Marshal.GetActiveObject("SolidEdge.Application");
                Console.Write("Connecté a Solid Edge \n");
            }
            catch (System.Exception)
            {
                Console.WriteLine("Solid Edge non actif...");
                StartSE();
            }


            // faire la magie icitte
            CreateDraft();
            //CreateDoc();

        }


        static void StartSE()
        {
            Console.WriteLine("Démarrage de Solid Edge");
            try
            {
                type = Type.GetTypeFromProgID("SolidEdge.Application");
                application = (SolidEdgeFramework.Application)Activator.CreateInstance(type);
                // Make Solid Edge visible
                application.Visible = true;
                Console.WriteLine("Solid edge démarré");
            }
            catch (System.Exception ex)
            {
                Console.WriteLine(ex.Message);
                Console.WriteLine("Solid non installer sur ce poste");
            }
        }


        static void CleanSE()
        {
            if (application != null)
            {
                Marshal.ReleaseComObject(application);
                application = null;
            }
        }


        static void CreateDoc()
        {
            Console.WriteLine("Creation de document test");
            documents = application.Documents;
            assembly = (SolidEdgeAssembly.AssemblyDocument)documents.Add("SolidEdge.AssemblyDocument", Missing.Value);
            draft = (SolidEdgeDraft.DraftDocument)documents.Add("SolidEdge.DraftDocument", Missing.Value);
            part = (SolidEdgePart.PartDocument)documents.Add("SolidEdge.PartDocument", Missing.Value);
            CleanSE();
        }

        static void CreateDraft()
        {
            //draft = (SolidEdgeDraft.DraftDocument)documents.Add("SolidEdge.DraftDocument", Missing.Value);
            try
            {
                //add a draft document
                documents = application.Documents;
                draft = (SolidEdgeDraft.DraftDocument)documents.Add("SolidEdge.DraftDocument", Missing.Value);
                
                // Get a reference to the sheets collection
                sheets = draft.Sheets;
                
                // Add sheets to draft document
                // Loop thru list of wall - to do


                // Add a new sheet
                sheet = sheets.Item(1);
                sheet.Activate();
                sheet.Name = "MUR A";
                // Add wall drawings

                //Insert next wall
                sheet = sheets.AddSheet("MUR B", SheetSectionTypeConstants.igWorkingSection, Missing.Value, Missing.Value);
                sheet.Activate();
                // Add wall drawing

                //Insert next wall
                sheet = sheets.AddSheet("MUR C", SheetSectionTypeConstants.igWorkingSection, Missing.Value, Missing.Value);
                sheet.Activate();
                // Add wall drawings

                //Insert next wall
                sheet = sheets.AddSheet("MUR D", SheetSectionTypeConstants.igWorkingSection, Missing.Value, Missing.Value);
                sheet.Activate();
                // Add wall drawings

                //Insert next wall
                sheet = sheets.AddSheet("CHEVRONS", SheetSectionTypeConstants.igWorkingSection, Missing.Value, Missing.Value);
                sheet.Activate();
                // Add wall drawings

            }

            catch (System.Exception ex)
            {
                Console.WriteLine(ex.Message);
            }
            finally
            {
                if (sheet != null)
                {
                    Marshal.ReleaseComObject(sheet);
                    sheet = null;
                }
                if (sheets != null)
                {
                    Marshal.ReleaseComObject(sheets);
                    sheets = null;
                }
                if (sectionSheets != null)
                {
                    Marshal.ReleaseComObject(sectionSheets);
                    sectionSheets = null;
                }
                if (section != null)
                {
                    Marshal.ReleaseComObject(section);
                    section = null;
                }
                if (draft != null)
                {
                    Marshal.ReleaseComObject(draft);
                    draft = null;
                }
                if (documents != null)
                {
                    Marshal.ReleaseComObject(documents);
                    documents = null;
                }
                if (application != null)
                {
                    Marshal.ReleaseComObject(application);
                    application = null;
                }


            }




        }

           
        
    }
}