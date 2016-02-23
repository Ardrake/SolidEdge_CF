using System;
using System.Reflection;
using System.Collections.Generic;
using System.Runtime.InteropServices;
using SolidEdgeDraft;
using System.IO;

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

            //OuvrirAsm();
            IterateFolder();
            // faire la magie icitte
            //CreateDraft();
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

        static void OuvrirAsm(string fname)
        {
            //string fName = @"K:\PROJET_IMAGE_WEB\PORTE_2016\P-04A\32X73\P04.asm";
            Console.WriteLine("Creation de document test");
            documents = application.Documents;
            assembly = (SolidEdgeAssembly.AssemblyDocument)documents.Open(fname);

            SolidEdgeFramework.Window window = (SolidEdgeFramework.Window)application.ActiveWindow;
            window.View.Fit();
            SaveAsImage(window, Path.GetDirectoryName(fname));
            assembly.Save();
            assembly.Close();
            documents.Close();
            Marshal.FinalReleaseComObject(documents);
            documents = null;
            application.DoIdle();

        }

        static void IterateFolder()
        {
            string folder = @"K:\PROJET_IMAGE_WEB\PORTE_2016\";
            IEnumerable<string> fnames = Directory.EnumerateFiles(folder, "*.*", SearchOption.AllDirectories);
            
            foreach (string items in fnames )
            {
                if (Path.GetExtension(items) == ".asm" && Path.GetFileName(items).Trim().StartsWith("P"))
                {
                    Console.WriteLine(items);
                    OuvrirAsm(items);
                }
                
            }
            Console.ReadKey();
        }

    static void SaveAsImage(SolidEdgeFramework.Window window, string folder)
        {
            //string[] extensions = { ".jpg", ".bmp", ".tif" };
            string[] extensions = { ".jpg" };
            SolidEdgeFramework.View view = null;
            Guid guid = Guid.NewGuid();
            //string folder = Environment.GetFolderPath(Environment.SpecialFolder.DesktopDirectory);
            //string folder = @"K:\PROJET_IMAGE_WEB\PORTE_2016\P-04A\32X73\";

            double resolution = 1;  // DPI - Larger values have better quality but also lead to larger file. 
            int colorDepth = 24;
            //int width = window.UsableWidth;
            //int height = window.UsableHeight;
            int width = 13201;
            int height = 6494;
            // Get a reference to the 3D view. 
            view = window.View;

            // Save each extension. 
            foreach (string extension in extensions)
            {
                // File saved to desktop. 
                string filename = Path.ChangeExtension(guid.ToString(), extension);
                filename = Path.Combine(folder, filename);

                // You can specify .bmp (Windows Bitmap), .tif (TIFF), or .jpg (JPEG). 
                view.SaveAsImage(
                    Filename: filename,
                    Width: width,
                    Height: height,
                    AltViewStyle: null,
                    Resolution: resolution,
                    ColorDepth: colorDepth,
                    ImageQuality: SolidEdgeFramework.SeImageQualityType.seImageQualityHigh,
                    Invert: false);

                Console.WriteLine("Saved '{0}'.", filename);
            }
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