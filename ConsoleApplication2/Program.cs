using  System;
using  System.Reflection;
using  System.Runtime.InteropServices;

namespace MyMacro
{
    class Program
    {
        static SolidEdgeFramework.Application application = null;
        static SolidEdgeFramework.Documents documents = null;
        static SolidEdgeAssembly.AssemblyDocument assembly = null;
        static SolidEdgeDraft.DraftDocument draft = null;
        static SolidEdgePart.PartDocument part = null;
        static Type type = null;

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
            CreateDoc();

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
    }
}