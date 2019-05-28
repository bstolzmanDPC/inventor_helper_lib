using System;
using Inventor;

namespace InventorHelperLib
{
    static public class InventorHelper
    {

        /// <summary>
        /// Derives an .ipt from a .iam and saves it in the same directory.
        /// </summary>
        /// <param name="AssemblyFileName">Path to the target assembly file</param>
        /// <param name="OptionalProgramId">Optional program identifier (ie. FCAM, ACAD...) Saves as AssemblyFileName_OptionalProgramId.ipt</param>
        static public void DerivePartFromAssembly(string AssemblyFileName, string OptionalProgramId = "")
        {
            //get running application instance
            Application app = GetInventorObject();
            app.Visible = true;

            //Create new empty part document to derive into
            Inventor.PartDocument doc = app.Documents.Add(DocumentTypeEnum.kPartDocumentObject) as PartDocument;

            //Create derived assembly definition from assembly file
            Inventor.DerivedAssemblyDefinition def = doc.ComponentDefinition.ReferenceComponents.DerivedAssemblyComponents.CreateDefinition(AssemblyFileName);

            //Derive from definition
            doc.ComponentDefinition.ReferenceComponents.DerivedAssemblyComponents.Add(def);

            //Save document with optional program use identifier
            doc.SaveAs($"{AssemblyFileName}_{OptionalProgramId}.ipt",false);

            doc.Close();

        }

        static private Inventor.Application GetInventorObject()
        {
            try
            {
                return System.Runtime.InteropServices.Marshal.GetActiveObject("Inventor.Application") as Inventor.Application;
            }
            catch
            {
                Type appTpye = System.Type.GetTypeFromProgID("Inventor.Application");
                return System.Activator.CreateInstance(appTpye) as Inventor.Application;
            }
        }
        
    }
}
