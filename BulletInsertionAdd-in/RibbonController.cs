using Microsoft.Office.Core;
using System;
using System.Collections.Generic;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Text;
using Office = Microsoft.Office.Core;

// TODO:  Follow these steps to enable the Ribbon (XML) item:

// 1: Copy the following code block into the ThisAddin, ThisWorkbook, or ThisDocument class.

//  protected override Microsoft.Office.Core.IRibbonExtensibility CreateRibbonExtensibilityObject()
//  {
//      return new RibbonController();
//  }

// 2. Create callback methods in the "Ribbon Callbacks" region of this class to handle user
//    actions, such as clicking a button. Note: if you have exported this Ribbon from the Ribbon designer,
//    move your code from the event handlers to the callback methods and modify the code to work with the
//    Ribbon extensibility (RibbonX) programming model.

// 3. Assign attributes to the control tags in the Ribbon XML file to identify the appropriate callback methods in your code.  

// For more information, see the Ribbon XML documentation in the Visual Studio Tools for Office Help.


namespace BulletInsertionAdd_in
{
    [ComVisible(true)]
    public class RibbonController : Office.IRibbonExtensibility
    {
        private Office.IRibbonUI ribbon;

        private bool _enableBulletInsertion;

        public RibbonController()
        {
            _enableBulletInsertion = false;
        }

        public void EnableBulletInsertion(bool isEnable)
        {
            _enableBulletInsertion = isEnable;
            ribbon.Invalidate();
        }

        public void BtnBulletTriangle_OnAction(Microsoft.Office.Core.IRibbonControl control)
        {
            Globals.ThisAddIn.InsertBullet("Triangle");
        }

        public void BtnBulletCircle_OnAction(Microsoft.Office.Core.IRibbonControl control)
        {
            Globals.ThisAddIn.InsertBullet("Circle");
        }

        public void BtnBulletDash_OnAction(Microsoft.Office.Core.IRibbonControl control)
        {
            Globals.ThisAddIn.InsertBullet("Dash");
        }

        public bool BulletInsertion_GetEnabled(Office.IRibbonControl control)
        {
            return _enableBulletInsertion;
        }

        public Bitmap ImageTriangle_Icon_GetImage(IRibbonControl control)
        {
            return GetImage("Triangle.png");
        }

        public Bitmap ImageCircle_Icon_GetImage(IRibbonControl control)
        {
            return GetImage("Circle.png");
        }

        public Bitmap ImageDash_Icon_GetImage(IRibbonControl control)
        {
            return GetImage("Dash.png");
        }

        public Bitmap GetImage(string filename)
        {
            // Use the assembly to get the resource stream for the image
            var assembly = Assembly.GetExecutingAssembly();
            var resourceName = $"BulletInsertionAdd_in.Res.{filename}"; // Adjust namespace accordingly

            using (Stream stream = assembly.GetManifestResourceStream(resourceName))
            {
                if (stream != null)
                {
                    return new Bitmap(stream);
                }
            }

            return null; // Return null if the image is not found
        }

        #region IRibbonExtensibility Members

        public string GetCustomUI(string ribbonID)
        {
            return GetResourceText("BulletInsertionAdd_in.RibbonController.xml");
        }

        #endregion

        #region Ribbon Callbacks
        //Create callback methods here. For more information about adding callback methods, visit https://go.microsoft.com/fwlink/?LinkID=271226

        public void Ribbon_Load(Office.IRibbonUI ribbonUI)
        {
            this.ribbon = ribbonUI;
        }

        #endregion

        #region Helpers

        private static string GetResourceText(string resourceName)
        {
            Assembly asm = Assembly.GetExecutingAssembly();
            string[] resourceNames = asm.GetManifestResourceNames();
            for (int i = 0; i < resourceNames.Length; ++i)
            {
                if (string.Compare(resourceName, resourceNames[i], StringComparison.OrdinalIgnoreCase) == 0)
                {
                    using (StreamReader resourceReader = new StreamReader(asm.GetManifestResourceStream(resourceNames[i])))
                    {
                        if (resourceReader != null)
                        {
                            return resourceReader.ReadToEnd();
                        }
                    }
                }
            }
            return null;
        }

        #endregion
    }
}
