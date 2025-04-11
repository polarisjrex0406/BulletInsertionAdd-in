using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Xml.Linq;
using PowerPoint = Microsoft.Office.Interop.PowerPoint;
using Office = Microsoft.Office.Core;
using Microsoft.Office.Interop.PowerPoint;
using Microsoft.Office.Core;

namespace BulletInsertionAdd_in
{
    public partial class ThisAddIn
    {
        private RibbonController _ribbonController;
        private void ThisAddIn_Startup(object sender, System.EventArgs e)
        {
            Application.WindowSelectionChange += OnSelectionChanged;
        }

        private void OnSelectionChanged(PowerPoint.Selection selection)
        {
            bool isEnabled = CheckSelection(selection);
            _ribbonController.EnableBulletInsertion(isEnabled);
        }

        private bool CheckSelection(PowerPoint.Selection selection)
        {
            // Combine checks from Steps 2 and 3
            //return (IsMultipleTextboxes(selection) || IsMultiLineText(selection));
            return IsMultiLineText(selection);
        }

        private bool IsMultipleTextboxes(PowerPoint.Selection selection)
        {
            if (selection.Type == PowerPoint.PpSelectionType.ppSelectionShapes)
            {
                // Check if more than one shape is selected
                if (selection.ShapeRange.Count > 1)
                {
                    bool allTextboxes = true;
                    foreach (PowerPoint.Shape shape in selection.ShapeRange)
                    {
                        // Verify each shape is a textbox with text
                        if (shape.HasTextFrame != Office.MsoTriState.msoTrue ||
                            shape.TextFrame.HasText != Office.MsoTriState.msoTrue)
                        {
                            allTextboxes = false;
                            break;
                        }
                    }
                    return allTextboxes; // True if multiple valid textboxes selected
                }
            }

            return false;
        }

        private bool IsMultiLineText(PowerPoint.Selection selection)
        {
            return selection.Type == PowerPoint.PpSelectionType.ppSelectionText;
        }

        public void InsertBullet(string bulletType)
        {
            PowerPoint.Selection selection = Application.ActiveWindow.Selection;
            if (selection.Type != PowerPoint.PpSelectionType.ppSelectionText)
            {
                System.Windows.Forms.MessageBox.Show("Please select text to insert a bullet.");
                return;
            }

            PowerPoint.TextRange textRange = selection.TextRange;
            PowerPoint.ParagraphFormat paragraphFormat = textRange.ParagraphFormat;

            paragraphFormat.Bullet.Visible = MsoTriState.msoTrue;
            paragraphFormat.Bullet.Type = PpBulletType.ppBulletUnnumbered;
            paragraphFormat.Bullet.UseTextColor = MsoTriState.msoFalse;
            paragraphFormat.Bullet.UseTextFont = MsoTriState.msoFalse;
            paragraphFormat.Bullet.Font.Color.RGB = 255;
            paragraphFormat.Bullet.Font.Bold = MsoTriState.msoFalse;
            paragraphFormat.Bullet.Font.Italic = MsoTriState.msoFalse;

            var fontSize = textRange.Font.Size;
            paragraphFormat.Bullet.Font.Size = fontSize;

            switch (bulletType)
            {
                case "Triangle":
                    textRange.IndentLevel = 2;
                    paragraphFormat.Bullet.Font.Name = "Wingdings 3";
                    paragraphFormat.Bullet.Character = 132;
                    break;
                case "Circle":
                    textRange.IndentLevel = 3;
                    paragraphFormat.Bullet.Character = 8226;                    
                    break;
                case "Dash":
                    textRange.IndentLevel = 4;
                    paragraphFormat.Bullet.Character = 8211;
                    break;
            }

            textRange.Font.Size = fontSize;
        }

        private void ThisAddIn_Shutdown(object sender, System.EventArgs e)
        {
        }

        protected override Microsoft.Office.Core.IRibbonExtensibility CreateRibbonExtensibilityObject()
        {
            _ribbonController = new RibbonController();
            return _ribbonController;
        }


        #region VSTO generated code

        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InternalStartup()
        {
            this.Startup += new System.EventHandler(ThisAddIn_Startup);
            this.Shutdown += new System.EventHandler(ThisAddIn_Shutdown);
        }
        
        #endregion
    }
}
