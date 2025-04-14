using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Xml.Linq;
using PowerPoint = Microsoft.Office.Interop.PowerPoint;
using Office = Microsoft.Office.Core;
using Microsoft.Office.Interop.PowerPoint;
using Microsoft.Office.Core;
using System.Reflection;
using System.Diagnostics;

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
            return IsMultiLineText(selection);
        }

        private bool IsMultiLineText(PowerPoint.Selection selection)
        {
            return selection.Type == PowerPoint.PpSelectionType.ppSelectionText;
        }

        public void InsertBullet(string bulletType)
        {
            if (bulletType != "Triangle" && bulletType != "Circle" && bulletType != "Dash")
            {
                return;
            }

            PowerPoint.Selection selection = Application.ActiveWindow.Selection;
            if (selection.Type != PowerPoint.PpSelectionType.ppSelectionText)
            {
                System.Windows.Forms.MessageBox.Show("Please select text to insert a bullet.");
                return;
            }

            PowerPoint.TextRange textRange = selection.TextRange;
            PowerPoint.ParagraphFormat paragraphFormat = textRange.ParagraphFormat;
            PowerPoint.TextFrame textFrame = textRange.Parent; // TextRange's parent is TextFrame
            PowerPoint.Ruler ruler = textFrame.Ruler;
            PowerPoint.RulerLevels levels = ruler.Levels;

            textRange.Paragraphs(1).IndentLevel = 1;

            paragraphFormat.Bullet.Visible = MsoTriState.msoTrue;
            paragraphFormat.Bullet.Type = PpBulletType.ppBulletUnnumbered;
            paragraphFormat.Bullet.UseTextColor = MsoTriState.msoFalse;
            paragraphFormat.Bullet.UseTextFont = MsoTriState.msoFalse;
            paragraphFormat.Bullet.Font.Color.RGB = 255;
            paragraphFormat.Bullet.Font.Bold = MsoTriState.msoFalse;
            paragraphFormat.Bullet.Font.Italic = MsoTriState.msoFalse;

            float fontSize = textRange.Font.Size;
            int indentIndex = 1;
            switch (bulletType)
            {
                case "Triangle":
                    indentIndex = 2;
                    paragraphFormat.Bullet.Font.Name = "Wingdings 3";
                    paragraphFormat.Bullet.Character = 132;
                    paragraphFormat.Bullet.Font.Size = fontSize * 0.7f;
                    //paragraphFormat.Bullet.RelativeSize = 0.7f;
                    break;
                case "Circle":
                    indentIndex = 3;
                    paragraphFormat.Bullet.Character = 8226;
                    paragraphFormat.Bullet.Font.Size = fontSize;
                    //paragraphFormat.Bullet.RelativeSize = 1.0f;
                    break;
                case "Dash":
                    indentIndex = 4;
                    paragraphFormat.Bullet.Character = 8211;
                    paragraphFormat.Bullet.Font.Size = fontSize;
                    //paragraphFormat.Bullet.RelativeSize = 1.0f;
                    break;
            }

            float bulletSize = fontSize;
            float baseMargin = bulletSize * 1.2f; // 1.2x bullet size as margin

            // Explicitly set parent level margins if needed
            levels[1].FirstMargin = 0;
            levels[1].LeftMargin = 0;
            for (int i = 2; i < indentIndex; i++)
            {
                levels[i].FirstMargin = baseMargin * 0.8f * i; // Adjust proportionally
                levels[i].LeftMargin = levels[i].FirstMargin * 1.5f;
            }

            levels[indentIndex].FirstMargin = baseMargin;
            levels[indentIndex].LeftMargin = baseMargin * 1.5f; // 1.5x for hanging indent

            textRange.IndentLevel = indentIndex;
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
