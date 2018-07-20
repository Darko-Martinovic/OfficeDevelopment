using System;
using System.IO;
using System.Reflection;
using System.Runtime.InteropServices;
using Office = Microsoft.Office.Core;
using Word = Microsoft.Office.Interop.Word;

// TODO:  Follow these steps to enable the Ribbon (XML) item:

// 1: Copy the following code block into the ThisAddin, ThisWorkbook, or ThisDocument class.

//  protected override Microsoft.Office.Core.IRibbonExtensibility CreateRibbonExtensibilityObject()
//  {
//      return new MyRibbon();
//  }

// 2. Create callback methods in the "Ribbon Callbacks" region of this class to handle user
//    actions, such as clicking a button. Note: if you have exported this Ribbon from the Ribbon designer,
//    move your code from the event handlers to the callback methods and modify the code to work with the
//    Ribbon extensibility (RibbonX) programming model.

// 3. Assign attributes to the control tags in the Ribbon XML file to identify the appropriate callback methods in your code.  

// For more information, see the Ribbon XML documentation in the Visual Studio Tools for Office Help.


namespace TestWordAddIn
{
    [ComVisible(true)]
    public class MyRibbon : Office.IRibbonExtensibility
    {
        public void OnTextButton(Office.IRibbonControl control)
        {
            var currentRange = Globals.ThisAddIn.Application.Selection.Range;
            currentRange.Text = "This text was added by the Ribbon.";
        }
        public void OnTableButton(Office.IRibbonControl control)
        {
            object missing = Type.Missing;
            Word.Range currentRange = Globals.ThisAddIn.Application.Selection.Range;
            Word.Table newTable = Globals.ThisAddIn.Application.ActiveDocument.Tables.Add(currentRange, 3, 4, ref missing, ref missing);
            // Get all of the borders except for the diagonal borders.
            Word.Border[] borders = new Word.Border[6];
            borders[0] = newTable.Borders[Word.WdBorderType.wdBorderLeft];
            borders[1] = newTable.Borders[Word.WdBorderType.wdBorderRight];
            borders[2] = newTable.Borders[Word.WdBorderType.wdBorderTop];
            borders[3] = newTable.Borders[Word.WdBorderType.wdBorderBottom];
            borders[4] = newTable.Borders[Word.WdBorderType.wdBorderHorizontal];
            borders[5] = newTable.Borders[Word.WdBorderType.wdBorderVertical];
            // Format each of the borders.
            foreach (Word.Border border in borders)
            {
                border.LineStyle = Word.WdLineStyle.wdLineStyleSingle;
                border.Color = Word.WdColor.wdColorBlue;
            }

        }


            #region IRibbonExtensibility Members

            public string GetCustomUI(string ribbonId)
        {
            return GetResourceText("TestWordAddIn.MyRibbon.xml");
        }

        #endregion

        #region Ribbon Callbacks
        //Create callback methods here. For more information about adding callback methods, visit https://go.microsoft.com/fwlink/?LinkID=271226

        public void Ribbon_Load(Office.IRibbonUI ribbonUi)
        {
        }

        #endregion

        #region Helpers

        private static string GetResourceText(string resourceName)
        {
            var asm = Assembly.GetExecutingAssembly();
            var resourceNames = asm.GetManifestResourceNames();
            foreach (var t in resourceNames)
            {
                if (string.Compare(resourceName, t, StringComparison.OrdinalIgnoreCase) != 0) continue;
                using (var resourceReader = new StreamReader(asm.GetManifestResourceStream(t) ?? throw new InvalidOperationException()))
                {
                    return resourceReader.ReadToEnd();
                }
            }
            return null;
        }

        #endregion
    }
}
