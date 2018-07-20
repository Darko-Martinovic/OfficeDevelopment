using System;
using System.Drawing;
using System.IO;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Windows.Forms;
using Office = Microsoft.Office.Core;
using Word = Microsoft.Office.Interop.Word;



namespace TestWordAddIn
{
    [ComVisible(true)]
    public class MyRibbon : Office.IRibbonExtensibility
    {


        /// <summary>
        /// Test spelling
        /// </summary>
        /// <param name="control"></param>
        public void OnSpellingButton(Office.IRibbonControl control)
        {
            object startLocation = Globals.ThisAddIn.Application.ActiveDocument.Content.Start;
            object endLocation = Globals.ThisAddIn.Application.ActiveDocument.Content.End;

            var spellCheck = Globals.ThisAddIn.Application.CheckSpelling(Globals.ThisAddIn.Application.ActiveDocument
                .Range(ref startLocation, ref endLocation).Text);

            MessageBox.Show(spellCheck ? @"Everything OK" : @"There are mistakes");
        }

        public void OnNumberOfWords(Office.IRibbonControl control)
        {

            var rng = Globals.ThisAddIn.Application.ActiveDocument.Content;
            rng.Select();


            MessageBox.Show($@"Characters including hidden in document : {Globals.ThisAddIn.Application.ActiveDocument.Characters.Count}",@"Info",MessageBoxButtons.OK,MessageBoxIcon.Information);



        }
        public void OnTableButton(Office.IRibbonControl control)
        {
            var missing = Type.Missing;
            var currentRange = Globals.ThisAddIn.Application.Selection.Range;
            var newTable = Globals.ThisAddIn.Application.ActiveDocument.Tables.Add(currentRange, 3, 4, ref missing, ref missing);
            // Get all of the borders except for the diagonal borders.
            var borders = new Word.Border[6];
            borders[0] = newTable.Borders[Word.WdBorderType.wdBorderLeft];
            borders[1] = newTable.Borders[Word.WdBorderType.wdBorderRight];
            borders[2] = newTable.Borders[Word.WdBorderType.wdBorderTop];
            borders[3] = newTable.Borders[Word.WdBorderType.wdBorderBottom];
            borders[4] = newTable.Borders[Word.WdBorderType.wdBorderHorizontal];
            borders[5] = newTable.Borders[Word.WdBorderType.wdBorderVertical];
            // Format each of the borders.
            foreach (var border in borders)
            {
                border.LineStyle = Word.WdLineStyle.wdLineStyleSingle;
                border.Color = Word.WdColor.wdColorBlue;
            }

        }

        public Bitmap spell_get(Office.IRibbonControl control)
        {
            return Properties.Resources.spell1;
        }

        public Bitmap word_get(Office.IRibbonControl control)
        {
            return Properties.Resources.word;
        }
        public Bitmap table_get(Office.IRibbonControl control)
        {
            return Properties.Resources.table;
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
