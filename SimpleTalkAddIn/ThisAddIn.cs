namespace SimpleTalkExcellAddin
{
    public partial class ThisAddIn
    {
        private void ThisAddInStartup(object sender, System.EventArgs e)
        {
            //put your code here
            //if (Debugger.IsAttached)
            //    Debugger.Break();
        }

        private void ThisAddInShutdown(object sender, System.EventArgs e)
        {
            //put your code here
            //if (Debugger.IsAttached)
            //    Debugger.Break();

        }

        #region VSTO generated code

        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InternalStartup()
        {
            Startup += ThisAddInStartup;
            Shutdown += ThisAddInShutdown;
        }
        
        #endregion
    }
}
