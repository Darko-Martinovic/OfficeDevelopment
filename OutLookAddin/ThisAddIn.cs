using Outlook = Microsoft.Office.Interop.Outlook;

namespace OutLookAddin
{
    public partial class ThisAddIn
    {
        Outlook.Inspectors _inspectors;


        private void InspectorsNewInspector(Outlook.Inspector inspector)
        {
            if (!(inspector.CurrentItem is Outlook.MailItem mailItem)) return;
            if (mailItem.EntryID != null) return;

            mailItem.Subject = @"Mail subject added by using Outlook VSTO";
            mailItem.Body = @"Mail body added by using Outlook VSTO";
        }
        private void ThisAddInStartup(object sender, System.EventArgs e)
        {
            _inspectors = Application.Inspectors;
            _inspectors.NewInspector +=
                InspectorsNewInspector;
        }
        private void ThisAddInShutdown(object sender, System.EventArgs e)
        {

            // Note: Outlook no longer raises this event. If you have code that 
            //    must run when Outlook shuts down, see https://go.microsoft.com/fwlink/?LinkId=506785
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
