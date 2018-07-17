using Outlook = Microsoft.Office.Interop.Outlook;
using Office = Microsoft.Office.Core;

namespace OutLookAddin
{
    public partial class ThisAddIn
    {
        Outlook.Inspectors _inspectors;



        void Inspectors_NewInspector(Outlook.Inspector inspector)
        {
            if (inspector.CurrentItem is Outlook.MailItem mailItem)
            {
                if (mailItem.EntryID == null)
                {
                    mailItem.Subject = @"Mal subject added by using Outlook add-in";
                    mailItem.Body = @"Mal body added by using Outlook add-in";
                }

            }
        }
        private void ThisAddIn_Startup(object sender, System.EventArgs e)
        {
            _inspectors = Application.Inspectors;
            _inspectors.NewInspector +=
                Inspectors_NewInspector;
        }
        private void ThisAddIn_Shutdown(object sender, System.EventArgs e)
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
            Startup += ThisAddIn_Startup;
            Shutdown += ThisAddIn_Shutdown;
        }
        
        #endregion
    }
}
