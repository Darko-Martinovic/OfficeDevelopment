

namespace SimpleTalkExcellAddin
{
    partial class UsingRibbon : Microsoft.Office.Tools.Ribbon.RibbonBase
    {
        /// <summary>
        /// Required designer variable.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        public UsingRibbon()
            : base(Globals.Factory.GetRibbonFactory())
        {
            InitializeComponent();
        }

        /// <summary> 
        /// Clean up any resources being used.
        /// </summary>
        /// <param name="disposing">true if managed resources should be disposed; otherwise, false.</param>
        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        #region Component Designer generated code

        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InitializeComponent()
        {
            this.tab1 = this.Factory.CreateRibbonTab();
            this.group1 = this.Factory.CreateRibbonGroup();
            this.btnCaptureUserInput = this.Factory.CreateRibbonButton();
            this.group2 = this.Factory.CreateRibbonGroup();
            this.btnNoInput = this.Factory.CreateRibbonButton();
            this.tab1.SuspendLayout();
            this.group1.SuspendLayout();
            this.group2.SuspendLayout();
            this.SuspendLayout();
            // 
            // tab1
            // 
            this.tab1.Groups.Add(this.group1);
            this.tab1.Groups.Add(this.group2);
            this.tab1.Label = "Simple Talk";
            this.tab1.Name = "tab1";
            // 
            // group1
            // 
            this.group1.Items.Add(this.btnCaptureUserInput);
            this.group1.Label = "Test user input";
            this.group1.Name = "group1";
            // 
            // btnCaptureUserInput
            // 
            this.btnCaptureUserInput.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.btnCaptureUserInput.Image = global::SimpleTalkExcellAddin.Properties.Resources.test;
            this.btnCaptureUserInput.Label = "With User Input";
            this.btnCaptureUserInput.Name = "btnCaptureUserInput";
            this.btnCaptureUserInput.ScreenTip = "Popup message bellow  icon";
            this.btnCaptureUserInput.ShowImage = true;
            this.btnCaptureUserInput.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.button1_Click);
            // 
            // group2
            // 
            this.group2.Items.Add(this.btnNoInput);
            this.group2.Label = "No user input";
            this.group2.Name = "group2";
            // 
            // btnNoInput
            // 
            this.btnNoInput.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.btnNoInput.Image = global::SimpleTalkExcellAddin.Properties.Resources.test2;
            this.btnNoInput.Label = "No Input";
            this.btnNoInput.Name = "btnNoInput";
            this.btnNoInput.ShowImage = true;
            this.btnNoInput.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnNoInput_Click);
            // 
            // UsingRibbon
            // 
            this.Name = "UsingRibbon";
            this.RibbonType = "Microsoft.Excel.Workbook";
            this.Tabs.Add(this.tab1);
            this.tab1.ResumeLayout(false);
            this.tab1.PerformLayout();
            this.group1.ResumeLayout(false);
            this.group1.PerformLayout();
            this.group2.ResumeLayout(false);
            this.group2.PerformLayout();
            this.ResumeLayout(false);

        }

        #endregion

        internal Microsoft.Office.Tools.Ribbon.RibbonTab tab1;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup group1;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnCaptureUserInput;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup group2;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnNoInput;
    }

}
