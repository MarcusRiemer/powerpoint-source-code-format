namespace pp_source_format
{
    partial class Ribbon : Microsoft.Office.Tools.Ribbon.RibbonBase
    {
        /// <summary>
        /// Required designer variable.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        public Ribbon()
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
            Microsoft.Office.Tools.Ribbon.RibbonDropDownItem ribbonDropDownItemImpl1 = this.Factory.CreateRibbonDropDownItem();
            Microsoft.Office.Tools.Ribbon.RibbonDropDownItem ribbonDropDownItemImpl2 = this.Factory.CreateRibbonDropDownItem();
            Microsoft.Office.Tools.Ribbon.RibbonDropDownItem ribbonDropDownItemImpl3 = this.Factory.CreateRibbonDropDownItem();
            Microsoft.Office.Tools.Ribbon.RibbonDropDownItem ribbonDropDownItemImpl4 = this.Factory.CreateRibbonDropDownItem();
            this.tab1 = this.Factory.CreateRibbonTab();
            this.group1 = this.Factory.CreateRibbonGroup();
            this.bxUnavailable = this.Factory.CreateRibbonBox();
            this.lblPygmentsNotAvailable = this.Factory.CreateRibbonLabel();
            this.btnHelpPygmentize = this.Factory.CreateRibbonButton();
            this.bxAvailable = this.Factory.CreateRibbonBox();
            this.lblPygmentsAvailable = this.Factory.CreateRibbonLabel();
            this.cmbLanguage = this.Factory.CreateRibbonComboBox();
            this.cmbStyle = this.Factory.CreateRibbonComboBox();
            this.btnFormatCurrent = this.Factory.CreateRibbonButton();
            this.tab1.SuspendLayout();
            this.group1.SuspendLayout();
            this.bxUnavailable.SuspendLayout();
            this.bxAvailable.SuspendLayout();
            this.SuspendLayout();
            // 
            // tab1
            // 
            this.tab1.ControlId.ControlIdType = Microsoft.Office.Tools.Ribbon.RibbonControlIdType.Office;
            this.tab1.Groups.Add(this.group1);
            this.tab1.Label = "TabAddIns";
            this.tab1.Name = "tab1";
            // 
            // group1
            // 
            this.group1.Items.Add(this.bxUnavailable);
            this.group1.Items.Add(this.bxAvailable);
            this.group1.Label = "PP Source Code";
            this.group1.Name = "group1";
            // 
            // bxUnavailable
            // 
            this.bxUnavailable.BoxStyle = Microsoft.Office.Tools.Ribbon.RibbonBoxStyle.Vertical;
            this.bxUnavailable.Items.Add(this.lblPygmentsNotAvailable);
            this.bxUnavailable.Items.Add(this.btnHelpPygmentize);
            this.bxUnavailable.Name = "bxUnavailable";
            // 
            // lblPygmentsNotAvailable
            // 
            this.lblPygmentsNotAvailable.Label = "❌ Pygments is not available";
            this.lblPygmentsNotAvailable.Name = "lblPygmentsNotAvailable";
            // 
            // btnHelpPygmentize
            // 
            this.btnHelpPygmentize.Label = "Show Online Help";
            this.btnHelpPygmentize.Name = "btnHelpPygmentize";
            this.btnHelpPygmentize.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.OnShowOnlineHelp);
            // 
            // bxAvailable
            // 
            this.bxAvailable.BoxStyle = Microsoft.Office.Tools.Ribbon.RibbonBoxStyle.Vertical;
            this.bxAvailable.Items.Add(this.lblPygmentsAvailable);
            this.bxAvailable.Items.Add(this.cmbLanguage);
            this.bxAvailable.Items.Add(this.cmbStyle);
            this.bxAvailable.Items.Add(this.btnFormatCurrent);
            this.bxAvailable.Name = "bxAvailable";
            // 
            // lblPygmentsAvailable
            // 
            this.lblPygmentsAvailable.Label = "✓ Pygments is available";
            this.lblPygmentsAvailable.Name = "lblPygmentsAvailable";
            // 
            // cmbLanguage
            // 
            ribbonDropDownItemImpl1.Label = "java";
            ribbonDropDownItemImpl2.Label = "c";
            this.cmbLanguage.Items.Add(ribbonDropDownItemImpl1);
            this.cmbLanguage.Items.Add(ribbonDropDownItemImpl2);
            this.cmbLanguage.Label = "Language";
            this.cmbLanguage.Name = "cmbLanguage";
            this.cmbLanguage.Text = "java";
            this.cmbLanguage.TextChanged += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.OnLanguageChanged);
            // 
            // cmbStyle
            // 
            ribbonDropDownItemImpl3.Label = "default";
            ribbonDropDownItemImpl4.Label = "vs";
            this.cmbStyle.Items.Add(ribbonDropDownItemImpl3);
            this.cmbStyle.Items.Add(ribbonDropDownItemImpl4);
            this.cmbStyle.Label = "Style";
            this.cmbStyle.Name = "cmbStyle";
            this.cmbStyle.Text = "default";
            this.cmbStyle.TextChanged += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.OnStyleChanged);
            // 
            // btnFormatCurrent
            // 
            this.btnFormatCurrent.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.btnFormatCurrent.Image = global::pp_source_format.Properties.Resources.single_target;
            this.btnFormatCurrent.Label = "Format Selected";
            this.btnFormatCurrent.Name = "btnFormatCurrent";
            this.btnFormatCurrent.ShowImage = true;
            this.btnFormatCurrent.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.OnFormatSelected);
            // 
            // Ribbon
            // 
            this.Name = "Ribbon";
            this.RibbonType = "Microsoft.PowerPoint.Presentation";
            this.Tabs.Add(this.tab1);
            this.Load += new Microsoft.Office.Tools.Ribbon.RibbonUIEventHandler(this.RibbonLoad);
            this.tab1.ResumeLayout(false);
            this.tab1.PerformLayout();
            this.group1.ResumeLayout(false);
            this.group1.PerformLayout();
            this.bxUnavailable.ResumeLayout(false);
            this.bxUnavailable.PerformLayout();
            this.bxAvailable.ResumeLayout(false);
            this.bxAvailable.PerformLayout();
            this.ResumeLayout(false);

        }

        #endregion

        internal Microsoft.Office.Tools.Ribbon.RibbonTab tab1;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup group1;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnFormatCurrent;
        internal Microsoft.Office.Tools.Ribbon.RibbonComboBox cmbLanguage;
        internal Microsoft.Office.Tools.Ribbon.RibbonLabel lblPygmentsAvailable;
        internal Microsoft.Office.Tools.Ribbon.RibbonLabel lblPygmentsNotAvailable;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnHelpPygmentize;
        internal Microsoft.Office.Tools.Ribbon.RibbonComboBox cmbStyle;
        internal Microsoft.Office.Tools.Ribbon.RibbonBox bxUnavailable;
        internal Microsoft.Office.Tools.Ribbon.RibbonBox bxAvailable;
    }

    partial class ThisRibbonCollection
    {
        internal Ribbon Ribbon1
        {
            get { return this.GetRibbon<Ribbon>(); }
        }
    }
}
