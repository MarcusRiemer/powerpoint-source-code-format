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
            this.lblPygmentsAvailable = this.Factory.CreateRibbonLabel();
            this.lblPygmentsNotAvailable = this.Factory.CreateRibbonLabel();
            this.cmbLanguage = this.Factory.CreateRibbonComboBox();
            this.btnHelpPygmentize = this.Factory.CreateRibbonButton();
            this.btnFormatCurrent = this.Factory.CreateRibbonButton();
            this.btnFormatAll = this.Factory.CreateRibbonButton();
            this.cmbStyle = this.Factory.CreateRibbonComboBox();
            this.tab1.SuspendLayout();
            this.group1.SuspendLayout();
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
            this.group1.Items.Add(this.lblPygmentsAvailable);
            this.group1.Items.Add(this.lblPygmentsNotAvailable);
            this.group1.Items.Add(this.btnHelpPygmentize);
            this.group1.Items.Add(this.cmbLanguage);
            this.group1.Items.Add(this.cmbStyle);
            this.group1.Items.Add(this.btnFormatCurrent);
            this.group1.Items.Add(this.btnFormatAll);
            this.group1.Label = "PP Source Code";
            this.group1.Name = "group1";
            // 
            // lblPygmentsAvailable
            // 
            this.lblPygmentsAvailable.Label = "✓ Pygments is available";
            this.lblPygmentsAvailable.Name = "lblPygmentsAvailable";
            this.lblPygmentsAvailable.Visible = false;
            // 
            // lblPygmentsNotAvailable
            // 
            this.lblPygmentsNotAvailable.Label = "❌ Pygments is not available";
            this.lblPygmentsNotAvailable.Name = "lblPygmentsNotAvailable";
            this.lblPygmentsNotAvailable.Visible = false;
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
            // 
            // btnHelpPygmentize
            // 
            this.btnHelpPygmentize.Label = "Show Online Help";
            this.btnHelpPygmentize.Name = "btnHelpPygmentize";
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
            // btnFormatAll
            // 
            this.btnFormatAll.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.btnFormatAll.Image = global::pp_source_format.Properties.Resources.multiple_targets;
            this.btnFormatAll.Label = "Format All";
            this.btnFormatAll.Name = "btnFormatAll";
            this.btnFormatAll.ShowImage = true;
            this.btnFormatAll.Visible = false;
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
            // 
            // Ribbon
            // 
            this.Name = "Ribbon";
            this.RibbonType = "Microsoft.PowerPoint.Presentation";
            this.Tabs.Add(this.tab1);
            this.Load += new Microsoft.Office.Tools.Ribbon.RibbonUIEventHandler(this.Ribbon1_Load);
            this.tab1.ResumeLayout(false);
            this.tab1.PerformLayout();
            this.group1.ResumeLayout(false);
            this.group1.PerformLayout();
            this.ResumeLayout(false);

        }

        #endregion

        internal Microsoft.Office.Tools.Ribbon.RibbonTab tab1;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup group1;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnFormatCurrent;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnFormatAll;
        internal Microsoft.Office.Tools.Ribbon.RibbonComboBox cmbLanguage;
        internal Microsoft.Office.Tools.Ribbon.RibbonLabel lblPygmentsAvailable;
        internal Microsoft.Office.Tools.Ribbon.RibbonLabel lblPygmentsNotAvailable;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnHelpPygmentize;
        internal Microsoft.Office.Tools.Ribbon.RibbonComboBox cmbStyle;
    }

    partial class ThisRibbonCollection
    {
        internal Ribbon Ribbon1
        {
            get { return this.GetRibbon<Ribbon>(); }
        }
    }
}
