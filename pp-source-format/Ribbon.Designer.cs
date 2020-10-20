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
            this.tab1 = this.Factory.CreateRibbonTab();
            this.group1 = this.Factory.CreateRibbonGroup();
            this.btnFormatCurrent = this.Factory.CreateRibbonButton();
            this.btnFormatAll = this.Factory.CreateRibbonButton();
            this.cmbLanguage = this.Factory.CreateRibbonComboBox();
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
            this.group1.Items.Add(this.btnFormatCurrent);
            this.group1.Items.Add(this.btnFormatAll);
            this.group1.Items.Add(this.cmbLanguage);
            this.group1.Label = "PP Source Code";
            this.group1.Name = "group1";
            // 
            // btnFormatCurrent
            // 
            this.btnFormatCurrent.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.btnFormatCurrent.Image = global::pp_source_format.Properties.Resources.single_target;
            this.btnFormatCurrent.Label = "Format Current";
            this.btnFormatCurrent.Name = "btnFormatCurrent";
            this.btnFormatCurrent.ShowImage = true;
            this.btnFormatCurrent.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.OnRenameSingle);
            // 
            // btnFormatAll
            // 
            this.btnFormatAll.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.btnFormatAll.Image = global::pp_source_format.Properties.Resources.multiple_targets;
            this.btnFormatAll.Label = "Format All";
            this.btnFormatAll.Name = "btnFormatAll";
            this.btnFormatAll.ShowImage = true;
            // 
            // cmbLanguage
            // 
            ribbonDropDownItemImpl1.Label = "java";
            ribbonDropDownItemImpl2.Label = "c";
            this.cmbLanguage.Items.Add(ribbonDropDownItemImpl1);
            this.cmbLanguage.Items.Add(ribbonDropDownItemImpl2);
            this.cmbLanguage.Label = "Language";
            this.cmbLanguage.Name = "cmbLanguage";
            this.cmbLanguage.Text = null;
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
    }

    partial class ThisRibbonCollection
    {
        internal Ribbon Ribbon1
        {
            get { return this.GetRibbon<Ribbon>(); }
        }
    }
}
