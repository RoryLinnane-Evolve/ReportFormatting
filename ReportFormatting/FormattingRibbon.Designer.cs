
namespace ReportFormatting
{
    partial class FormattingRibbon : Microsoft.Office.Tools.Ribbon.RibbonBase
    {
        /// <summary>
        /// Required designer variable.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        public FormattingRibbon()
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
            this.FormattingTab = this.Factory.CreateRibbonTab();
            this.group1 = this.Factory.CreateRibbonGroup();
            this.btnSelectImages = this.Factory.CreateRibbonButton();
            this.btnSelectFigures = this.Factory.CreateRibbonButton();
            this.btnFormatLines = this.Factory.CreateRibbonButton();
            this.FormattingTab.SuspendLayout();
            this.group1.SuspendLayout();
            this.SuspendLayout();
            // 
            // FormattingTab
            // 
            this.FormattingTab.ControlId.ControlIdType = Microsoft.Office.Tools.Ribbon.RibbonControlIdType.Office;
            this.FormattingTab.Groups.Add(this.group1);
            this.FormattingTab.Label = "Formatting";
            this.FormattingTab.Name = "FormattingTab";
            // 
            // group1
            // 
            this.group1.Items.Add(this.btnSelectImages);
            this.group1.Items.Add(this.btnSelectFigures);
            this.group1.Items.Add(this.btnFormatLines);
            this.group1.Name = "group1";
            // 
            // btnSelectImages
            // 
            this.btnSelectImages.Label = "Select Plates";
            this.btnSelectImages.Name = "btnSelectImages";
            this.btnSelectImages.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnSelectImages_Click);
            // 
            // btnSelectFigures
            // 
            this.btnSelectFigures.Label = "Select Figures";
            this.btnSelectFigures.Name = "btnSelectFigures";
            this.btnSelectFigures.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnSelectFigures_Click);
            // 
            // btnFormatLines
            // 
            this.btnFormatLines.Label = "Format Lines";
            this.btnFormatLines.Name = "btnFormatLines";
            this.btnFormatLines.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnFormatLines_Click);
            // 
            // FormattingRibbon
            // 
            this.Name = "FormattingRibbon";
            this.RibbonType = "Microsoft.Word.Document";
            this.Tabs.Add(this.FormattingTab);
            this.Load += new Microsoft.Office.Tools.Ribbon.RibbonUIEventHandler(this.FormattingRibbon_Load);
            this.FormattingTab.ResumeLayout(false);
            this.FormattingTab.PerformLayout();
            this.group1.ResumeLayout(false);
            this.group1.PerformLayout();
            this.ResumeLayout(false);

        }

        #endregion

        internal Microsoft.Office.Tools.Ribbon.RibbonTab FormattingTab;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup group1;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnSelectImages;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnSelectFigures;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnFormatLines;
    }

    partial class ThisRibbonCollection
    {
        internal FormattingRibbon FormattingRibbon
        {
            get { return this.GetRibbon<FormattingRibbon>(); }
        }
    }
}
