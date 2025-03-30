namespace WordNepaliSpellCheck
{
    partial class SpellingRibbon : Microsoft.Office.Tools.Ribbon.RibbonBase
    {
        /// <summary>
        /// Required designer variable.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        public SpellingRibbon()
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
            this.nepaliLanguageTab = this.Factory.CreateRibbonTab();
            this.nepaliSpelling = this.Factory.CreateRibbonGroup();
            this.spellingCheckButton = this.Factory.CreateRibbonToggleButton();
            this.nepaliLanguageTab.SuspendLayout();
            this.nepaliSpelling.SuspendLayout();
            this.SuspendLayout();
            // 
            // nepaliLanguageTab
            // 
            this.nepaliLanguageTab.ControlId.ControlIdType = Microsoft.Office.Tools.Ribbon.RibbonControlIdType.Office;
            this.nepaliLanguageTab.Groups.Add(this.nepaliSpelling);
            this.nepaliLanguageTab.Label = "सायक- Nepali Language Tools";
            this.nepaliLanguageTab.Name = "nepaliLanguageTab";
            // 
            // nepaliSpelling
            // 
            this.nepaliSpelling.Items.Add(this.spellingCheckButton);
            this.nepaliSpelling.Label = "Spelling Action: OFF";
            this.nepaliSpelling.Name = "nepaliSpelling";
            // 
            // spellingCheckButton
            // 
            this.spellingCheckButton.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.spellingCheckButton.Description = "Nepali Spelling Checker State";
            this.spellingCheckButton.Image = global::WordNepaliSpellCheck.Properties.Resources.checker;
            this.spellingCheckButton.Label = "नेपाली हिज्जे जाँच";
            this.spellingCheckButton.Name = "spellingCheckButton";
            this.spellingCheckButton.ShowImage = true;
            this.spellingCheckButton.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.spellingCheckButton_Click);
            // 
            // SpellingRibbon
            // 
            this.Name = "SpellingRibbon";
            this.RibbonType = "Microsoft.Word.Document";
            this.Tabs.Add(this.nepaliLanguageTab);
            this.Load += new Microsoft.Office.Tools.Ribbon.RibbonUIEventHandler(this.SpellingRibbon_Load);
            this.nepaliLanguageTab.ResumeLayout(false);
            this.nepaliLanguageTab.PerformLayout();
            this.nepaliSpelling.ResumeLayout(false);
            this.nepaliSpelling.PerformLayout();
            this.ResumeLayout(false);

        }

        #endregion

        internal Microsoft.Office.Tools.Ribbon.RibbonTab nepaliLanguageTab;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup nepaliSpelling;
        internal Microsoft.Office.Tools.Ribbon.RibbonToggleButton spellingCheckButton;
    }

    partial class ThisRibbonCollection
    {
        internal SpellingRibbon SpellingRibbon
        {
            get { return this.GetRibbon<SpellingRibbon>(); }
        }
    }
}
