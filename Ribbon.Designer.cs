namespace BenryPPT
{
    partial class Ribbon : Microsoft.Office.Tools.Ribbon.RibbonBase
    {
        /// <summary>
        /// 必要なデザイナー変数です。
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        public Ribbon()
            : base(Globals.Factory.GetRibbonFactory())
        {
            InitializeComponent();
        }

        /// <summary> 
        /// 使用中のリソースをすべてクリーンアップします。
        /// </summary>
        /// <param name="disposing">マネージド リソースを破棄する場合は true を指定し、その他の場合は false を指定します。</param>
        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        #region コンポーネント デザイナーで生成されたコード

        /// <summary>
        /// デザイナー サポートに必要なメソッドです。このメソッドの内容を
        /// コード エディターで変更しないでください。
        /// </summary>
        private void InitializeComponent()
        {
            this.tab_Benry = this.Factory.CreateRibbonTab();
            this.group1 = this.Factory.CreateRibbonGroup();
            this.label_UnifyFontsTargetFont = this.Factory.CreateRibbonLabel();
            this.dropDown_UnifyFontsTargetFont = this.Factory.CreateRibbonDropDown();
            this.UnifyFonts = this.Factory.CreateRibbonButton();
            this.separator1 = this.Factory.CreateRibbonSeparator();
            this.tab_Benry.SuspendLayout();
            this.group1.SuspendLayout();
            this.SuspendLayout();
            // 
            // tab_Benry
            // 
            this.tab_Benry.ControlId.ControlIdType = Microsoft.Office.Tools.Ribbon.RibbonControlIdType.Office;
            this.tab_Benry.Groups.Add(this.group1);
            this.tab_Benry.Label = "【Benry】";
            this.tab_Benry.Name = "tab_Benry";
            // 
            // group1
            // 
            this.group1.Items.Add(this.label_UnifyFontsTargetFont);
            this.group1.Items.Add(this.dropDown_UnifyFontsTargetFont);
            this.group1.Items.Add(this.separator1);
            this.group1.Items.Add(this.UnifyFonts);
            this.group1.Label = "全ページのフォント統一";
            this.group1.Name = "group1";
            // 
            // label_UnifyFontsTargetFont
            // 
            this.label_UnifyFontsTargetFont.Label = "すべてのページをこのフォントに統一する：";
            this.label_UnifyFontsTargetFont.Name = "label_UnifyFontsTargetFont";
            // 
            // dropDown_UnifyFontsTargetFont
            // 
            this.dropDown_UnifyFontsTargetFont.Label = "TargetFont";
            this.dropDown_UnifyFontsTargetFont.Name = "dropDown_UnifyFontsTargetFont";
            this.dropDown_UnifyFontsTargetFont.ShowLabel = false;
            this.dropDown_UnifyFontsTargetFont.SizeString = "wwwwwwwwwwwwwwwwwwww";
            this.dropDown_UnifyFontsTargetFont.SelectionChanged += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.dropDown_UnifyFontsTargetFont_SelectionChanged);
            // 
            // UnifyFonts
            // 
            this.UnifyFonts.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.UnifyFonts.Label = "統一開始";
            this.UnifyFonts.Name = "UnifyFonts";
            this.UnifyFonts.OfficeImageId = "FontsReplaceFonts";
            this.UnifyFonts.ShowImage = true;
            this.UnifyFonts.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.UnifyFont_Click);
            // 
            // separator1
            // 
            this.separator1.Name = "separator1";
            // 
            // Ribbon
            // 
            this.Name = "Ribbon";
            this.RibbonType = "Microsoft.PowerPoint.Presentation";
            this.Tabs.Add(this.tab_Benry);
            this.Load += new Microsoft.Office.Tools.Ribbon.RibbonUIEventHandler(this.Ribbon_Load);
            this.tab_Benry.ResumeLayout(false);
            this.tab_Benry.PerformLayout();
            this.group1.ResumeLayout(false);
            this.group1.PerformLayout();
            this.ResumeLayout(false);

        }

        #endregion

        internal Microsoft.Office.Tools.Ribbon.RibbonTab tab_Benry;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup group1;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton UnifyFonts;
        internal Microsoft.Office.Tools.Ribbon.RibbonDropDown dropDown_UnifyFontsTargetFont;
        internal Microsoft.Office.Tools.Ribbon.RibbonLabel label_UnifyFontsTargetFont;
        internal Microsoft.Office.Tools.Ribbon.RibbonSeparator separator1;
    }

    partial class ThisRibbonCollection
    {
        internal Ribbon Ribbon
        {
            get { return this.GetRibbon<Ribbon>(); }
        }
    }
}
