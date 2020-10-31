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
            this.label_font = this.Factory.CreateRibbonLabel();
            this.dropDown_UnifyFontsTargetFont = this.Factory.CreateRibbonDropDown();
            this.separator2 = this.Factory.CreateRibbonSeparator();
            this.label_fontFarEast = this.Factory.CreateRibbonLabel();
            this.dropDown_UnifyFontsTargetFontFarEast = this.Factory.CreateRibbonDropDown();
            this.separator1 = this.Factory.CreateRibbonSeparator();
            this.button_UnifyFonts = this.Factory.CreateRibbonButton();
            this.group_hankaku = this.Factory.CreateRibbonGroup();
            this.button_zenkakuToHankaku = this.Factory.CreateRibbonButton();
            this.group_multiple = this.Factory.CreateRibbonGroup();
            this.checkBox_unifyFonts = this.Factory.CreateRibbonCheckBox();
            this.checkBox_zenkakuToHankaku = this.Factory.CreateRibbonCheckBox();
            this.button_multiple = this.Factory.CreateRibbonButton();
            this.tab_Benry.SuspendLayout();
            this.group1.SuspendLayout();
            this.group_hankaku.SuspendLayout();
            this.group_multiple.SuspendLayout();
            this.SuspendLayout();
            // 
            // tab_Benry
            // 
            this.tab_Benry.ControlId.ControlIdType = Microsoft.Office.Tools.Ribbon.RibbonControlIdType.Office;
            this.tab_Benry.Groups.Add(this.group1);
            this.tab_Benry.Groups.Add(this.group_hankaku);
            this.tab_Benry.Groups.Add(this.group_multiple);
            this.tab_Benry.Label = "【Benry】";
            this.tab_Benry.Name = "tab_Benry";
            // 
            // group1
            // 
            this.group1.Items.Add(this.label_font);
            this.group1.Items.Add(this.dropDown_UnifyFontsTargetFont);
            this.group1.Items.Add(this.separator2);
            this.group1.Items.Add(this.label_fontFarEast);
            this.group1.Items.Add(this.dropDown_UnifyFontsTargetFontFarEast);
            this.group1.Items.Add(this.separator1);
            this.group1.Items.Add(this.button_UnifyFonts);
            this.group1.Label = "全ページのフォント統一";
            this.group1.Name = "group1";
            // 
            // label_font
            // 
            this.label_font.Label = "英文";
            this.label_font.Name = "label_font";
            // 
            // dropDown_UnifyFontsTargetFont
            // 
            this.dropDown_UnifyFontsTargetFont.Label = "TargetFont";
            this.dropDown_UnifyFontsTargetFont.Name = "dropDown_UnifyFontsTargetFont";
            this.dropDown_UnifyFontsTargetFont.ShowLabel = false;
            this.dropDown_UnifyFontsTargetFont.SizeString = "wwwwwwwwwwwwwww";
            this.dropDown_UnifyFontsTargetFont.SelectionChanged += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.dropDown_UnifyFontsTargetFont_SelectionChanged);
            // 
            // separator2
            // 
            this.separator2.Name = "separator2";
            // 
            // label_fontFarEast
            // 
            this.label_fontFarEast.Label = "和文";
            this.label_fontFarEast.Name = "label_fontFarEast";
            // 
            // dropDown_UnifyFontsTargetFontFarEast
            // 
            this.dropDown_UnifyFontsTargetFontFarEast.Label = "TargetFontFarEast";
            this.dropDown_UnifyFontsTargetFontFarEast.Name = "dropDown_UnifyFontsTargetFontFarEast";
            this.dropDown_UnifyFontsTargetFontFarEast.ShowLabel = false;
            this.dropDown_UnifyFontsTargetFontFarEast.SizeString = "wwwwwwwwwwwwwww";
            this.dropDown_UnifyFontsTargetFontFarEast.SelectionChanged += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.dropDown_UnifyFontsTargetFontFarEast_SelectionChanged);
            // 
            // separator1
            // 
            this.separator1.Name = "separator1";
            // 
            // button_UnifyFonts
            // 
            this.button_UnifyFonts.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.button_UnifyFonts.Label = "統一開始";
            this.button_UnifyFonts.Name = "button_UnifyFonts";
            this.button_UnifyFonts.OfficeImageId = "FontsReplaceFonts";
            this.button_UnifyFonts.ShowImage = true;
            this.button_UnifyFonts.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.UnifyFont_Click);
            // 
            // group_hankaku
            // 
            this.group_hankaku.Items.Add(this.button_zenkakuToHankaku);
            this.group_hankaku.Label = "全ページの全角英数字を半角化";
            this.group_hankaku.Name = "group_hankaku";
            // 
            // button_zenkakuToHankaku
            // 
            this.button_zenkakuToHankaku.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.button_zenkakuToHankaku.Label = "半角化開始";
            this.button_zenkakuToHankaku.Name = "button_zenkakuToHankaku";
            this.button_zenkakuToHankaku.OfficeImageId = "AsianLayoutMenu";
            this.button_zenkakuToHankaku.ShowImage = true;
            this.button_zenkakuToHankaku.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.button_zenkakuToHankaku_Click);
            // 
            // group_multiple
            // 
            this.group_multiple.Items.Add(this.checkBox_unifyFonts);
            this.group_multiple.Items.Add(this.checkBox_zenkakuToHankaku);
            this.group_multiple.Items.Add(this.button_multiple);
            this.group_multiple.Label = "全ページに複数の処理をまとめて実行";
            this.group_multiple.Name = "group_multiple";
            // 
            // checkBox_unifyFonts
            // 
            this.checkBox_unifyFonts.Label = "フォント統一";
            this.checkBox_unifyFonts.Name = "checkBox_unifyFonts";
            this.checkBox_unifyFonts.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.checkBox_unifyFonts_Click);
            // 
            // checkBox_zenkakuToHankaku
            // 
            this.checkBox_zenkakuToHankaku.Label = "全角英数字を半角化";
            this.checkBox_zenkakuToHankaku.Name = "checkBox_zenkakuToHankaku";
            this.checkBox_zenkakuToHankaku.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.checkBox_zenkakuToHankaku_Click);
            // 
            // button_multiple
            // 
            this.button_multiple.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.button_multiple.Label = "まとめて開始";
            this.button_multiple.Name = "button_multiple";
            this.button_multiple.OfficeImageId = "WorkTrackingForm";
            this.button_multiple.ShowImage = true;
            this.button_multiple.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.button_multiple_Click);
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
            this.group_hankaku.ResumeLayout(false);
            this.group_hankaku.PerformLayout();
            this.group_multiple.ResumeLayout(false);
            this.group_multiple.PerformLayout();
            this.ResumeLayout(false);

        }

        #endregion

        internal Microsoft.Office.Tools.Ribbon.RibbonTab tab_Benry;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup group1;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton button_UnifyFonts;
        internal Microsoft.Office.Tools.Ribbon.RibbonDropDown dropDown_UnifyFontsTargetFont;
        internal Microsoft.Office.Tools.Ribbon.RibbonSeparator separator1;
        internal Microsoft.Office.Tools.Ribbon.RibbonLabel label_font;
        internal Microsoft.Office.Tools.Ribbon.RibbonLabel label_fontFarEast;
        internal Microsoft.Office.Tools.Ribbon.RibbonDropDown dropDown_UnifyFontsTargetFontFarEast;
        internal Microsoft.Office.Tools.Ribbon.RibbonSeparator separator2;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup group_hankaku;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton button_zenkakuToHankaku;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup group_multiple;
        internal Microsoft.Office.Tools.Ribbon.RibbonCheckBox checkBox_unifyFonts;
        internal Microsoft.Office.Tools.Ribbon.RibbonCheckBox checkBox_zenkakuToHankaku;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton button_multiple;
    }

    partial class ThisRibbonCollection
    {
        internal Ribbon Ribbon
        {
            get { return this.GetRibbon<Ribbon>(); }
        }
    }
}
