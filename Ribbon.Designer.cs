﻿namespace BenryPPT
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
            this.group_hankaku = this.Factory.CreateRibbonGroup();
            this.group_multiple = this.Factory.CreateRibbonGroup();
            this.checkBox_unifyFonts = this.Factory.CreateRibbonCheckBox();
            this.checkBox_zenkakuToHankaku = this.Factory.CreateRibbonCheckBox();
            this.group2 = this.Factory.CreateRibbonGroup();
            this.button_UnifyFonts = this.Factory.CreateRibbonButton();
            this.button_zenkakuToHankaku = this.Factory.CreateRibbonButton();
            this.button_multiple = this.Factory.CreateRibbonButton();
            this.button_align_same_height_horizontal = this.Factory.CreateRibbonButton();
            this.button_align_same_width_vertical = this.Factory.CreateRibbonButton();
            this.checkBox_widthAlso = this.Factory.CreateRibbonCheckBox();
            this.checkBox_heightAlso = this.Factory.CreateRibbonCheckBox();
            this.separator3 = this.Factory.CreateRibbonSeparator();
            this.group_info = this.Factory.CreateRibbonGroup();
            this.label_ProductVersion = this.Factory.CreateRibbonLabel();
            this.label_assemblyFileversion = this.Factory.CreateRibbonLabel();
            this.label_versionTitle = this.Factory.CreateRibbonLabel();
            this.tab_Benry.SuspendLayout();
            this.group1.SuspendLayout();
            this.group_hankaku.SuspendLayout();
            this.group_multiple.SuspendLayout();
            this.group2.SuspendLayout();
            this.group_info.SuspendLayout();
            this.SuspendLayout();
            // 
            // tab_Benry
            // 
            this.tab_Benry.ControlId.ControlIdType = Microsoft.Office.Tools.Ribbon.RibbonControlIdType.Office;
            this.tab_Benry.Groups.Add(this.group1);
            this.tab_Benry.Groups.Add(this.group_hankaku);
            this.tab_Benry.Groups.Add(this.group_multiple);
            this.tab_Benry.Groups.Add(this.group2);
            this.tab_Benry.Groups.Add(this.group_info);
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
            this.dropDown_UnifyFontsTargetFont.SizeString = "wwwwwwwwwwww";
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
            this.dropDown_UnifyFontsTargetFontFarEast.SizeString = "wwwwwwwwwwww";
            this.dropDown_UnifyFontsTargetFontFarEast.SelectionChanged += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.dropDown_UnifyFontsTargetFontFarEast_SelectionChanged);
            // 
            // separator1
            // 
            this.separator1.Name = "separator1";
            // 
            // group_hankaku
            // 
            this.group_hankaku.Items.Add(this.button_zenkakuToHankaku);
            this.group_hankaku.Label = "全ページの全角英数字を半角化";
            this.group_hankaku.Name = "group_hankaku";
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
            // group2
            // 
            this.group2.Items.Add(this.checkBox_widthAlso);
            this.group2.Items.Add(this.button_align_same_height_horizontal);
            this.group2.Items.Add(this.separator3);
            this.group2.Items.Add(this.checkBox_heightAlso);
            this.group2.Items.Add(this.button_align_same_width_vertical);
            this.group2.Label = "左上の選択図形とサイズを合わせて均等に再配置";
            this.group2.Name = "group2";
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
            // button_zenkakuToHankaku
            // 
            this.button_zenkakuToHankaku.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.button_zenkakuToHankaku.Label = "半角化開始";
            this.button_zenkakuToHankaku.Name = "button_zenkakuToHankaku";
            this.button_zenkakuToHankaku.OfficeImageId = "AsianLayoutMenu";
            this.button_zenkakuToHankaku.ShowImage = true;
            this.button_zenkakuToHankaku.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.button_zenkakuToHankaku_Click);
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
            // button_align_same_height_horizontal
            // 
            this.button_align_same_height_horizontal.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.button_align_same_height_horizontal.Label = "高さを合わせて横に並べる";
            this.button_align_same_height_horizontal.Name = "button_align_same_height_horizontal";
            this.button_align_same_height_horizontal.OfficeImageId = "PrintPreviewZoomTwoPages";
            this.button_align_same_height_horizontal.ShowImage = true;
            this.button_align_same_height_horizontal.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.button_align_same_height_horizontal_Click);
            // 
            // button_align_same_width_vertical
            // 
            this.button_align_same_width_vertical.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.button_align_same_width_vertical.Label = "幅を合わせて縦に並べる";
            this.button_align_same_width_vertical.Name = "button_align_same_width_vertical";
            this.button_align_same_width_vertical.OfficeImageId = "WindowsTileHorizontally";
            this.button_align_same_width_vertical.ShowImage = true;
            this.button_align_same_width_vertical.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.button_align_same_width_vertical_Click);
            // 
            // checkBox_widthAlso
            // 
            this.checkBox_widthAlso.Label = "幅も合わせる";
            this.checkBox_widthAlso.Name = "checkBox_widthAlso";
            this.checkBox_widthAlso.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.checkBox_widthAlso_Click);
            // 
            // checkBox_heightAlso
            // 
            this.checkBox_heightAlso.Label = "高さも合わせる";
            this.checkBox_heightAlso.Name = "checkBox_heightAlso";
            this.checkBox_heightAlso.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.checkBox_heightAlso_Click);
            // 
            // separator3
            // 
            this.separator3.Name = "separator3";
            // 
            // group_info
            // 
            this.group_info.Items.Add(this.label_versionTitle);
            this.group_info.Items.Add(this.label_ProductVersion);
            this.group_info.Items.Add(this.label_assemblyFileversion);
            this.group_info.Label = "バージョン情報";
            this.group_info.Name = "group_info";
            // 
            // label_ProductVersion
            // 
            this.label_ProductVersion.Label = "product version";
            this.label_ProductVersion.Name = "label_ProductVersion";
            // 
            // label_assemblyFileversion
            // 
            this.label_assemblyFileversion.Label = "label_assemblyFileVersion";
            this.label_assemblyFileversion.Name = "label_assemblyFileversion";
            // 
            // label_versionTitle
            // 
            this.label_versionTitle.Label = "BenryPPT";
            this.label_versionTitle.Name = "label_versionTitle";
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
            this.group2.ResumeLayout(false);
            this.group2.PerformLayout();
            this.group_info.ResumeLayout(false);
            this.group_info.PerformLayout();
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
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup group2;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton button_align_same_height_horizontal;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton button_align_same_width_vertical;
        internal Microsoft.Office.Tools.Ribbon.RibbonCheckBox checkBox_widthAlso;
        internal Microsoft.Office.Tools.Ribbon.RibbonCheckBox checkBox_heightAlso;
        internal Microsoft.Office.Tools.Ribbon.RibbonSeparator separator3;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup group_info;
        internal Microsoft.Office.Tools.Ribbon.RibbonLabel label_ProductVersion;
        internal Microsoft.Office.Tools.Ribbon.RibbonLabel label_assemblyFileversion;
        internal Microsoft.Office.Tools.Ribbon.RibbonLabel label_versionTitle;
    }

    partial class ThisRibbonCollection
    {
        internal Ribbon Ribbon
        {
            get { return this.GetRibbon<Ribbon>(); }
        }
    }
}
