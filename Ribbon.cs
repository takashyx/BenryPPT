using Microsoft.Office.Interop.PowerPoint;
using Microsoft.Office.Tools.Ribbon;
using Microsoft.VisualBasic;
using System;
using System.Diagnostics;
using System.Drawing;
using System.Drawing.Text;
using System.Linq;
using System.Text.RegularExpressions;
using System.Windows.Forms;

namespace BenryPPT
{
    public partial class Ribbon
    {

        private void Ribbon_Load(object sender, RibbonUIEventArgs e)
        {
            bool targetFontMatch = false;
            bool targetFontFarEastMatch = false;
            int targetFontIndex = 0;
            int targetFontFarEastIndex = 0;
            string targetFont = Settings.Default.targetFontForUnifyFonts.ToString();
            string targetFontFarEast = Settings.Default.targetFontFarEastForUnifyFonts.ToString();

            // set fontfamily to dropdown
            InstalledFontCollection fonts = new InstalledFontCollection();
            FontFamily[] ffArray = fonts.Families;


            RibbonDropDownItem item;
            dropDown_UnifyFontsTargetFont.Items.Clear();
            dropDown_UnifyFontsTargetFontFarEast.Items.Clear();

            foreach (FontFamily ff in ffArray)
            {
                item = Factory.CreateRibbonDropDownItem();
                item.Label = ff.Name;
                dropDown_UnifyFontsTargetFont.Items.Add(item);

                item = Factory.CreateRibbonDropDownItem();
                item.Label = ff.Name;
                dropDown_UnifyFontsTargetFontFarEast.Items.Add(item);

                if (targetFont.Equals(item.Label.ToString()))
                {
                    targetFontMatch = true;
                    targetFontIndex = dropDown_UnifyFontsTargetFont.Items.IndexOf(item);
                }
                if (targetFontFarEast.Equals(item.Label.ToString()))
                {
                    targetFontFarEastMatch = true;
                    targetFontFarEastIndex = dropDown_UnifyFontsTargetFontFarEast.Items.IndexOf(item);
                }
            }

            // set target font
            if (targetFontMatch) dropDown_UnifyFontsTargetFont.SelectedItemIndex = targetFontIndex;
            if (targetFontFarEastMatch) dropDown_UnifyFontsTargetFontFarEast.SelectedItemIndex = targetFontFarEastIndex;
        }

        private void convertShapeFont(Shape shape, string targetFont, string targetFontFarEast)
        {
            if (shape.HasTextFrame == Microsoft.Office.Core.MsoTriState.msoTrue)
            {
                shape.TextFrame.TextRange.Font.Name = targetFont;
                shape.TextFrame.TextRange.Font.NameFarEast = targetFontFarEast;
                shape.TextFrame2.TextRange.Font.Name = targetFont;
                shape.TextFrame2.TextRange.Font.NameFarEast = targetFontFarEast;
            }
        }

        private void UnifyFont_ConvertFonts(FormProgress pr)
        {

            string targetFont = Settings.Default.targetFontForUnifyFonts;
            string targetFontFarEast = Settings.Default.targetFontFarEastForUnifyFonts;

            try
            {
                // convert slide items
                var slides = Globals.ThisAddIn.Application.ActivePresentation.Slides;
                int slide_count_all = slides.Count;
                int slide_count_processed = 0;

                foreach (Slide slide in slides)
                {
                    pr.setProgressBarMessage("作業中: "+(slide_count_processed + 1)+"枚目 / "+slide_count_all+"枚中");
                    pr.setProgressBarPercentage((100 * slide_count_processed) / slide_count_all);

                    foreach (Shape shape in slide.Shapes)
                    {
                        // Grouped Shape and Smart Art
                        if (shape.Type == Microsoft.Office.Core.MsoShapeType.msoGroup || shape.Type == Microsoft.Office.Core.MsoShapeType.msoSmartArt)
                        {
                            foreach (Shape item in shape.GroupItems)
                            {
                                convertShapeFont(item, targetFont, targetFontFarEast);
                            }
                        }

                        // Shapes with texts
                        convertShapeFont(shape, targetFont, targetFontFarEast);

                        // Tables
                        if (shape.HasTable == Microsoft.Office.Core.MsoTriState.msoTrue)
                        {
                            foreach (int i in Enumerable.Range(1, shape.Table.Rows.Count))
                            {
                                foreach (int j in Enumerable.Range(1, shape.Table.Columns.Count))
                                {
                                    Cell cell = shape.Table.Cell(i, j);
                                    convertShapeFont(cell.Shape, targetFont, targetFontFarEast);
                                }
                            }
                        }
                        // Charts
                        if (shape.HasChart == Microsoft.Office.Core.MsoTriState.msoTrue)
                        {
                            // workaround
                            if (shape.Chart.HasTitle) shape.Chart.ChartTitle.Font.Name = targetFontFarEast;
                            if (shape.Chart.HasTitle) shape.Chart.ChartTitle.Font.Name = targetFont;
                            if (shape.Chart.HasLegend) shape.Chart.Legend.Font.Name = targetFontFarEast;
                            if (shape.Chart.HasLegend) shape.Chart.Legend.Font.Name = targetFont;
                        }
                    }
                    slide_count_processed += 1;

                }
            }
            catch (Exception ex)
            {
                System.Windows.Forms.MessageBox.Show("Benry Error: \n" + ex);
            }
            Debug.WriteLine("Converting Done.");
        }

        private void UnifyFont_Click(object sender, RibbonControlEventArgs e)
        {
            // Disable controls
            this.button_UnifyFonts.Enabled = false;
            this.dropDown_UnifyFontsTargetFont.Enabled = false;
            this.dropDown_UnifyFontsTargetFontFarEast.Enabled = false;

            // show progress bar and convert
            var progress = new FormProgress();

            progress.setFormTitle("フォントを統一しています");
            progress.Show();

            UnifyFont_ConvertFonts(progress);

            progress.exitForm();

            // Enable controls
            this.dropDown_UnifyFontsTargetFont.Enabled = true;
            this.dropDown_UnifyFontsTargetFontFarEast.Enabled = true;
            this.button_UnifyFonts.Enabled = true;
        }

        private void dropDown_UnifyFontsTargetFont_SelectionChanged(object sender, RibbonControlEventArgs e)
        {
            Settings.Default.targetFontForUnifyFonts = dropDown_UnifyFontsTargetFont.SelectedItem.ToString();
            Settings.Default.Save();
        }

        private void dropDown_UnifyFontsTargetFontFarEast_SelectionChanged(object sender, RibbonControlEventArgs e)
        {
            Settings.Default.targetFontFarEastForUnifyFonts = dropDown_UnifyFontsTargetFontFarEast.SelectedItem.ToString();
            Settings.Default.Save();
        }

        private static string abc123ToHankaku(string s)
        {
            Regex re = new Regex("[０-９Ａ-Ｚａ-ｚ：－　]+");
            string output = re.Replace(s, myReplacer);

            return output;
        }

        private static string myReplacer(Match m)
        {
            return Strings.StrConv(m.Value, VbStrConv.Narrow, 0);
        }

        private void convert_shape_zenkakuToHankaku(Shape shape)
        {
            if (shape.HasTextFrame == Microsoft.Office.Core.MsoTriState.msoTrue) shape.TextFrame.TextRange.Text = abc123ToHankaku(shape.TextFrame.TextRange.Text); 
        }

        private void convertZenkakuToHankaku(FormProgress pr)
        {
            try
            {
                // convert slide items
                var slides = Globals.ThisAddIn.Application.ActivePresentation.Slides;
                int slide_count_all = slides.Count;
                int slide_count_processed = 0;

                foreach (Slide slide in slides)
                {
                    pr.setProgressBarMessage("作業中: "+(slide_count_processed + 1)+"枚目 / "+slide_count_all+"枚中");
                    pr.setProgressBarPercentage((100 * slide_count_processed) / slide_count_all);

                    foreach (Shape shape in slide.Shapes)
                    {
                        // Grouped Shape and Smart Art
                        if (shape.Type == Microsoft.Office.Core.MsoShapeType.msoGroup || shape.Type == Microsoft.Office.Core.MsoShapeType.msoSmartArt)
                        {
                            foreach (Shape item in shape.GroupItems)
                            {
                                convert_shape_zenkakuToHankaku(item);
                            }
                        }

                        // Tables
                        if (shape.HasTable == Microsoft.Office.Core.MsoTriState.msoTrue)
                        {
                            foreach (int i in Enumerable.Range(1, shape.Table.Rows.Count))
                            {
                                foreach (int j in Enumerable.Range(1, shape.Table.Columns.Count))
                                {
                                    Cell cell = shape.Table.Cell(i, j);
                                    convert_shape_zenkakuToHankaku(cell.Shape);
                                }
                            }
                        }

                        // Charts
                        if (shape.HasChart == Microsoft.Office.Core.MsoTriState.msoTrue)
                        {
                            if (shape.Chart.HasTitle) shape.Chart.ChartTitle.Text = abc123ToHankaku(shape.Chart.ChartTitle.Text);
                        }

                        // Shapes with texts
                        else
                        { 
                            convert_shape_zenkakuToHankaku(shape);
                        }

                    }
                    slide_count_processed += 1;

                }
            }
            catch (Exception ex)
            {
                System.Windows.Forms.MessageBox.Show("Benry Error: \n" + ex);
            }
            Debug.WriteLine("Converting Done.");
        }

        private void button_zenkakuToHankaku_Click(object sender, RibbonControlEventArgs e)
        {
            // Disable controls
            this.button_zenkakuToHankaku.Enabled = false;

            // show progress bar and convert
            var progress = new FormProgress();

            progress.setFormTitle("全角の英字・数字を半角に変換中");
            progress.Show();

            convertZenkakuToHankaku(progress);

            progress.exitForm();

            // Enable controls
            this.button_zenkakuToHankaku.Enabled = true;
        }
    }
}
