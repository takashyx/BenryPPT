using Microsoft.Office.Interop.PowerPoint;
using Microsoft.Office.Tools.Ribbon;
using System;
using System.Diagnostics;
using System.Drawing;
using System.Drawing.Text;
using System.Linq;
using System.Windows.Forms;

namespace BenryPPT
{
    public partial class Ribbon
    {

        private void Ribbon_Load(object sender, RibbonUIEventArgs e)
        {
            bool targetSettingsMatch = false;
            int targetIndex = 0;
            string targetFont = Settings.Default.targetFontForUnifyFonts.ToString();

            // set fontfamily to dropdown
            InstalledFontCollection fonts = new InstalledFontCollection();
            FontFamily[] ffArray = fonts.Families;


            RibbonDropDownItem item;
            dropDown_UnifyFontsTargetFont.Items.Clear();

            foreach (FontFamily ff in ffArray)
            {
                item = Factory.CreateRibbonDropDownItem();
                item.Label = ff.Name;
                dropDown_UnifyFontsTargetFont.Items.Add(item);

                if (targetFont.Equals(item.Label.ToString()))
                {
                    targetSettingsMatch = true;
                    targetIndex = dropDown_UnifyFontsTargetFont.Items.IndexOf(item);
                }
            }

            // set target font
            if (targetSettingsMatch)
            {
                dropDown_UnifyFontsTargetFont.SelectedItemIndex = targetIndex;
            }
        }

        private void convertShapeFont(Shape shape, string targetFont)
        {
            if (shape.HasTextFrame == Microsoft.Office.Core.MsoTriState.msoTrue)
            {
                shape.TextFrame.TextRange.Font.Name = targetFont;
                shape.TextFrame.TextRange.Font.NameFarEast = targetFont;
                shape.TextFrame2.TextRange.Font.Name = targetFont;
                shape.TextFrame2.TextRange.Font.NameFarEast = targetFont;
            }
        }

        private void UnifyFont_ConvertFonts(FormProgress pr)
        {

            string targetFont = Settings.Default.targetFontForUnifyFonts;

            try
            {
                // convert slide items
                var slides = Globals.ThisAddIn.Application.ActivePresentation.Slides;
                int slide_count_all = slides.Count;
                int slide_count_processed = 0;

                foreach (Slide slide in slides)
                {
                    pr.setProgressBarMessage("作業中: "+(slide_count_processed + 1)+"枚目 / "+slide_count_all+"枚中");
                    pr.setProgressBarPercentage(100 * slide_count_processed / slide_count_all);

                    foreach (Shape shape in slide.Shapes)
                    {
                        if (shape.Type == Microsoft.Office.Core.MsoShapeType.msoGroup || shape.Type == Microsoft.Office.Core.MsoShapeType.msoSmartArt)
                        {
                            foreach (Shape item in shape.GroupItems)
                            {
                                convertShapeFont(item, targetFont);
                            }
                        }

                        // Shapes with texts
                        convertShapeFont(shape, targetFont);

                        // Tables
                        if (shape.HasTable == Microsoft.Office.Core.MsoTriState.msoTrue)
                        {
                            foreach (int i in Enumerable.Range(1, shape.Table.Rows.Count))
                            {
                                foreach (int j in Enumerable.Range(1, shape.Table.Columns.Count))
                                {
                                    Cell cell = shape.Table.Cell(i, j);
                                    convertShapeFont(cell.Shape, targetFont);
                                }
                            }
                        }
                        // Charts
                        if (shape.HasChart == Microsoft.Office.Core.MsoTriState.msoTrue)
                        {
                            if (shape.Chart.HasTitle) shape.Chart.ChartTitle.Font.Name = targetFont;
                            if (shape.Chart.HasLegend) shape.Chart.Legend.Font.Name = targetFont;
                        }
                    }
                    slide_count_processed += 1;

                }
            }
            catch (Exception ex)
            {
                System.Windows.Forms.MessageBox.Show("Benry Error: " + ex);
            }
            Debug.WriteLine("Converting Done.");
        }

        private void UnifyFont_Click(object sender, RibbonControlEventArgs e)
        {
            this.RibbonButton_UnifyFonts.Enabled = false;
            this.dropDown_UnifyFontsTargetFont.Enabled = false;
            // read target font
            string FontFamilyName = Settings.Default.targetFontForUnifyFonts;

            var progress = new FormProgress();

            progress.setFormTitle("フォントを統一しています");
            progress.Show();

            // splash screen
            UnifyFont_ConvertFonts(progress);

            progress.exitForm();

            this.dropDown_UnifyFontsTargetFont.Enabled = true;
            this.RibbonButton_UnifyFonts.Enabled = true;
        }

        private void dropDown_UnifyFontsTargetFont_SelectionChanged(object sender, RibbonControlEventArgs e)
        {
            Settings.Default.targetFontForUnifyFonts = dropDown_UnifyFontsTargetFont.SelectedItem.ToString();
            Settings.Default.Save();
        }
    }
}
