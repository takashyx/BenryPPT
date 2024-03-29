﻿using Microsoft.Office.Tools.Ribbon;
using Microsoft.VisualBasic;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Drawing;
using System.Drawing.Text;
using System.Linq;
using System.Text.RegularExpressions;
using PowerPoint = Microsoft.Office.Interop.PowerPoint;
using Office = Microsoft.Office;
using System.Reflection;

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

            // set multiple do checkboxes
            this.checkBox_unifyFonts.Checked = Settings.Default.multipleDoFontUnify;
            this.checkBox_zenkakuToHankaku.Checked = Settings.Default.multipleDoZenkakuToHankaku;

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

                if (targetFont.Equals(item.Label.ToString()))
                {
                    targetFontMatch = true;
                    targetFontIndex = dropDown_UnifyFontsTargetFont.Items.IndexOf(item);
                }

                item = Factory.CreateRibbonDropDownItem();
                item.Label = ff.Name;
                dropDown_UnifyFontsTargetFontFarEast.Items.Add(item);

                if (targetFontFarEast.Equals(item.Label.ToString()))
                {
                    targetFontFarEastMatch = true;
                    targetFontFarEastIndex = dropDown_UnifyFontsTargetFontFarEast.Items.IndexOf(item);
                }
            }

            // set target font
            if (targetFontMatch) dropDown_UnifyFontsTargetFont.SelectedItemIndex = targetFontIndex;
            if (targetFontFarEastMatch) dropDown_UnifyFontsTargetFontFarEast.SelectedItemIndex = targetFontFarEastIndex;
            
            // show version info
            FileVersionInfo ver = FileVersionInfo.GetVersionInfo(Assembly.GetExecutingAssembly().Location);
            this.label_ProductVersion.Label =       "ver:   "+ver.ProductVersion;
            this.label_assemblyFileversion.Label =  "build: "+ ver.FileVersion;
        }

        private void convertShapeFont(PowerPoint.Shape shape, string targetFont, string targetFontFarEast)
        {
            if (shape.HasTextFrame == Office.Core.MsoTriState.msoTrue)
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

                foreach (PowerPoint.Slide slide in slides)
                {
                    pr.setProgressBarMessage("作業中: " + (slide_count_processed + 1) + "枚目 / " + slide_count_all + "枚中");
                    pr.setProgressBarPercentage((100 * slide_count_processed) / slide_count_all);

                    foreach (PowerPoint.Shape shape in slide.Shapes)
                    {
                        // Grouped Shape and Smart Art
                        if (shape.Type == Microsoft.Office.Core.MsoShapeType.msoGroup || shape.Type == Microsoft.Office.Core.MsoShapeType.msoSmartArt)
                        {
                            foreach (PowerPoint.Shape item in shape.GroupItems)
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
                                    PowerPoint.Cell cell = shape.Table.Cell(i, j);
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
            Regex re = new Regex("[０-９Ａ-Ｚａ-ｚ：；－＿　]+");
            return re.Replace(s, myReplacer);
        }

        private static string myReplacer(Match m)
        {
            return Strings.StrConv(m.Value, VbStrConv.Narrow, 0);
        }

        private void convert_shape_zenkakuToHankaku(PowerPoint.Shape shape)
        {
            if (shape.HasTextFrame == Microsoft.Office.Core.MsoTriState.msoTrue)
            {
                int count = shape.TextFrame.TextRange.Runs(-1,-1).Count;
                for (int i = count; i > 0; i--)
                {
                    PowerPoint.TextRange run = shape.TextFrame.TextRange.Runs(i);
                    /*
                    Debug.WriteLine(run.Text);
                    Debug.WriteLine("-");
                    Debug.WriteLine(abc123ToHankaku(run.Text));
                    Debug.WriteLine("--------------------");
                    */
                    run.Text = abc123ToHankaku(run.Text);
                }
            }
        }

        private void convert_shape_bufont(PowerPoint.Shape shape)
        {
            if (shape.HasTextFrame == Microsoft.Office.Core.MsoTriState.msoTrue) shape.TextFrame.TextRange.ParagraphFormat.Bullet.Font.Name = "Arial";
        }

        private void convertZenkakuToHankaku(FormProgress pr)
        {
            try
            {
                // convert slide items
                var slides = Globals.ThisAddIn.Application.ActivePresentation.Slides;
                int slide_count_all = slides.Count;
                int slide_count_processed = 0;

                foreach (PowerPoint.Slide slide in slides)
                {
                    pr.setProgressBarMessage("作業中: " + (slide_count_processed + 1) + "枚目 / " + slide_count_all + "枚中");
                    pr.setProgressBarPercentage((100 * slide_count_processed) / slide_count_all);

                    foreach (PowerPoint.Shape shape in slide.Shapes)
                    {
                        // Grouped Shape and Smart Art
                        if (shape.Type == Microsoft.Office.Core.MsoShapeType.msoGroup || shape.Type == Microsoft.Office.Core.MsoShapeType.msoSmartArt)
                        {
                            foreach (PowerPoint.Shape item in shape.GroupItems)
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
                                    PowerPoint.Cell cell = shape.Table.Cell(i, j);
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

        private void convert_bufont(FormProgress pr)
        {
            try
            {
                // For  <a:buFont typeface="Noto Sans Symbols"/>
                // convert slide items
                var slides = Globals.ThisAddIn.Application.ActivePresentation.Slides;
                int slide_count_all = slides.Count;
                int slide_count_processed = 0;

                foreach (PowerPoint.Slide slide in slides)
                {
                    pr.setProgressBarMessage("作業中: " + (slide_count_processed + 1) + "枚目 / " + slide_count_all + "枚中");
                    pr.setProgressBarPercentage((100 * slide_count_processed) / slide_count_all);

                    foreach (PowerPoint.Shape shape in slide.Shapes)
                    {
                        // Grouped Shape and Smart Art
                        if (shape.Type == Microsoft.Office.Core.MsoShapeType.msoGroup || shape.Type == Microsoft.Office.Core.MsoShapeType.msoSmartArt)
                        {
                            foreach (PowerPoint.Shape item in shape.GroupItems)
                            {
                                convert_shape_bufont(item);
                            }
                        }

                        // Tables
                        if (shape.HasTable == Microsoft.Office.Core.MsoTriState.msoTrue)
                        {
                            foreach (int i in Enumerable.Range(1, shape.Table.Rows.Count))
                            {
                                foreach (int j in Enumerable.Range(1, shape.Table.Columns.Count))
                                {
                                    PowerPoint.Cell cell = shape.Table.Cell(i, j);
                                    convert_shape_bufont(cell.Shape);
                                }
                            }
                        }

                        // Shapes with texts
                        else
                        {
                            convert_shape_bufont(shape);
                        }

                    }
                    slide_count_processed += 1;

                }
            }
            catch (Exception ex)
            {
                System.Windows.Forms.MessageBox.Show("Benry Error: \n" + ex);
            }
            Debug.Write("Notosans killer done.");
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

        private void checkBox_unifyFonts_Click(object sender, RibbonControlEventArgs e)
        {
            Settings.Default.multipleDoFontUnify = this.checkBox_unifyFonts.Checked;
            Settings.Default.Save();
        }

        private void checkBox_zenkakuToHankaku_Click(object sender, RibbonControlEventArgs e)
        {
            Settings.Default.multipleDoZenkakuToHankaku = this.checkBox_zenkakuToHankaku.Checked;
            Settings.Default.Save();
        }

        private void button_multiple_Click(object sender, RibbonControlEventArgs e)
        {
            if (this.checkBox_unifyFonts.Checked) UnifyFont_Click(sender, e);
            if (this.checkBox_zenkakuToHankaku.Checked) button_zenkakuToHankaku_Click(sender, e);
        }

        private PowerPoint.Selection GetSelection()
        {
            try
            {
                return Globals.ThisAddIn.Application.ActiveWindow.Selection;
            }
            catch (System.Runtime.InteropServices.COMException exc)
            {
                // TODO
            }
            return null;
        }

        private PowerPoint.ShapeRange GetSelectedShapeRange()
        {
            PowerPoint.Selection selection = GetSelection();
            try
            {
                return (selection != null) ? selection.ShapeRange : null;
            }
            catch (System.Runtime.InteropServices.COMException exc)
            {
                // TODO
            }
            return null;
        }

        private static int CompareTopLeft(PowerPoint.Shape a, PowerPoint.Shape b)
        {
            // Shapeの上端位置で比較
            if (a.Top > b.Top)
            {
                return 1;
            }
            else if (a.Top < b.Top)
            {
                return -1;
            }
            else
            {
                // Shapeの上端位置が同じ場合は、左端位置で比較
                if (a.Left > b.Left)
                    return 1;
                else if (a.Left < b.Left)
                    return -1;
                else
                    return 0;
            }
        }

        private static int CompareLeftTop(PowerPoint.Shape a, PowerPoint.Shape b)
        {
            // Shapeの左端位置で比較
            if (a.Left > b.Left)
            {
                return 1;
            }
            else if (a.Left < b.Left)
            {
                return -1;
            }
            else
            {
                // Shapeの左端位置が同じ場合は、上端位置で比較
                if (a.Top > b.Top)
                    return 1;
                else if (a.Top < b.Top)
                    return -1;
                else
                    return 0;
            }
        }

        private static bool isSelectedShapePortrait(List<PowerPoint.Shape> ss)
        {
            ss.Sort(CompareTopLeft);
            float selectedHeight = ss.Last().Top + ss.Last().Height - ss.First().Top;
            ss.Sort(CompareLeftTop);
            float selectedWidth = ss.Last().Left + ss.Last().Width - ss.First().Left;

            if (selectedHeight >= selectedWidth) return true;
            else return false;
        }


        private void button_relocate_horizontal_Click(object sender, RibbonControlEventArgs e)
        {
            var ss = new List<PowerPoint.Shape>(); ;

            var sr = GetSelectedShapeRange();
            if (sr == null) return;

            int c = sr.Count;

            if (c > 1)
            {
                foreach (PowerPoint.Shape s in sr) ss.Add(s);
                ss.Sort(CompareLeftTop);
                float horizontalCenter = ss.First().Top + (ss.First().Height / 2);

                float selectedWidth = (ss.Last().Left + ss.Last().Width) - ss.First().Left;

                float SumOfShapeWidth = 0;
                for (int i = 0; i < c; i++) SumOfShapeWidth += ss[i].Width;

                float margin = (selectedWidth - SumOfShapeWidth) / (c - 1);
                float currentRight = ss[0].Left + ss[0].Width;
                for (int i = 1; i < c; i++)
                {
                    ss[i].Top = horizontalCenter - (ss[i].Height/2);
                    ss[i].Left = currentRight + margin;
                    currentRight = ss[i].Left + ss[i].Width;
                }
            }
        }

        private void button_relocate_vertical_Click(object sender, RibbonControlEventArgs e)
        {
            var ss = new List<PowerPoint.Shape>(); ;

            var sr = GetSelectedShapeRange();
            if (sr == null) return;

            int c = sr.Count;

            if (c > 1)
            {
                foreach (PowerPoint.Shape s in sr) ss.Add(s);
                ss.Sort(CompareTopLeft);
                float verticalCenter = ss.First().Left + (ss.First().Width / 2);

                float selectedHeight = (ss.Last().Top + ss.Last().Height) - ss.First().Top;

                float SumOfShapeHeight = 0;
                for (int i = 0; i < c; i++) SumOfShapeHeight += ss[i].Height;

                float margin = (selectedHeight - SumOfShapeHeight) / (c - 1);
                float currentBottom = ss[0].Top + ss[0].Height;
                for (int i = 1; i < c; i++)
                {
                    ss[i].Left = verticalCenter - (ss[i].Width/2);
                    ss[i].Top = currentBottom + margin;
                    currentBottom = ss[i].Top + ss[i].Height;
                }
            }
        }

        private void button_resize_width_Click(object sender, RibbonControlEventArgs e)
        {
            var ss = new List<PowerPoint.Shape>(); ;

            var sr = GetSelectedShapeRange();
            if (sr == null) return;

            int c = sr.Count;


            if (c > 1)
            {
                foreach (PowerPoint.Shape s in sr) ss.Add(s);
                if (isSelectedShapePortrait(ss)) ss.Sort(CompareTopLeft);
                else ss.Sort(CompareLeftTop);

                for (int i = 1; i < c; i++)
                {
                    float widthDiff = (ss.First().Width - ss[i].Width);
                    ss[i].Left -= (widthDiff / 2);
                    ss[i].Width = ss.First().Width;
                }
            }
        }

        private void button_resize_height_Click(object sender, RibbonControlEventArgs e)
        {
            var ss = new List<PowerPoint.Shape>();

            var sr = GetSelectedShapeRange();
            if (sr == null) return;

            int c = sr.Count;

            if (c > 1)
            {
                foreach (PowerPoint.Shape s in sr) ss.Add(s);
                if (isSelectedShapePortrait(ss)) ss.Sort(CompareTopLeft);
                else ss.Sort(CompareLeftTop);

                for (int i = 1; i < c; i++)
                {
                    float heightDiff = (ss.First().Height - ss[i].Height);
                    ss[i].Top -= (heightDiff / 2);
                    ss[i].Height = ss.First().Height;
                }
            }


        }

        private void button_kill_bufont_issue_Click(object sender, RibbonControlEventArgs e)
        {
            this.button_kill_font_issue.Enabled = false;
            // show progress bar and convert
            var progress = new FormProgress();

            progress.setFormTitle("除霊中");
            progress.Show();

            convert_bufont(progress);

            progress.exitForm();

            this.button_kill_font_issue.Enabled = true;
        }

        private void button_swap_objects_Click(object sender, RibbonControlEventArgs e)
        {
            var sr = GetSelectedShapeRange();
            if (sr == null) return;

            if (sr.Count == 2)
            {
                var ss = new List<PowerPoint.Shape>();
                foreach (PowerPoint.Shape s in sr) ss.Add(s);
                // save locations
                float center_x0 = ss[0].Left + (ss[0].Width / 2);
                float center_y0 = ss[0].Top+ (ss[0].Height/ 2);
                float height_0 = ss[0].Height;
                float width_0 = ss[0].Width;

                float center_x1 = ss[1].Left + (ss[1].Width / 2);
                float center_y1 = ss[1].Top+ (ss[1].Height/ 2);
                float height_1 = ss[1].Height;
                float width_1 = ss[1].Width;

                // apply swap
                ss[0].Left = center_x1 - (width_0 / 2); 
                ss[0].Top = center_y1 - (height_0 / 2); 

                ss[1].Left = center_x0 - (width_1 / 2); 
                ss[1].Top = center_y0 - (height_1 / 2); 
            }


        }
    }
}
