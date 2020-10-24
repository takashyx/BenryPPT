using System;
using System.Collections.Generic;
using System.Drawing;
using System.Drawing.Text;
using System.Linq;
using System.Security.Principal;
using System.Text;
using System.Windows.Forms;
using Microsoft.Office.Tools.Ribbon;
using Microsoft.Office.Interop.PowerPoint;
using System.Diagnostics;
using System.Security.Cryptography;

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
        
        private void UnifyFont_ConvertFonts()
        {

            string targetFont = Settings.Default.targetFontForUnifyFonts;

            try
            {
                // convert slide items
                var slides = Globals.ThisAddIn.Application.ActivePresentation.Slides;

                foreach (Slide slide in slides)
                {
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
                            foreach(int i in  Enumerable.Range(1,shape.Table.Rows.Count))
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
                            if (shape.Chart.HasTitle)  shape.Chart.ChartTitle.Font.Name = targetFont; 
                            if (shape.Chart.HasLegend) shape.Chart.Legend.Font.Name = targetFont;
                        }
                    }
                }
            }
            catch(Exception ex)
            {
                System.Windows.Forms.MessageBox.Show("Benry Error: " + ex);
            }
            Debug.WriteLine("Converting Done.");
        }

        private void UnifyFont_Click(object sender, RibbonControlEventArgs e)
        {
            // read target font
            string FontFamilyName = Settings.Default.targetFontForUnifyFonts;

            // splash screen
            UnifyFont_ConvertFonts();
        }

        private void dropDown_UnifyFontsTargetFont_SelectionChanged(object sender, RibbonControlEventArgs e)
        {
            Settings.Default.targetFontForUnifyFonts = dropDown_UnifyFontsTargetFont.SelectedItem.ToString();
            Settings.Default.Save();
        }
    }
}
