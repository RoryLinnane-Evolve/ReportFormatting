using Microsoft.Office.Tools.Ribbon;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

using Word = Microsoft.Office.Interop.Word;
using System.Windows.Forms;

namespace ReportFormatting
{
    public partial class FormattingRibbon
    {
        private void FormattingRibbon_Load(object sender, RibbonUIEventArgs e)
        {

        }

        private void btnSelectImages_Click(object sender, RibbonControlEventArgs e)
        {
            //Init
            Word.Document doc = Globals.ThisAddIn.Application.ActiveDocument;
            Word.Application app = Globals.ThisAddIn.Application;
            doc.Paragraphs.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter;

            doc.Paragraphs.SpaceAfter = 0;
            doc.Paragraphs.SpaceBefore = 0;


            //File dialog
            OpenFileDialog fileDialog = new OpenFileDialog();
            fileDialog.Filter = "image files (*.bmp;*.jpg;*.jpeg;*.png;*.gif;*.tif;*.tiff)|*.bmp;*.jpg;*.jpeg;*.png;*.gif;*.tif;*.tiff";
            fileDialog.Multiselect = true;
            fileDialog.InitialDirectory = @"C:\";
            var result = fileDialog.ShowDialog();

            //Checking for files
            if (result != DialogResult.OK)
            {
                return;
            }

            //Insert all images
            foreach (var sfileName in fileDialog.FileNames)
            {
                //Get Selection
                Word.Range act_Image_Range = app.Selection.Range;

                //Insert Image
                Word.InlineShape inlineShape = doc.InlineShapes.AddPicture(sfileName, LinkToFile: false, SaveWithDocument: true, Range: act_Image_Range);

                //scaling
                inlineShape.LockAspectRatio = Microsoft.Office.Core.MsoTriState.msoTrue;
                inlineShape.Height = app.CentimetersToPoints(10);

                inlineShape.Line.Weight = 1;
                inlineShape.Line.Visible = Microsoft.Office.Core.MsoTriState.msoTrue;
            }
        }

        private void btnSelectFigures_Click(object sender, RibbonControlEventArgs e)
        {
            //Init
            Word.Document doc = Globals.ThisAddIn.Application.ActiveDocument;
            Word.Application app = Globals.ThisAddIn.Application;
            doc.Paragraphs.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter;

            //File dialog
            OpenFileDialog fileDialog = new OpenFileDialog();
            fileDialog.Filter = "image files (*.bmp;*.jpg;*.jpeg;*.png;*.gif;*.tif;*.tiff)|*.bmp;*.jpg;*.jpeg;*.png;*.gif;*.tif;*.tiff";
            fileDialog.Multiselect = true;
            fileDialog.InitialDirectory = @"C:\";
            var result = fileDialog.ShowDialog();

            //Checking for files
            if (result != DialogResult.OK)
            {
                return;
            }

            //Insert all images
            foreach (var sfileName in fileDialog.FileNames)
            {
                //Get Selection
                Word.Range act_Image_Range = app.Selection.Range;

                //Insert Image
                Word.InlineShape inlineShape = doc.InlineShapes.AddPicture(sfileName, LinkToFile: false, SaveWithDocument: true, Range: act_Image_Range);

                //scaling
                inlineShape.LockAspectRatio = Microsoft.Office.Core.MsoTriState.msoTrue;
                inlineShape.Height = app.CentimetersToPoints(10);

                inlineShape.Line.Weight = 1.5f;
                inlineShape.Line.Visible = Microsoft.Office.Core.MsoTriState.msoTrue;
            }
        }

        private void btnFormatLines_Click(object sender, RibbonControlEventArgs e)
        {
            Word.Document doc = Globals.ThisAddIn.Application.ActiveDocument;
            Word.Application app = Globals.ThisAddIn.Application;
            doc.Paragraphs.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter;
            doc.Paragraphs.LineSpacingRule = Word.WdLineSpacing.wdLineSpace1pt5;
            doc.Paragraphs.SpaceAfter = 0;
            doc.Paragraphs.SpaceBefore = 0;
            doc.Paragraphs.SpaceAfterAuto = 0;
            doc.Paragraphs.SpaceBeforeAuto = 0;
        }
    }
}
