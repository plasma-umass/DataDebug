using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Excel = Microsoft.Office.Interop.Excel;
using System.Text.RegularExpressions;
using Microsoft.FSharp.Core;
using System.Drawing;
using System.Drawing.Drawing2D;
using System.Drawing.Text;

using Workbook = Microsoft.Office.Interop.Excel.Workbook;
using Worksheet = Microsoft.Office.Interop.Excel.Worksheet;

namespace DataDebugMethods
{
    public static class Utility
    {
        public static AST.Reference ParseReferenceOfXLRange(Excel.Range rng, Workbook wb)
        {
            string rng_r1c1 = rng.Address[true, true, Excel.XlReferenceStyle.xlR1C1, false];
            FSharpOption<AST.Reference> r = ExcelParser.GetReference(rng_r1c1, wb, rng.Worksheet);

            if (FSharpOption<AST.Reference>.get_IsNone(r))
            {
                throw new Exception("Unimplemented address feature in address string: '" + rng_r1c1 + "'");
            }

            return r.Value;
        }

        public static AST.Address ParseXLAddress(Excel.Range rng, Workbook wb)
        {
            // we'll get an exception from the parser if this is not, in fact an address
            string rng_r1c1 = rng.Address[true, true, Excel.XlReferenceStyle.xlR1C1, false];
            return ExcelParser.GetAddress(rng_r1c1, wb, rng.Worksheet);
        }

        public static bool InsideRectangle(Excel.Range rng, AST.Reference rect, Workbook wb)
        {
            return ParseReferenceOfXLRange(rng, wb).InsideRef(rect);
        }

        public static bool InsideUsedRange(Excel.Range rng, Workbook wb)
        {
            return InsideRectangle(rng, UsedRange(rng, wb), wb);
        }

        public static AST.Reference UsedRange(Excel.Range rng, Workbook wb)
        {
            return ParseReferenceOfXLRange(rng.Worksheet.UsedRange, wb);
        }

        // borrowed from: http://chiragrdarji.wordpress.com/2008/05/09/generate-image-from-text-using-c-or-convert-text-in-to-image-using-c/
        public static Bitmap CreateBitmapImage(string sImageText, int fontsize)
        {
            Bitmap objBmpImage = new Bitmap(1, 1);

            int intWidth = 0;
            int intHeight = 0;

            // Create the Font object for the image text drawing.
            Font objFont = new Font("Arial", fontsize, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Pixel);

            // Create a graphics object to measure the text's width and height.
            Graphics objGraphics = Graphics.FromImage(objBmpImage);

            // This is where the bitmap size is determined.
            intWidth = (int)objGraphics.MeasureString(sImageText, objFont).Width;
            intHeight = (int)objGraphics.MeasureString(sImageText, objFont).Height;

            // Create the bmpImage again with the correct size for the text and font.
            objBmpImage = new Bitmap(objBmpImage, new Size(intWidth, intHeight));

            // Add the colors to the new bitmap.
            objGraphics = Graphics.FromImage(objBmpImage);

            // Set Background color
            objGraphics.Clear(Color.White);
            objGraphics.SmoothingMode = SmoothingMode.AntiAlias;
            objGraphics.TextRenderingHint = TextRenderingHint.AntiAlias;
            objGraphics.DrawString(sImageText, objFont, new SolidBrush(Color.FromArgb(102, 102, 102)), 0, 0);
            objGraphics.Flush();

            return (objBmpImage);
        }

    }
}
