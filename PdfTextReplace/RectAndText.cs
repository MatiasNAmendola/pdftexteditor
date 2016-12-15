using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace PdfTextReplace
{
    /// <summary>
    /// Helper class that stores our rectangle and text
    /// </summary>
    public class RectAndText
    {
        public iTextSharp.text.Rectangle Rect;
        public String Text;
        public RectAndText(iTextSharp.text.Rectangle rect, String text)
        {
            this.Rect = rect;
            this.Text = text;
        }
    }
}
