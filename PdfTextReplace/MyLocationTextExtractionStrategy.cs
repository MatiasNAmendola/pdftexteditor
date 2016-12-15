//	PdfTextReplace.MyLocationTextExtractionStrategy
//	Copyright (c) 2016, Matias Nahuel Améndola soporte.esolutions@gmail.com
//
//	MIT license (http://en.wikipedia.org/wiki/MIT_License)
//
//	Permission is hereby granted, free of charge, to any person obtaining a copy
//	of this software and associated documentation files (the "Software"), to deal
//	in the Software without restriction, including without limitation the rights 
//	to use, copy, modify, merge, publish, distribute, sublicense, and/or sell copies
//	of the Software, and to permit persons to whom the Software is furnished to do so, 
//	subject to the following conditions:
//
//	The above copyright notice and this permission notice shall be included in all 
//	copies or substantial portions of the Software.
//
//	THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR IMPLIED, 
//	INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY, FITNESS FOR A PARTICULAR
//	PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE AUTHORS OR COPYRIGHT HOLDERS BE LIABLE 
//	FOR ANY CLAIM, DAMAGES OR OTHER LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE,
//	ARISING FROM, OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE SOFTWARE.

using iTextSharp.text.pdf.parser;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace PdfTextReplace
{
    public class MyLocationTextExtractionStrategy : LocationTextExtractionStrategy
    {
        // Hold each coordinate
        public List<RectAndText> myPoints = new List<RectAndText>();

        // Automatically called for each chunk of text in the PDF
        public override void RenderText(TextRenderInfo renderInfo)
        {
            base.RenderText(renderInfo);

            // Get the bounding box for the chunk of text
            var bottomLeft = renderInfo.GetDescentLine().GetStartPoint();
            var topRight = renderInfo.GetAscentLine().GetEndPoint();

            // Create a rectangle from it
            var rect = new iTextSharp.text.Rectangle(
                                                    bottomLeft[Vector.I1],
                                                    bottomLeft[Vector.I2],
                                                    topRight[Vector.I1],
                                                    topRight[Vector.I2]
                                                    );

            // Add this to our main collection
            this.myPoints.Add(new RectAndText(rect, renderInfo.GetText()));
        }
    }
}
