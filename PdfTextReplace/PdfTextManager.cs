//	PdfTextReplace.PdfTextManager
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

using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace PdfTextReplace
{
    public class PdfTextManager
    {
        #region Constants

        public const string DEFAULT_FONTNAME = iTextSharp.text.FontFactory.HELVETICA;
        public const float DEFAULT_FONTSIZE = 15f;
        public const float ADJUST_MARGIN = 15f;
        public const float MARGIN_PAGE = 15f;

        #endregion
        #region Attributes

        private string _fontname = DEFAULT_FONTNAME;
        private float _maxFontSize = DEFAULT_FONTSIZE;
        private bool _responsive = false;
        private iTextSharp.text.BaseColor _backColor = iTextSharp.text.BaseColor.WHITE;
        private IExceptionHandler _exceptionHandler = null;      

        #endregion
        #region Properties

        /// <summary>
        /// Exception handler
        /// </summary>
        public IExceptionHandler ExceptionHandler
        {
            get { return _exceptionHandler; }
            set { _exceptionHandler = value; }
        }

        /// <summary>
        /// Fontname
        /// </summary>
        public string Fontname
        {
            get { return _fontname; }
            set
            { 
                if (string.IsNullOrEmpty(value))
                {
                    value = DEFAULT_FONTNAME;
                }

                _fontname = value; 
            }
        }

        /// <summary>
        /// Max font size
        /// </summary>
        public float MaxFontSize
        {
            get { return _maxFontSize; }
            set 
            {
                if (value <= 0)
                {
                    value = DEFAULT_FONTSIZE;
                }

                _maxFontSize = value; 
            }
        }

        /// <summary>
        /// ¿Responsive?
        /// </summary>
        public bool Responsive
        {
            get { return _responsive; }
            set { _responsive = value; }
        }

        /// <summary>
        /// Background color
        /// </summary>
        public iTextSharp.text.BaseColor BackColor
        {
            get { return _backColor; }
            set
            {
                if (value == null)
                {
                    value = iTextSharp.text.BaseColor.WHITE;
                }

                _backColor = value;
            }
        }

        #endregion
        #region Constructor

        /// <summary>
        /// Constructor
        /// </summary>
        /// <param name="exceptionHandler">Exception handler</param>
        public PdfTextManager(IExceptionHandler exceptionHandler)
        {
            _exceptionHandler = exceptionHandler;
        }

        #endregion
        #region Private methods

        /// <summary>
        /// Get characters location
        /// </summary>
        /// <param name="reader">PDF reader</param>
        /// <param name="page">Page N°</param>
        /// <param name="targetedText">Targeted text</param>
        /// <returns>RectAndText collection or null</returns>
        private List<RectAndText> GetCharactersLocation(iTextSharp.text.pdf.PdfReader reader, int page, string targetedText)
        {
            // Create an instance of our strategy
            var t = new MyLocationTextExtractionStrategy();

            // Parse page N of the document above
            string content = iTextSharp.text.pdf.parser.PdfTextExtractor.GetTextFromPage(reader, page, t);

            string temp = content.Replace("\n", "");
            bool founded = temp.Contains(targetedText);

            if (!temp.Contains(targetedText))
            {
                return null;
            }

            char[] target = targetedText.ToCharArray();

            int first = -1;
            int count = 0;

            List<RectAndText> found = new List<RectAndText>();

            for (int i = 0; i < t.myPoints.Count; i++)
            {
                if (t.myPoints[i].Text == target[count].ToString())
                {
                    found.Add(t.myPoints[i]);

                    if (first == -1)
                    {
                        first = i;
                    }


                    if (count == target.Length - 1)
                    {
                        break;
                    }

                    count++;
                }
                else
                {
                    found.Clear();
                    count = 0;
                    first = -1;
                }
            }

            return found;
        }

        /// <summary>
        /// Generate rectangle
        /// </summary>
        /// <param name="characters">RectAndText collection</param>
        /// <returns>Rectangle</returns>
        private iTextSharp.text.Rectangle GenRectangle(List<RectAndText> characters, iTextSharp.text.Rectangle pageSize)
        {
            float llx = -1;
            float lly = -1;
            float urx = -1;
            float ury = -1;

            foreach (RectAndText character in characters)
            {
                iTextSharp.text.Rectangle rect = character.Rect;

                if (llx == -1 || rect.Left < llx)
                {
                    llx = rect.Left;
                }

                if (urx == -1 || urx < rect.Right)
                {
                    urx = rect.Right;
                }

                if (lly == -1 || rect.Bottom < lly)
                {
                    lly = rect.Bottom;
                }

                if (ury == -1 || ury < rect.Top)
                {
                    ury = rect.Top;
                }
            }

            if (llx == -1 || lly == -1 || urx == -1 || ury == -1)
            {
                return null;
            }
            else
            {
                if (_responsive)
                {
                    llx = pageSize.Left + MARGIN_PAGE;
                    urx = pageSize.Right - MARGIN_PAGE;
                }

                return new iTextSharp.text.Rectangle(llx, lly, urx + ADJUST_MARGIN, ury);
            }
        }

        /// <summary>
        /// Generate column text
        /// </summary>
        /// <param name="cb">Content byte</param>
        /// <param name="container">Container</param>
        /// <param name="text">Text to add</param>
        /// <param name="font">Font</param>
        /// <returns>ColumnText object</returns>
        private iTextSharp.text.pdf.ColumnText GenColumnText(iTextSharp.text.pdf.PdfContentByte cb, iTextSharp.text.Rectangle container, string text, iTextSharp.text.Font font)
        {
            int width = Convert.ToInt32(container.Width);
            int height = Convert.ToInt32(container.Height);

            iTextSharp.text.pdf.ColumnText ct = new iTextSharp.text.pdf.ColumnText(cb);
                        
            ct.SetSimpleColumn(new iTextSharp.text.Rectangle(container.Left, container.Top, width, height));

            iTextSharp.text.Paragraph ptext = new iTextSharp.text.Paragraph(text, font);

            ct.AddElement(ptext);

            return ct;
        }

        /// <summary>
        /// Calculate font size
        /// </summary>
        /// <param name="width">Width</param>
        /// <param name="height">Height</param>
        /// <param name="text">Text to add</param>
        /// <returns></returns>
        private float CalFontSize(float width, float height, string text)
        {
            iTextSharp.text.Font font = iTextSharp.text.FontFactory.GetFont(_fontname);
            iTextSharp.text.Rectangle rectangle = new iTextSharp.text.Rectangle(width, height);
            float newSize = iTextSharp.text.pdf.ColumnText.FitText(font, text, rectangle, _maxFontSize, iTextSharp.text.pdf.PdfWriter.RUN_DIRECTION_NO_BIDI);
            return newSize;
        }

        /// <summary>
        /// Generate font
        /// </summary>
        /// <param name="container">Container</param>
        /// <param name="text">Text to add</param>
        /// <returns>Font object</returns>
        private iTextSharp.text.Font GenFont(iTextSharp.text.Rectangle container, string text)
        {
            int width = Convert.ToInt32(container.Width);
            int height = Convert.ToInt32(container.Height);

            float newSize = this.CalFontSize(width, height, text);

            iTextSharp.text.Font font = iTextSharp.text.FontFactory.GetFont(_fontname, newSize, iTextSharp.text.Font.NORMAL);
            return font;
        }

        /// <summary>
        /// Get text size
        /// </summary>
        /// <param name="text">Text to add</param>
        /// <param name="baseFont">Base font</param>
        /// <param name="fontSize">Font size</param>
        /// <returns>Text size</returns>
        private float GetTextSize(string text, iTextSharp.text.pdf.BaseFont baseFont, float fontSize)
        {
            float ascend = baseFont.GetAscentPoint(text, fontSize);
            float descend = baseFont.GetDescentPoint(text, fontSize);
            return ascend - descend;
        }

        /// <summary>
        /// Gets the font size and verifies if this size fits the container (simulated)
        /// </summary>
        /// <param name="container">Container</param>
        /// <param name="text">Text to add</param>
        /// <param name="font">Font</param>
        /// <returns>¿Is font size adjust? Returns new/adjusted font</returns>
        private KeyValuePair<iTextSharp.text.Font, bool> SimulateFont(iTextSharp.text.Rectangle container, string text, iTextSharp.text.Font font = null)
        {
            if (font == null)
            {
                font = this.GenFont(container, text);
            }

            float simulatedSize = 0;

            using (System.IO.MemoryStream ms = new System.IO.MemoryStream())
            {
                //Create an iTextSharp Document which is an abstraction of a PDF but **NOT** a PDF
                using (iTextSharp.text.Document doc = new iTextSharp.text.Document(iTextSharp.text.PageSize.A4, 10f, 10f, 10f, 0f))
                {
                    //Create a writer that's bound to our PDF abstraction and our stream
                    using (iTextSharp.text.pdf.PdfWriter tmpWriter = iTextSharp.text.pdf.PdfWriter.GetInstance(doc, ms))
                    {
                        doc.Open();
                        doc.Add(new iTextSharp.text.Chunk(""));

                        iTextSharp.text.pdf.PdfContentByte cb = new iTextSharp.text.pdf.PdfContentByte(tmpWriter);
                        iTextSharp.text.pdf.ColumnText ct = this.GenColumnText(cb, container, text, font);

                        float textSize = this.GetTextSize(text, font.BaseFont, font.Size);

                        ct.Go();
                        tmpWriter.CloseStream = false;

                        simulatedSize = ((ct.LinesWritten - 1) * ct.Leading) + font.Size;
                    }
                }
            }

            bool success = simulatedSize <= container.Height;

            return new KeyValuePair<iTextSharp.text.Font, bool>(font, success);
        }

        /// <summary>
        /// Get measured font
        /// </summary>
        /// <param name="container">Container</param>
        /// <param name="text">Text to add</param>
        /// <returns></returns>
        private iTextSharp.text.Font GetMeasuredFont(iTextSharp.text.Rectangle container, string text)
        {
            KeyValuePair<iTextSharp.text.Font, bool> simulatedFont = this.SimulateFont(container, text);
            
            if (simulatedFont.Value == false)
            {
                float minSize = 0.01f;

                do
                {
                    simulatedFont.Key.Size -= minSize;
                    simulatedFont = this.SimulateFont(container, text, simulatedFont.Key);

                } while (!simulatedFont.Value);
            }
            
            return simulatedFont.Key;

        }

        /// <summary>
        /// Creates an image that is the size i need to hide the text i'm interested in removing
        /// </summary>
        /// <param name="container">Container</param>
        /// <param name="cb">Content byte</param>
        private void AddBackgroud(iTextSharp.text.Rectangle container, iTextSharp.text.pdf.PdfContentByte cb)
        {
            int width = Convert.ToInt32(container.Width - ADJUST_MARGIN);
            int height = Convert.ToInt32(container.Height);

            if (_backColor == null)
            {
                _backColor = iTextSharp.text.BaseColor.WHITE;
            }

            iTextSharp.text.Image backgroud = iTextSharp.text.Image.GetInstance(new System.Drawing.Bitmap(width, height), _backColor);

            // Sets the position that the image needs to be placed (ie the location of the text to be removed)
            backgroud.SetAbsolutePosition(container.Left, container.Bottom);

            // Adds the image to the output pdf
            cb.AddImage(backgroud);
        }

        #endregion
        #region Public methods

        /// <summary>
        /// Replace text.
        /// </summary>
        /// <param name="src">Source</param>
        /// <param name="dest">Destination</param>
        /// <param name="targetedText">Targeted text</param>
        /// <param name="newText">New text</param>
        /// <returns>True if success</returns>
        public bool ReplaceText(string src, string dest, string targetedText, string newText)
        {
            iTextSharp.text.pdf.PdfReader reader = null;

            try
            {
                reader = new iTextSharp.text.pdf.PdfReader(src);
            }
            catch(Exception ex)
            {
                _exceptionHandler.Manage(ex);
                return false;
            }

            int page = 0;
            List<RectAndText> found = null;

            try
            {
                for (int i = 1; i <= reader.NumberOfPages; i++)
                {
                    found = this.GetCharactersLocation(reader, i, targetedText);

                    if (found == null)
                    {
                        continue;
                    }
                    else
                    {
                        page = i;
                        break;
                    }
                }

                if (found == null)
                {
                    throw new TargetedTextNotFound("Text to replace not found");
                }
            }
            catch (TargetedTextNotFound ex)
            {
                _exceptionHandler.Manage(ex);
                return false;
            }
            

            iTextSharp.text.pdf.PdfDictionary dict = reader.GetPageN(page);
            iTextSharp.text.Rectangle pageSize = reader.GetPageSize(dict);

            iTextSharp.text.Rectangle container = this.GenRectangle(found, pageSize);
            iTextSharp.text.Font font = this.GetMeasuredFont(container, newText);

            try
            {
                using (System.IO.Stream outputPdfStream = new System.IO.FileStream(dest, System.IO.FileMode.Create, System.IO.FileAccess.Write, System.IO.FileShare.ReadWrite))
                using (iTextSharp.text.pdf.PdfStamper stamper = new iTextSharp.text.pdf.PdfStamper(reader, outputPdfStream) { FormFlattening = true, FreeTextFlattening = true })
                {
                    iTextSharp.text.pdf.PdfContentByte over = stamper.GetOverContent(page);

                    // Creates an image that is the size i need to hide the text i'm interested in removing
                    this.AddBackgroud(container, over);

                    iTextSharp.text.pdf.ColumnText ct = this.GenColumnText(over, container, newText, font);

                    ct.Go();
                }
            }
            catch(System.IO.IOException ex)
            {
                _exceptionHandler.Manage(ex);
                return false;
            }
            catch(System.Security.SecurityException ex)
            {
                _exceptionHandler.Manage(ex);
                return false;
            }
            catch (UnauthorizedAccessException ex)
            {
                _exceptionHandler.Manage(ex);
                return false;
            }
            catch(Exception ex)
            {
                _exceptionHandler.Manage(ex);
                return false;
            }
            
            reader.Close();
            return true;
        }

        #endregion
        #region Public static methods

        /// <summary>
        /// Get available fonts
        /// </summary>
        public static string GetAvailableFonts()
        {
            int totalfonts = iTextSharp.text.FontFactory.RegisterDirectory("C:\\WINDOWS\\Fonts");

            StringBuilder sb = new StringBuilder();
            sb.Append("All Fonts:");
            sb.Append(Environment.NewLine);
            sb.Append(Environment.NewLine);

            foreach (string fontname in iTextSharp.text.FontFactory.RegisteredFonts)
            {
                sb.Append(fontname);
                sb.Append(Environment.NewLine);
            }

            sb.Append(Environment.NewLine);
            sb.Append(Environment.NewLine);
            sb.Append(string.Format("Total Fonts: {0}", totalfonts));

            return sb.ToString();
        }

        /// <summary>
        /// Register font
        /// </summary>
        /// <param name="fontpath">Path</param>
        /// <param name="fontname">Filename (without extension, *.ttf)</param>
        /// <returns></returns>
        public static iTextSharp.text.pdf.BaseFont RegisterFont(string fontpath, string fontname)
        {
            iTextSharp.text.pdf.BaseFont customfont = iTextSharp.text.pdf.BaseFont.CreateFont(fontpath + fontname + ".ttf", iTextSharp.text.pdf.BaseFont.CP1252, iTextSharp.text.pdf.BaseFont.EMBEDDED);
            return customfont;
        }

        #endregion
    }
}
