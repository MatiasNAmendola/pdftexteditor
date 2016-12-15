//	PdfTextEditor.Program
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

namespace PdfTextEditor
{
    class Program
    {
        static void Main(string[] args)
        {
            string src = @"C:\Users\mamendola\Documents\csv\source.pdf";
            string dest = @"C:\Users\mamendola\Documents\csv\changed.pdf";
            string csv = @"C:\Users\mamendola\Documents\csv\process.csv";

            DictionaryCollection<int, string> dictionaryCol = ReadCsv(csv, ',');

            foreach (Dictionary<int, string> dictionary in dictionaryCol)
            {
                string targetedText = dictionary[1];
                string newText = dictionary[2];

                ExceptionHandler exHandler = new ExceptionHandler();
                PdfTextReplace.PdfTextManager manager = new PdfTextReplace.PdfTextManager(exHandler);

                manager.ReplaceText(src, dest, targetedText, newText);
            }
        }

        /// <summary>
        /// Read CSV File
        /// </summary>
        /// <param name="src">Source</param>
        /// <param name="delimiter">Delimiter</param>
        /// <returns>DictionaryCollection</returns>
        static DictionaryCollection<int, string> ReadCsv(string src, char delimiter = ';')
        {
            DictionaryCollection<int, string> dictionaryCol = new DictionaryCollection<int, string>();

            using (CsvHandler.Csv.CsvReader csv = new CsvHandler.Csv.CsvReader(new System.IO.StreamReader(src), true, delimiter))
            {
                int fieldCount = csv.FieldCount;

                string[] headers = csv.GetFieldHeaders();

                Dictionary<int, string> dictionary = new Dictionary<int, string>();

                while (csv.ReadNextRecord())
                {
                    for (int i = 0; i < fieldCount; i++)
                    {
                        dictionary.Add(i, csv[i]); // headers[i]
                    }

                    dictionaryCol.Add(dictionary);
                }
            }

            return dictionaryCol;
        }
    }
}
