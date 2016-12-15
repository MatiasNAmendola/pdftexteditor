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
