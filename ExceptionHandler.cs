using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace PdfTextEditor
{
    public class ExceptionHandler : PdfTextReplace.IExceptionHandler
    {
        public void Manage(System.Runtime.InteropServices._Exception ex)
        {
            Console.WriteLine(string.Format("{0}: {1}", ex.GetType().FullName, ex.Message));
            Console.ReadKey();
        }
    }
}
