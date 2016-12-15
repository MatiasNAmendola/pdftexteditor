using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace PdfTextReplace
{
    public class TargetedTextNotFound : Exception
    {
        public TargetedTextNotFound(string message) : base(message)
        {
        }
    }
}
