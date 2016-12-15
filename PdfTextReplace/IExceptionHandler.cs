using System;
namespace PdfTextReplace
{
    public interface IExceptionHandler
    {
        void Manage(System.Runtime.InteropServices._Exception ex);
    }
}
