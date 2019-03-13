using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Word = Microsoft.Office.Interop.Word;

namespace wordtopdf
{
    class Program
    {
        static void Main(string[] args)
        {
            Word.Application app = new Word.Application();

            string[] docs = Directory.GetFiles(Environment.CurrentDirectory + @"\docs","*.docx",SearchOption.AllDirectories);

            foreach(string docPath in docs)
            {
                var doc = app.Documents.Open(docPath);
                string pdfPath = docPath.Replace("docx", "pdf");
                doc.SaveAs2(pdfPath, Word.WdSaveFormat.wdFormatPDF);
                doc.Close();
                Console.WriteLine($"Converted {pdfPath}");
                //use RAM disk? 
                //https://blogs.technet.microsoft.com/windowsinternals/2017/08/25/how-to-create-a-ram-disk-in-windows-server/
            }

        }
    }
}
