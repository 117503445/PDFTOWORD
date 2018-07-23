using Spire.Pdf;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Demo1
{
    class Program
    {
        static void Main(string[] args)
        {
            PdfDocument doc = new PdfDocument();
            doc.LoadFromFile(@"C:\Users\117503445\Desktop\答案4.pdf");
            doc.SaveToFile(@"C:\Users\117503445\Desktop\4.doc", FileFormat.DOC);
            //System.Diagnostics.Process.Start("图文版丽江旅游攻略大全.doc");
            Console.ReadLine();
        }
    }
}
