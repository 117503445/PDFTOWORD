
using System;
using System.Collections.Generic;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Text;
using System.Threading.Tasks;
using TLib.Software;
using Word = Microsoft.Office.Interop.Word;
using Spire.Doc;
using Spire.Doc.Documents;
using Spire.Doc.Fields;

namespace Demo1
{

    class Program
    {
        public static void WriteWord(string file_doc, string dir_pic)
        {
            //实例化一个Document对象
            Document doc = new Document();

            //添加section和段落
            Section section = doc.AddSection();
            Paragraph para = section.AddParagraph();
            DirectoryInfo info = new DirectoryInfo(dir_pic);
            var f = info.GetFiles();

            var pics = (from x in f orderby int.Parse(x.Name.Substring(0, x.Name.Length - 4)) select x.FullName).ToList();//不然11.jpg在2.jpg前面
            //加载图片到System.Drawing.Image对象, 使用AppendPicture方法将图片插入到段落
            for (int i = 0; i < pics.Count; i++)
            {
                Image image = Image.FromFile(pics[i]);
                DocPicture picture = doc.Sections[0].Paragraphs[0].AppendPicture(image);
                //设置文字环绕方式
                picture.TextWrappingStyle = TextWrappingStyle.Square;
                doc.Sections[0].Paragraphs[0].AppendText(Environment.NewLine);
                //指定图片位置
                //picture.HorizontalPosition = 50.0f;
                //picture.VerticalPosition = 50.0f;

                //设置图片大小
                //picture.Width = 100;
                //picture.Height = 100;
            }




            //保存到文档
            doc.SaveToFile("Image.docx", FileFormat.Docx2013);
        }
        static void Main(string[] args)
        {
            //List<Bitmap> b = new List<Bitmap>();
            //for (int i = 0; i < 60; i++)
            //{
            //    b.Add(new Bitmap($@"C:\Users\117503445\Desktop\pics\{i}.jpg"));
            //}
            //var m = CombineImages(b);
            //using (FileStream stream = File.Create(@"C:\Users\117503445\Desktop\3.png"))
            //{
            //    m.Save(stream, System.Drawing.Imaging.ImageFormat.Bmp);
            //}

            //Console.WriteLine("213.pdf".Substring(0, "213.pdf".Length - 4));
            //WriteWord($@"C:\Users\117503445\Desktop\pics\1.doc", @"C:\Users\117503445\Desktop\pics1");


            WriteWord($@"{AppDomain.CurrentDomain.BaseDirectory}{TimeStamp.Now}.doc", @"C:\User\Project\PDFTOWORD\PDFTOWORD\bin\x64\Debug\File\PDF\笔记本电脑推荐-258293\pics");

            Console.WriteLine("OJBK");
            Console.ReadLine();
            
        }
    }
}
