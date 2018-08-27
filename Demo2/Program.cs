
using Spire.Doc;
using Spire.Doc.Documents;
using Spire.Doc.Fields;
using System;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Reflection;
using Word = Microsoft.Office.Interop.Word;


namespace Demo2
{
    class Program
    {
        public static void WriteWord(string file_doc, string dir_pic)
        {
            #region Write
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
            doc.SaveToFile(file_doc, FileFormat.Docx2013);
            #endregion



            object path = file_doc;
            Object Nothing = Missing.Value;

            Word.Application wordApp = new Word.ApplicationClass
            {
                Visible = true//使文档不可见
            }; //初始化


            object unite = Word.WdUnits.wdStory;
            Word.Document wordDoc = wordApp.Documents.Open(file_doc, Type.Missing, false);
            #region 插入图片、居中显示，设置图片的绝对尺寸和缩放尺寸，并给图片添加标题

            wordApp.Selection.Find.Replacement.ClearFormatting();
            wordApp.Selection.Find.ClearFormatting();
            wordApp.Selection.Find.Text = "Evaluation Warning: The document was created with Spire.Doc for .NET.";//需要被替换的文本
            wordApp.Selection.Find.Replacement.Text = "";//替换文本 
            object oMissing = Missing.Value;
            object replace = Word.WdReplace.wdReplaceAll;
            //执行替换操作
            wordApp.Selection.Find.Execute(
            ref oMissing, ref oMissing,
            ref oMissing, ref oMissing,
            ref oMissing, ref oMissing,
            ref oMissing, ref oMissing, ref oMissing,
            ref oMissing, ref replace,
            ref oMissing, ref oMissing,
            ref oMissing, ref oMissing);
            //wordApp.Selection.HomeKey(ref unite, ref Nothing); //将光标移动到文档末尾
            //wordApp.Selection.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter;//居中显示图片
            //设置图片宽高的绝对大小

            //wordDoc.InlineShapes[1].Width = 200;
            //wordDoc.InlineShapes[1].Height = 150;
            //按比例缩放大小

            //wordDoc.InlineShapes[1].ScaleWidth = 20;//缩小到20% ？
            //wordDoc.InlineShapes[1].ScaleHeight = 20;

            //在图下方居中添加图片标题

            //wordDoc.Content.InsertAfter("\n");//这一句与下一句的顺序不能颠倒，原因还没搞透
            //wordApp.Selection.EndKey(ref unite, ref Nothing);
            //wordApp.Selection.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter;
            //wordApp.Selection.Font.Size = 10;//字体大小
            //wordApp.Selection.TypeText("图1 测试图片\n");

            #endregion
            object format = Word.WdSaveFormat.wdFormatDocument;
            wordDoc.SaveAs(ref path, ref format, ref Nothing, ref Nothing, ref Nothing, ref Nothing, ref Nothing, ref Nothing, ref Nothing, ref Nothing, ref Nothing, ref Nothing, ref Nothing, ref Nothing, ref Nothing, ref Nothing);
            wordDoc.Close(ref Nothing, ref Nothing, ref Nothing);
            //关闭wordApp组件对象
            wordApp.Quit(ref Nothing, ref Nothing, ref Nothing);

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


            WriteWord($@"{AppDomain.CurrentDomain.BaseDirectory}{1}.doc", @"C:\User\Project\PDFTOWORD\PDFTOWORD\bin\x64\Debug\File\PDF\笔记本电脑推荐-258293\pics");

            Console.WriteLine("OJBK");
            Console.ReadLine();

        }
    }
}
