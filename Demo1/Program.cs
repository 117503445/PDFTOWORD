using Spire.Pdf;
using System;
using System.Collections.Generic;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Text;
using System.Threading.Tasks;
using Word = Microsoft.Office.Interop.Word;
namespace Demo1
{

    class Program
    {
        public static void EditImage(string file_pic, string file_save, int x1, int y1, int x2, int y2)
        {
            Bitmap bitmap = new Bitmap(file_pic);
            Bitmap b = bitmap.Clone(new Rectangle(x1,y1,x2-x1,y2-y1),System.Drawing.Imaging.PixelFormat.Undefined);
            b.Save(file_save);
            bitmap.Dispose();
            b.Dispose();
        }
        static void Main(string[] args)
        {
            EditImage(@"C:\Users\117503445\Desktop\1.jpg", @"C:\Users\117503445\Desktop\2.jpg",10,10,10000,1000);
            //Console.WriteLine(a.Substring(0, a.Length - 4));
            return;
            object path = @"C:\Users\117503445\Desktop\1.doc";
            //string strContent;

            Object Nothing = Missing.Value;

            Word.Application wordApp = new Word.ApplicationClass
            {
                Visible = true//使文档可见
            }; //初始化
            object unite = Word.WdUnits.wdStory;
            Word.Document wordDoc = wordApp.Documents.Add(ref Nothing, ref Nothing, ref Nothing, ref Nothing);
            #region 插入图片、居中显示，设置图片的绝对尺寸和缩放尺寸，并给图片添加标题

            wordApp.Selection.EndKey(ref unite, ref Nothing); //将光标移动到文档末尾
            //图片文件的路径
            string filename = @"C:\Users\117503445\Desktop\1.PNG";
            string filename2 = @"C:\Users\117503445\Desktop\2.PNG";
            string filename3 = @"C:\Users\117503445\Desktop\3.PNG";
            //要向Word文档中插入图片的位置
            Object range = wordDoc.Paragraphs.Last.Range;
            //定义该插入的图片是否为外部链接
            Object linkToFile = false;               //默认,这里貌似设置为bool类型更清晰一些
            //定义要插入的图片是否随Word文档一起保存
            Object saveWithDocument = true;              //默认
            //使用InlineShapes.AddPicture方法(【即“嵌入型”】)插入图片
            wordDoc.InlineShapes.AddPicture(filename, ref linkToFile, ref saveWithDocument, ref range);
            //wordApp.Selection.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter;//居中显示图片
            wordDoc.InlineShapes.AddPicture(filename2, ref linkToFile, ref saveWithDocument, ref range);
            wordDoc.InlineShapes.AddPicture(filename3, ref linkToFile, ref saveWithDocument, ref range);
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
            Console.WriteLine("OJBK");
            Console.ReadLine();
        }
    }
}
