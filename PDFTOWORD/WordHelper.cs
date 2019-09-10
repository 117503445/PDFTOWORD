using Spire.Doc;
using Spire.Doc.Documents;
using Spire.Doc.Fields;
using System;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Reflection;
using Word = Microsoft.Office.Interop.Word;

namespace PDFTOWORD
{
    public static class WordHelper
    {
        public static void WriteWord(string file_doc, string dir_pic)
        {
            #region 生成Word
            //实例化一个Document对象
            Document doc = new Document();
            //添加section和段落
            Section section = doc.AddSection();
            Paragraph para = section.AddParagraph();
            DirectoryInfo info = new DirectoryInfo(dir_pic);
            var f = info.GetFiles();
            var pics = (from x in f orderby int.Parse(x.Name.Substring(0, x.Name.Length - 4)) select x.FullName).ToList();//不然11.jpg在2.jpg前面
            for (int i = 0; i < pics.Count; i++)
            {
                Image image = Image.FromFile(pics[i]);
                DocPicture picture = doc.Sections[0].Paragraphs[0].AppendPicture(image);
                image.Dispose();
                picture.TextWrappingStyle = TextWrappingStyle.Inline;//设置文字环绕方式
                doc.Sections[0].Paragraphs[0].AppendText(Environment.NewLine);
            }

            doc.SaveToFile(file_doc, FileFormat.Docx2013);//保存到文档
            #endregion
            #region 去水印
            object path = file_doc;
            Object Nothing = Missing.Value;
            Word.Application wordApp = new Word.ApplicationClass
            {
                Visible = false//使文档不可见
            }; //初始化
            object unite = Word.WdUnits.wdStory;
            Word.Document wordDoc = wordApp.Documents.Open(file_doc, Type.Missing, false);
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
            object format = Word.WdSaveFormat.wdFormatDocumentDefault;
            wordDoc.SaveAs(ref path, ref format, ref Nothing, ref Nothing, ref Nothing, ref Nothing, ref Nothing, ref Nothing, ref Nothing, ref Nothing, ref Nothing, ref Nothing, ref Nothing, ref Nothing, ref Nothing, ref Nothing);
            wordDoc.Close(ref Nothing, ref Nothing, ref Nothing);
            //关闭wordApp组件对象
            wordApp.Quit(ref Nothing, ref Nothing, ref Nothing);
            #endregion
        }
    }
}
