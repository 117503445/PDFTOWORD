using System;
using System.IO;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;

using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Shapes;
using static System.Environment;
using O2S.Components.PDFRender4NET;
using System.Drawing;

namespace PDFTOWORD
{
    /// <summary>
    /// WdProcess.xaml 的交互逻辑
    /// </summary>
    public partial class WdProcess : Window
    {
        public string Dir_SourcePdf { get; set; }
        public string Dir_WorkPlace { get; set; }
        public long PdfSize { get; set; }



        public WdProcess(string dir_pdf)
        {
            InitializeComponent();
            Dir_SourcePdf = dir_pdf;
            Title = dir_pdf;
        }


        private async void Window_Loaded(object sender, RoutedEventArgs e)
        {
            TbInfo.Text = GetTbInfoText(Dir_SourcePdf, "正在处理", false);
            FileInfo fileInfo = new FileInfo(Dir_SourcePdf);
            PdfSize = fileInfo.Length;
            Dir_WorkPlace = App.Dir_File_PDF + $"{fileInfo.Name.Substring(0, fileInfo.Name.Length - 4)}-{PdfSize}/";
            Directory.CreateDirectory(Dir_WorkPlace);
            if (!File.Exists(Dir_WorkPlace + "pdf.jpg"))
            {
                await Task.Run(() =>
                {
                    var pdf = PDFFile.Open(Dir_SourcePdf);
                    List<Bitmap> pics = new List<Bitmap>();
                    int totalWidth = 0;
                    for (int i = 0; i < pdf.PageCount; i++)
                    {
                        TbInfo.Dispatcher.Invoke(() => { TbInfo.Text = GetTbInfoText(Dir_SourcePdf, $"{i + 1}/{pdf.PageCount}", false); });
                        var p = pdf.GetPageImage(i, 150);
                        pics.Add(p);
                        totalWidth += p.Width;
                    }
                    TbInfo.Dispatcher.Invoke(() => { TbInfo.Text = GetTbInfoText(Dir_SourcePdf, "正在合成图片,即将完成", false); });
                    var p1 = ImageHelper.CombineImages(pics);
                    p1.Save($@"{Dir_WorkPlace}pdf.jpg");

                    TbInfo.Dispatcher.Invoke(() => { TbInfo.Text = GetTbInfoText(Dir_SourcePdf, "预处理完成", false); });

                });
            }
            BtnEdit.IsEnabled = true;
            BtnEdit.Content = "开始编辑";


        }
        private string GetTbInfoText(string dir_SourcePdf, string strHandle, bool IsBuilt)
        {
            return $"正在处理的PDF文件路径:{dir_SourcePdf + NewLine}" +
             $"预处理状态:{(strHandle) + NewLine}" +
             $"是否生成过了DOC:{(IsBuilt ? "生成过了" : "未生成")}";
        }

        private void BtnEdit_Click(object sender, RoutedEventArgs e)
        {
            WdEdit wdEdit = new WdEdit(Dir_WorkPlace);
            wdEdit.Show();
        }
    }
}
