﻿using System;
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
        public long PdfSize { get; set; }
        /// <summary>
        /// 横向连接多个图片,要求高度相同
        /// </summary>
        /// <param name="imgs"></param>
        /// <returns></returns>
        private Bitmap CombineImages(Bitmap[] imgs)
        {
            int totalWidth = 0;
            for (int i = 0; i < imgs.Length; i++)
            {
                totalWidth += imgs[i].Width;
                if (imgs[i].Height!=imgs[0].Height)
                {
                    throw new Exception("高度不同!");
                }
            }
            
            Bitmap bitmap = new Bitmap(totalWidth,imgs[0].Height);
            int twidth = 0;
            Console.WriteLine("1");
            for (int i = 0; i < imgs.Length; i++)
            {
                for (int x = 0; x < imgs[i].Width; x++)
                {
                    for (int y = 0; y < imgs[i].Height; y++)
                    {
                        try
                        {
                            bitmap.SetPixel(x + twidth, y, imgs[i].GetPixel(x, y));
                        }
                        catch (Exception ex)
                        {
                            Console.WriteLine(ex);
                            TLib.Software.Logger.WriteException(ex);
                        }
                        
                    }
                    
                }
               
                twidth += imgs[i].Width;
            }
            Console.WriteLine("Fin");
            return bitmap;
        }


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
            string dir_WorkPlace = App.Dir_File_PDF + $"{fileInfo.Name.Substring(0, fileInfo.Name.Length - 4)}-{PdfSize}";
            if (Directory.Exists(dir_WorkPlace))
            {
                Console.WriteLine("已存在dir_WorkPlace");
            }
            else
            {
                Directory.CreateDirectory(dir_WorkPlace);
            }
            await Task.Run(() =>
            {
                var pdf = PDFFile.Open(Dir_SourcePdf);
                List<Bitmap> pics = new List<Bitmap>();
                int totalWidth = 0;
                for (int i = 0; i < pdf.PageCount; i++)
                {
                TbInfo.Dispatcher.Invoke(()=> { TbInfo.Text = GetTbInfoText(Dir_SourcePdf, $"{i+1}/{pdf.PageCount}", false); });
                    var p = pdf.GetPageImage(i, 150);
                    pics.Add(p);
                    totalWidth += p.Width;
                }
                Console.WriteLine(totalWidth);
                // Bitmap bitmap = new Bitmap(totalWidth, pics[0].Height);
                var p1 = CombineImages(pics.ToArray());
                p1.Save(@"C:\Users\117503445\Desktop\1.jpg");

                for (int i = 0; i < pics.Count; i++)
                {
                    Console.WriteLine(i);
                    pics[i].Save($@"C:\Users\117503445\Desktop\pics\{i}.jpg");
                }

                TbInfo.Dispatcher.Invoke(() => { TbInfo.Text = GetTbInfoText(Dir_SourcePdf, "预处理完成", false); });
            });
        }
        private string GetTbInfoText(string dir_SourcePdf, string strHandle, bool IsBuilt)
        {
            return $"正在处理的PDF文件路径:{dir_SourcePdf + NewLine}" +
             $"预处理状态:{(strHandle) + NewLine}" +
             $"是否生成过了DOC:{(IsBuilt ? "生成过了" : "未生成")}";
        }

        private void BtnEdit_Click(object sender, RoutedEventArgs e)
        {

        }
    }
}
