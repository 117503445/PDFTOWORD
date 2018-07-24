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
        /// <summary>
        /// 横向连接多个图片,要求尺寸相同
        /// </summary>
        /// <param name="imgs"></param>
        /// <returns></returns>
        private static Bitmap CombineImages(List<Bitmap> imgs)
        {
            int totalWidth = 0;
            foreach (var item in imgs)
            {
                //Console.WriteLine(item.Width);
                if (imgs[0].Size!=item.Size)
                {
                    throw new Exception("不同尺寸的图片");
                }
                totalWidth += item.Width;
            }
            Console.WriteLine(imgs[0].Width);
            //Console.WriteLine(totalWidth);
            Bitmap bitNew = new Bitmap(totalWidth, imgs[0].Height);
            Graphics g = Graphics.FromImage(bitNew);//Create graphics object
            int printedWidth = 0;
            for (int i = 0; i < imgs.Count; i++)
            {
                g.DrawImage(imgs[i], imgs[0].Width*i, 0, imgs[0].Width, imgs[0].Height);
                printedWidth += imgs[i].Width;
            }
            g.Dispose();
            return bitNew;
        }

        static void Main(string[] args)
        {
            List<Bitmap> b = new List<Bitmap>();
            for (int i = 0; i < 60; i++)
            {
                b.Add(new Bitmap($@"C:\Users\117503445\Desktop\pics\{i}.jpg"));
            }
            var m = CombineImages(b);
            m.Save(@"C:\Users\117503445\Desktop\3.png", System.Drawing.Imaging.ImageFormat.Png);
            Console.WriteLine("OJBK");
            Console.ReadLine();
        }
    }
}
