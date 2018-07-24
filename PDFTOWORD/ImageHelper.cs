using System;
using System.Collections.Generic;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace PDFTOWORD
{
    public class ImageHelper
    {
        /// <summary>
        /// 横向连接多个图片,要求尺寸相同
        /// </summary>
        /// <param name="imgs"></param>
        /// <returns></returns>
        public static Bitmap CombineImages(List<Bitmap> imgs)
        {
            int totalWidth = 0;
            foreach (var item in imgs)
            {
                //Console.WriteLine(item.Width);
                if (imgs[0].Size != item.Size)
                {
                    throw new Exception("不同尺寸的图片");
                }
                totalWidth += item.Width;
            }
            //Console.WriteLine(imgs[0].Height);
            //Console.WriteLine(totalWidth);
            Bitmap bitNew = new Bitmap(totalWidth, imgs[0].Height);
            Graphics g = Graphics.FromImage(bitNew);//Create graphics object
            int printedWidth = 0;
            for (int i = 0; i < imgs.Count; i++)
            {
                g.DrawImage(imgs[i], imgs[0].Width * i, 0, imgs[0].Width, imgs[0].Height);
                printedWidth += imgs[i].Width;
            }
            g.Dispose();
            return bitNew;
        }


        /// <summary>
        /// 截取图片
        /// </summary>
        /// <param name="file_pic"></param>
        /// <param name="file_save"></param>
        /// <param name="x1"></param>
        /// <param name="y1"></param>
        /// <param name="x2"></param>
        /// <param name="y2"></param>
        public static Bitmap EditImage(string file_pic, int x1, int y1, int x2, int y2)
        {
            //Console.WriteLine(x1);
            //Console.WriteLine(y1);

            //Console.WriteLine(x2);
            //Console.WriteLine(y2);
            Bitmap bitmap = new Bitmap(file_pic);
            Bitmap b = bitmap.Clone(new Rectangle(x1, y1, x2 - x1, y2 - y1), System.Drawing.Imaging.PixelFormat.Undefined);
            bitmap.Dispose();
            return b;
        }
        /// <summary>
        /// 批量截取图片
        /// </summary>
        public static List<Bitmap> EditImages(Bitmap p, List<TPoint> ps)
        {
            List<Bitmap> bs = new List<Bitmap>();
            for (int i = 0; i < ps.Count; i += 2)
            {
                bs.Add(p.Clone(new Rectangle((int)(ps[i].X * p.Width), (int)(ps[i].Y * p.Height), (int)((ps[i + 1].X - ps[i].X) * p.Width), (int)((ps[i + 1].Y - ps[i].Y) * p.Height)), System.Drawing.Imaging.PixelFormat.Undefined));
            }
            return bs;
            //Bitmap b = bitmap.Clone(new Rectangle(x1, y1, x2 - x1, y2 - y1), System.Drawing.Imaging.PixelFormat.Undefined);
        }
    }
    /// <summary>
    /// 
    /// </summary>
    public class TPoint
    {
        public TPoint(double x, double y)
        {
            X = x;
            Y = y;
        }

        /// <summary>
        /// X轴的比例
        /// </summary>
        public double X { get; }
        /// <summary>
        /// Y轴的比例
        /// </summary>
        public double Y { get; }

    }
}
