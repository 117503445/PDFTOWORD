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
        /// 截取图片
        /// </summary>
        /// <param name="file_pic"></param>
        /// <param name="file_save"></param>
        /// <param name="x1"></param>
        /// <param name="y1"></param>
        /// <param name="x2"></param>
        /// <param name="y2"></param>
        public static Bitmap EditImage(string file_pic,int x1,int y1,int x2,int y2) {
            Console.WriteLine(x1);
            Console.WriteLine(y1);

            Console.WriteLine(x2);
            Console.WriteLine(y2);
            Bitmap bitmap = new Bitmap(file_pic);
            Bitmap b = bitmap.Clone(new Rectangle(x1, y1, x2 - x1, y2 - y1), System.Drawing.Imaging.PixelFormat.Undefined);
            bitmap.Dispose();
            return b;
        }
    }
}
