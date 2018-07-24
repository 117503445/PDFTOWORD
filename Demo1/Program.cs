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
                if (imgs[i].Height != imgs[0].Height)
                {
                    throw new Exception("高度不同!");
                }
            }
            Bitmap bitmap = new Bitmap(totalWidth, imgs[0].Height);
            int twidth = 0;
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

        static void Main(string[] args)
        {

            Console.WriteLine("OJBK");
            Console.ReadLine();
        }
    }
}
