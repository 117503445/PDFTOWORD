using System;
using System.Collections.Generic;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Shapes;

namespace PDFTOWORD
{
    /// <summary>
    /// WdEdit.xaml 的交互逻辑
    /// </summary>
    public partial class WdEdit : Window
    {
        /// <summary>
        /// 要操作的bitmap
        /// </summary>
        private Bitmap bitmap = new Bitmap(@"C:\Users\117503445\Desktop\1.jpg");
        /// <summary>
        /// 维护点集
        /// </summary>
        private List<TPoint> points = new List<TPoint>();
        public WdEdit()
        {
            InitializeComponent();
        }

        private void Window_Loaded(object sender, RoutedEventArgs e)
        {
            Img.Source = new BitmapImage(new Uri(@"C:\Users\117503445\Desktop\1.jpg", UriKind.Absolute));//strImagePath 就绝对路径
        }
        private void Img_MouseDown(object sender, MouseButtonEventArgs e)
        {
            //double xScale = 0;
            //double yScale = 0;
            //xScale = Img.ActualWidth / bitmap.Width;
            //yScale = Img.ActualHeight / bitmap.Height;
            //Console.WriteLine(xScale);
            Line l = new Line();
            G.Children.Add(l);
            l.X1 = e.GetPosition(Img).X;
            l.X2 = e.GetPosition(Img).X + 100;
            l.Y1 = e.GetPosition(Img).Y;
            l.Y2 = e.GetPosition(Img).Y + 100;
            l.Stroke = System.Windows.Media.Brushes.LightSteelBlue;
            l.StrokeThickness = 10;
            Panel.SetZIndex(l, 2);

            points.Add(new TPoint(e.GetPosition(Img).X/Img.ActualWidth, e.GetPosition(Img).Y/Img.ActualHeight));
            Console.WriteLine(e.GetPosition(Img).X);
            Console.WriteLine(e.GetPosition(Img).Y);
            //if (i == 0)
            //{
            //    point = e.GetPosition(Img);
            //}
            //i += 1;
            //if (i == 2)
            //{
            //    var p1 = point;
            //    var p2 = e.GetPosition(Img);
            //    var b = ImageHelper.EditImage(@"C:\Users\117503445\Desktop\1.jpg", (int)(p1.X / xScale), (int)(p1.Y / yScale), (int)(p2.X / xScale), (int)(p2.Y / yScale));
            //    b.Save(@"C:\Users\117503445\Desktop\2.jpg");
            //}
        }

        /// <summary>
        /// 使ScrollViewer支持横向滚动
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void Scv_PreviewMouseWheel(object sender, MouseWheelEventArgs e)
        {
            ScrollViewer sv = sender as ScrollViewer;
            if (e.Delta > 0)
            {
                sv.LineLeft();
                sv.LineLeft();
            }
            else
            {
                sv.LineRight();
                sv.LineRight();
            }
            e.Handled = true;
        }

        private void BtnSave_Click(object sender, RoutedEventArgs e)
        {
            if (points.Count%2!=0)
            {
                Console.WriteLine("点不对应");
            }
            var bs = ImageHelper.EditImages(bitmap,points);
            for (int i = 0; i < bs.Count; i++)
            {
                bs[i].Save($@"C:\Users\117503445\Desktop\1\{i}.jpg");
            }
        }
    }

}
