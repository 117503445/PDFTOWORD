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
using System.IO;

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
        private Bitmap bitmap;
        /// <summary>
        /// 维护点集
        /// </summary>
        private List<TPoint> points = new List<TPoint>();
        public string Dir_WorkPlace { get; set; }
        public WdEdit(string dir_WorkPlace)
        {
            string file_img = dir_WorkPlace + "pdf.jpg";
            Dir_WorkPlace = dir_WorkPlace;
            InitializeComponent();
            bitmap = new Bitmap(file_img);
            Img.Source = new BitmapImage(new Uri(file_img, UriKind.Absolute));//strImagePath 就绝对路径
        }

        private void Window_Loaded(object sender, RoutedEventArgs e)
        {
            
        }
        private void Img_MouseDown(object sender, MouseButtonEventArgs e)
        {
            Line l = new Line();
            G.Children.Add(l);
            l.X1 = e.GetPosition(Img).X;
            l.X2 = e.GetPosition(Img).X + 100;
            l.Y1 = e.GetPosition(Img).Y;
            l.Y2 = e.GetPosition(Img).Y + 100;
            l.Stroke = System.Windows.Media.Brushes.LightSteelBlue;
            l.StrokeThickness = 10;
            Panel.SetZIndex(l, 2);
            points.Add(new TPoint(e.GetPosition(Img).X / Img.ActualWidth, e.GetPosition(Img).Y / Img.ActualHeight));
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

        private async void BtnSave_Click(object sender, RoutedEventArgs e)
        {
            await Task.Run(() =>
            {
                if (points.Count % 2 != 0)
                {
                    Console.WriteLine("点不对应");
                    return;
                }
                TbInfo.Dispatcher.Invoke(() =>
                {
                    TbInfo.Text = "切割图片中";
                });
                var bs = ImageHelper.EditImages(bitmap, points);
                string dir_pic = Dir_WorkPlace+@"pics\";
                Directory.CreateDirectory(dir_pic);
                for (int i = 0; i < bs.Count; i++)
                {
                    bs[i].Save($@"{dir_pic}{i}.jpg");
                }
                TbInfo.Dispatcher.Invoke(() =>
                {
                    TbInfo.Text = "生成DOC中";
                });
                WordHelper.WriteWord($@"{App.Dir_Desktop}1.doc", dir_pic);
                TbInfo.Dispatcher.Invoke(() =>
                {
                    TbInfo.Text = "运行中";
                });
            });

        }

        private void BtnUndo_Click(object sender, RoutedEventArgs e)
        {

        }




    }

}
