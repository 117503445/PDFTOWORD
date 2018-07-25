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
using System.Windows.Threading;

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

            DispatcherTimer timer = new DispatcherTimer()
            {
                IsEnabled = true,
                Interval = TimeSpan.FromMilliseconds(20)
            };
            timer.Tick += (s, e) =>
            {
                if (Mouse.GetPosition(GMain).X >= 50)
                {
                    l1.Visibility = Visibility.Visible;
                    l2.Visibility = Visibility.Visible;
                    l1.X1 = 0;
                    l1.Y1 = Mouse.GetPosition(GMain).Y;
                    l1.X2 = bitmap.Width;
                    l1.Y2 = Mouse.GetPosition(GMain).Y;

                    l2.X1 = Mouse.GetPosition(GMain).X - 50;
                    l2.Y1 = 0;
                    l2.X2 = Mouse.GetPosition(GMain).X - 50;
                    l2.Y2 = bitmap.Height;
                }
                else
                {
                    l1.Visibility = Visibility.Hidden;
                    l2.Visibility = Visibility.Hidden;
                }
            };

        }

        private void Window_Loaded(object sender, RoutedEventArgs e)
        {
            G.Children.Add(l1);
            G.Children.Add(l2);
            Panel.SetZIndex(l1, 3);
            Panel.SetZIndex(l2, 3);
        }
        private void Img_MouseDown(object sender, MouseButtonEventArgs e)
        {

        }

        /// <summary>
        /// 使ScrollViewer支持横向滚动
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void Scv_PreviewMouseWheel(object sender, MouseWheelEventArgs e)
        {
            ScrollViewer sv = sender as ScrollViewer;
            for (int i = 0; i < 6; i++)
            {
                if (e.Delta > 0)
                {
                    sv.LineLeft();
                }
                else
                {
                    sv.LineRight();
                }

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
                BtnSave.Dispatcher.Invoke(() =>
                {
                    BtnSave.IsEnabled = false;
                });
                var bs = ImageHelper.EditImages(bitmap, points);
                string dir_pic = Dir_WorkPlace + @"pics\";
                if (Directory.Exists(dir_pic))
                {
                    Directory.Delete(dir_pic, true);
                }
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
            BtnSave.IsEnabled = true;
        }

        private Line l1 = new Line()
        {
            Visibility = Visibility.Collapsed,
            Stroke = System.Windows.Media.Brushes.Blue,
            StrokeThickness = 2,
            Opacity = 0.3
        };//鼠标指示,横
        private Line l2 = new Line()
        {
            Visibility = Visibility.Collapsed,
            Stroke = System.Windows.Media.Brushes.Blue,
            StrokeThickness = 2,
            Opacity = 0.3
        };//鼠标指示,竖
        private void BtnUndo_Click(object sender, RoutedEventArgs e)
        {

        }

        private void G_MouseDown(object sender, MouseButtonEventArgs e)
        {
            //Console.WriteLine("!");
            //Console.WriteLine(e.GetPosition(Img).X);
            //Console.WriteLine(e.GetPosition(Img).Y);
            //Console.WriteLine(bitmap.Width - e.GetPosition(Img).X + 20);
            //Console.WriteLine(Img.ActualHeight - e.GetPosition(Img).Y - 2);
            //Console.WriteLine("!!");
            //Console.WriteLine(Img.ActualHeight);
            //Console.WriteLine("!");


            //Line l = new Line();
            //G.Children.Add(l);
            //l.X1 = e.GetPosition(Img).X;
            //l.X2 = e.GetPosition(Img).X ;
            //l.Y1 = e.GetPosition(Img).Y;
            //l.Y2 = e.GetPosition(Img).Y ;
            //l.Stroke = System.Windows.Media.Brushes.LightSteelBlue;
            //l.StrokeThickness = 10;
            //Panel.SetZIndex(l, 2);
            points.Add(new TPoint(e.GetPosition(Img).X / Img.ActualWidth, e.GetPosition(Img).Y / Img.ActualHeight));
            UpdateEllipse();
        }
        private List<Ellipse> eps = new List<Ellipse>();
        private void UpdateEllipse()
        {
            for (int i = 0; i < eps.Count; i++)
            {
                G.Children.Remove(eps[i]);
            }
            for (int i = 0; i < points.Count; i++)
            {
                Ellipse ep = new Ellipse();
                ep.Visibility = Visibility.Visible;

                ep.Margin = new Thickness(points[i].X*Img.ActualWidth, points[i].Y * Img.ActualHeight, bitmap.Width - (points[i].X * Img.ActualWidth) + 30, Img.ActualHeight - (points[i].Y * Img.ActualHeight) - 5);
                //Img.ActualHeight- e.GetPosition(Img).Y+2
                //ep.Width = 200;
                //ep.Height = 200;
                ep.Stroke = new SolidColorBrush(Colors.Red);
                ep.StrokeThickness = 3;
                ep.Fill = new SolidColorBrush(Colors.Red);
                G.Children.Add(ep);
                Panel.SetZIndex(ep, 2);
                eps.Add(ep);
            }
        }

        private void G_SizeChanged(object sender, SizeChangedEventArgs e)
        {
            UpdateEllipse();
        }
    }

}
