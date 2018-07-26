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
using TLib.Software;
using System.Collections.ObjectModel;

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

        private System.Drawing.Size bpSize;

        /// <summary>
        /// 维护点集
        /// </summary>
        public ObservableCollection<TPoint> Tps = new ObservableCollection<TPoint>();
        public string Dir_WorkPlace { get; set; }
        private string Dir_tpsTXT { get { return Dir_WorkPlace + "tps.txt"; } }
        public WdEdit(string dir_WorkPlace)
        {
            string file_img = dir_WorkPlace + "pdf.jpg";
            Dir_WorkPlace = dir_WorkPlace;
            InitializeComponent();
            bitmap = new Bitmap(file_img);
            bpSize = bitmap.Size;
            Img.Source = new BitmapImage(new Uri(file_img, UriKind.Absolute));//strImagePath 就绝对路径


        }

        private void Window_Loaded(object sender, RoutedEventArgs e)
        {

            var i = new List<string> { "Tps" };
            if (File.Exists(Dir_tpsTXT))
            {
                Tps = LoadTps();
            }
            Tps.CollectionChanged += (sd, arg) =>
            {
                SaveTps();
            };
            UpdateEllipse();
        }

        private ObservableCollection<TPoint> LoadTps()
        {
            ObservableCollection<TPoint> p = new ObservableCollection<TPoint>();
            if (!File.Exists(Dir_tpsTXT))
            {
                return p;
            }
            var s = File.ReadAllText(Dir_tpsTXT, Encoding.UTF8);
            string[] ss = s.Split(';');
            foreach (var item in ss)
            {
                var sss = item.Split(',');
                if (sss.Length == 2)
                {
                    TPoint point = new TPoint(double.Parse(sss[0]), double.Parse(sss[1]));
                    p.Add(point);
                }
            }
            return p;
        }

        private void SaveTps()
        {
            //File.Create(Dir_tpsTXT);
            string s = "";
            foreach (var item in Tps)
            {
                s += item.X + "," + item.Y + ";";
            }
            File.WriteAllText(Dir_tpsTXT, s, Encoding.UTF8);
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
                if (Tps.Count % 2 != 0)
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
                var bs = ImageHelper.EditImages(bitmap, Tps.ToList());
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
            Tps.Add(new TPoint(e.GetPosition(Img).X / Img.ActualWidth, e.GetPosition(Img).Y / Img.ActualHeight));
            UpdateEllipse();
        }
        private List<Ellipse> eps = new List<Ellipse>();
        private void UpdateEllipse()
        {
            for (int i = 0; i < eps.Count; i++)
            {
                G.Children.Remove(eps[i]);
            }
            for (int i = 0; i < Tps.Count; i++)
            {
                Ellipse ep = new Ellipse();
                ep.Visibility = Visibility.Visible;

                ep.Margin = new Thickness(Tps[i].X * Img.ActualWidth, Tps[i].Y * Img.ActualHeight, bpSize.Width - (Tps[i].X * Img.ActualWidth) + 30, Img.ActualHeight - (Tps[i].Y * Img.ActualHeight) - 5);
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
