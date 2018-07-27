﻿using System;
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
using TLib.UI.WPF_MessageBox;
using Rec = System.Windows.Shapes.Rectangle;
namespace PDFTOWORD
{
    /// <summary>
    /// WdEdit.xaml 的交互逻辑
    /// </summary>
    public partial class WdEdit : Window
    {


        /// <summary>
        /// 维护点集
        /// </summary>
        public ObservableCollection<TPoint> Tps = new ObservableCollection<TPoint>();
        private ObservableCollection<TPoint> LoadTps()
        {
            ObservableCollection<TPoint> p = new ObservableCollection<TPoint>();
            if (!File.Exists(File_tpstxt))
            {
                return p;
            }
            var s = File.ReadAllText(File_tpstxt, Encoding.UTF8);
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
            string s = "";
            foreach (var item in Tps)
            {
                s += item.X + "," + item.Y + ";";
            }
            File.WriteAllText(File_tpstxt, s, Encoding.UTF8);
        }
        public string Dir_WorkPlace { get; set; }
        private string File_tpstxt { get { return Dir_WorkPlace + "tps.txt"; } }

        public WdEdit(string dir_WorkPlace)
        {
            InitializeComponent();

            Dir_WorkPlace = dir_WorkPlace;

        }
        int maxPgIndex = 0;
        private void Window_Loaded(object sender, RoutedEventArgs e)
        {
            DirectoryInfo info = new DirectoryInfo(Dir_WorkPlace + "temp");
            var f = info.GetFiles();

            maxPgIndex = (from x in f orderby int.Parse(x.Name.Substring(0, x.Name.Length - 4)) select int.Parse(x.Name.Substring(0, x.Name.Length - 4))).ToList().Last();//不然11.jpg在2.jpg前面


            if (File.Exists(File_tpstxt))
            {
                Tps = LoadTps();
            }
            Tps.CollectionChanged += (sd, arg) =>
            {
                SaveTps();
            };
            UpdateEpsAndRs();
            Console.WriteLine(2);
            UpdateImg();
        }
        int PgIndex { get { return pgIndex; } set { pgIndex = value; UpdateImg(); } }
        int pgIndex;
        private void UpdateImg()
        {
            Img.Source = null;
            Img.Source = GetImage(Dir_WorkPlace + $"temp/{PgIndex}.jpg");
            //Img1.Source = new BitmapImage(new Uri(Dir_WorkPlace + $"temp/{PgIndex + 1}.jpg", UriKind.Absolute));
        }
        public static BitmapImage GetImage(string imagePath)
        {
            BitmapImage bitmap = new BitmapImage();

            if (File.Exists(imagePath))
            {
                bitmap.BeginInit();
                bitmap.CacheOption = BitmapCacheOption.OnLoad;

                using (Stream ms = new MemoryStream(File.ReadAllBytes(imagePath)))
                {
                    bitmap.StreamSource = ms;
                    bitmap.EndInit();
                    bitmap.Freeze();
                }
            }

            return bitmap;
        }

        private void Img_MouseDown(object sender, MouseButtonEventArgs e)
        {
            if (e.RightButton == MouseButtonState.Released)//左键
            {
                double x = e.GetPosition(Img).X / Img.ActualWidth;
                double y = e.GetPosition(Img).Y / Img.ActualHeight;

                bool isLegal = true;
                if (Tps.Count > 0)
                {
                    isLegal = x > Tps.Last().X && y > Tps.Last().Y;
                }

                if (Tps.Count % 2 == 0 || (isLegal))
                {
                    Tps.Add(new TPoint(x, y,pgIndex));
                }
            }
            else
            {//右键
                RemoveAPoint();
            }

            UpdateEpsAndRs();

        }
        private async void BtnSave_Click(object sender, RoutedEventArgs e)
        {
            if (Tps.Count % 2 != 0)
            {
                await WdMessageBox.Display("消息", "请选中偶数个点", "", "", "知道了");
                return;
            }
            await Task.Run(() =>
            {

                TbInfo.Dispatcher.Invoke(() =>
                {
                    TbInfo.Text = "切割图片中";
                });
                BtnSave.Dispatcher.Invoke(() =>
                {
                    BtnSave.IsEnabled = false;
                });
                var bs = ImageHelper.EditImages(null, Tps.ToList());//!!!!!!!!!!!!!!!!!!!!!!!!

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
                WordHelper.WriteWord($@"{App.Dir_Desktop}{TimeStamp.Now}.doc", dir_pic);
                TbInfo.Dispatcher.Invoke(() =>
                {
                    TbInfo.Text = "运行中";
                });
            });
            BtnSave.IsEnabled = true;
        }
        private async void BtnDelete_Click(object sender, RoutedEventArgs e)
        {
            var i = await WdMessageBox.Display("消息", "真的要删除所有点吗?不过会留下备份的.", "", "取消", "删除");
            if (i == 2)
            {
                File.Copy(File_tpstxt, File_tpstxt + TimeStamp.Now);
                Tps = new ObservableCollection<TPoint>();
                SaveTps();
                Tps.CollectionChanged += (sd, arg) =>
                {
                    SaveTps();
                };
                UpdateEpsAndRs();
            }
        }
        private void BtnUndo_Click(object sender, RoutedEventArgs e)
        {
            RemoveAPoint();
        }

        private void G_SizeChanged(object sender, SizeChangedEventArgs e)
        {
            UpdateEpsAndRs();
        }
        private void Window_Closed(object sender, EventArgs e)
        {

        }

        private void RemoveAPoint()
        {
            if (Tps.Count > 0)
            {
                Tps.RemoveAt(Tps.Count - 1);
                UpdateEpsAndRs();
            }
            else
            {
                WdMessageBox.Display("消息", "无法再撤销了");
            }
        }
        private List<Ellipse> eps = new List<Ellipse>();
        private List<Rec> rs = new List<Rec>();
        private void UpdateEpsAndRs()
        {
            for (int i = 0; i < eps.Count; i++)
            {
                //G.Children.Remove(eps[i]);
            }
            for (int i = 0; i < rs.Count; i++)
            {
                //G.Children.Remove(rs[i]);
            }
            for (int i = 0; i < Tps.Count; i += 2)
            {
                //Ellipse ep = new Ellipse
                //{
                //    Visibility = Visibility.Visible,
                //    Margin = new Thickness(Tps[i].X * Img.ActualWidth, Tps[i].Y * Img.ActualHeight, bpSize.Width - (Tps[i].X * Img.ActualWidth) + 30, Img.ActualHeight - (Tps[i].Y * Img.ActualHeight) - 5),
                //    //Img.ActualHeight- e.GetPosition(Img).Y+2
                //    //ep.Width = 200;
                //    //ep.Height = 200;
                //    Stroke = new SolidColorBrush(Colors.Red),
                //    StrokeThickness = 3,
                //    Fill = new SolidColorBrush(Colors.Red)
                //};
                //G.Children.Add(ep);
                //Panel.SetZIndex(ep, 2);
                //eps.Add(ep);
            }

            //for (int i = 0; i < Tps.Count; i += 2)
            //{
            //    if (i + 1 < Tps.Count)
            //    {
            //        Rec r = new Rec
            //        {
            //            Fill = new SolidColorBrush(System.Windows.Media.Color.FromArgb(0, 0, 0, 0)),
            //            HorizontalAlignment = HorizontalAlignment.Left,
            //            VerticalAlignment = VerticalAlignment.Top,
            //            Stroke = new SolidColorBrush(Colors.Red),
            //            StrokeThickness = 3,
            //            Margin = new Thickness(Tps[i].X * Img.ActualWidth, Tps[i].Y * Img.ActualHeight, 0, 0),
            //            Width = (Tps[i + 1].X - Tps[i].X) * Img.ActualWidth,
            //            Height = (Tps[i + 1].Y - Tps[i].Y) * Img.ActualHeight
            //        };
            //        G.Children.Add(r);
            //        Panel.SetZIndex(r, 2);
            //        rs.Add(r);

            //    }
            //}
        }

        private void GMain_MouseWheel(object sender, MouseWheelEventArgs e)
        {

        }

        private void Window_MouseWheel(object sender, MouseWheelEventArgs e)
        {

            if (e.Delta > 0)
            {
                if (PgIndex > 0)
                {

                    PgIndex -= 1;
                }
            }
            else
            {
                if (pgIndex < maxPgIndex)
                {

                    PgIndex += 1;
                }
            }
        }

    }
}


