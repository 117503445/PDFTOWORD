using System;
using System.Collections.Generic;
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
using System.Windows.Navigation;
using System.Windows.Shapes;
using System.IO;
using PDFTOWORD;
using System.Collections.ObjectModel;

namespace TpsEditor
{
    /// <summary>
    /// MainWindow.xaml 的交互逻辑
    /// </summary>
    public partial class MainWindow : Window
    {
        /// <summary>
        /// 维护点集
        /// </summary>
        public ObservableCollection<TPoint> Tps = new ObservableCollection<TPoint>();
        string File_tpstxt = AppDomain.CurrentDomain.BaseDirectory + "/tps.txt";
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
                if (sss.Length == 3)
                {
                    TPoint point = new TPoint(double.Parse(sss[0]), double.Parse(sss[1]), int.Parse(sss[2]));
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
                s += item.X + "," + item.Y + "," + item.PgIndex + ";";
            }
            File.WriteAllText(File_tpstxt, s, Encoding.UTF8);
        }

        public MainWindow()
        {
            InitializeComponent();
        }
        private void Window_Loaded(object sender, RoutedEventArgs e)
        {
            Tps = LoadTps();
            foreach (var item in Tps)
            {
                Lst.Items.Add($"{item.PgIndex}   {item.X}   {item.Y}");
            }
            for (int i = 0; i < Tps.Count; i += 2)
            {
                if (Tps[i].Y == Tps[i + 1].Y)
                {
                    Console.WriteLine($"{Tps[i].PgIndex}   {Tps[i].X}   {Tps[i].Y}");
                    Tps.RemoveAt(i + 1);
                    Tps.RemoveAt(i);
                    SaveTps();
                    break;
                }
            }
        }
    }
}
