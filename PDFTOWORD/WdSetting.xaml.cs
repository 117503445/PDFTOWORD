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
using System.Windows.Shapes;

namespace PDFTOWORD
{
    /// <summary>
    /// WdSetting.xaml 的交互逻辑
    /// </summary>
    public partial class WdSetting : Window
    {
        public WdSetting()
        {
            InitializeComponent();
            //TLib.Software.Serializer serializer = new TLib.Software.Serializer(this, App.Dir_File + "Setting.xml", new List<string>() { "Sharpness" });
            TbAbout.Text += System.Reflection.Assembly.GetExecutingAssembly().GetName().Version.ToString();
        }
        /// <summary>
        /// 清晰度
        /// </summary>
        public int Sharpness { get; set; } = 150;

        private void Window_Loaded(object sender, RoutedEventArgs e)
        {

            TbSharpness.Text = Sharpness.ToString();
        }
        private void TbSharpness_TextChanged(object sender, TextChangedEventArgs e)
        {

            bool b = int.TryParse(TbSharpness.Text, out int r);
            if (b && r > 0)
            {
                Sharpness = r;
                TbInfoSharpness.Text = "清晰度设置(默认150,数值越大,越清晰,处理时间越慢)";
            }
            else
            {
                TbInfoSharpness.Text = "清晰度设置(默认150,数值越大,越清晰,处理时间越慢),但你需要输入正确的数值啊";
            }

        }

        private void Window_Closing(object sender, System.ComponentModel.CancelEventArgs e)
        {
            e.Cancel = true;
            Hide();
        }
    }
}
