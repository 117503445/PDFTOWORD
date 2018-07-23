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
    /// WdEdit.xaml 的交互逻辑
    /// </summary>
    public partial class WdEdit : Window
    {
        public WdEdit()
        {
            InitializeComponent();
        }

        private void Window_Loaded(object sender, RoutedEventArgs e)
        {
           
            Img.Source= new BitmapImage(new Uri(@"C:\Users\117503445\Desktop\1.jpg", UriKind.Absolute));//strImagePath 就绝对路径
        }
    }
}
