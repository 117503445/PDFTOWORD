using System;
using System.Collections.Generic;
using System.IO;
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

namespace PDFTOWORD
{
    /// <summary>
    /// WdMain.xaml 的交互逻辑
    /// </summary>
    public partial class WdMain : Window
    {
        public WdMain()
        {
            InitializeComponent();
        }

        private void BtnExplorer_Click(object sender, RoutedEventArgs e)
        {
            var openFileDialog = new Microsoft.Win32.OpenFileDialog()
            {
                Filter = "Portable Document Format|*.pdf"
            };
            var result = openFileDialog.ShowDialog();
            if (result == true)
            {
                if (!App.LstPaths.Contains(openFileDialog.FileName))
                {
                    App.LstPaths.Add(openFileDialog.FileName);
                }
                OpenWdProcess(openFileDialog.FileName);
            }
        }
        private void LstPPT_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            OpenWdProcess((string)LstPPT.SelectedItem);
        }

        private void Window_Loaded(object sender, RoutedEventArgs e)
        {
            for (int i = App.LstPaths.Count - 1; i >= 0; i--)
            {
                if (!File.Exists(App.LstPaths[i]))
                {
                    App.LstPaths.RemoveAt(i);
                }
            }
            LstPPT.ItemsSource = App.LstPaths;
        }

        public void OpenWdProcess(string dir_pdf)
        {
            var index = App.LstPaths.IndexOf(dir_pdf);
            var temp = App.LstPaths[0];
            App.LstPaths[0] = App.LstPaths[index];
            App.LstPaths[index] = temp;

            WdProcess wdProcess = new WdProcess(dir_pdf);
            wdProcess.Show();
        }

        private void BtnSetting_Click(object sender, RoutedEventArgs e)
        {

                App.WdSetting.Show();

        }

        private void Window_Closed(object sender, EventArgs e)
        {
            App.Current.Shutdown();
        }
    }
}
