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
using Microsoft.WindowsAPICodePack.Dialogs;
using O2S.Components.PDFRender4NET;
using System.Drawing.Imaging;
using MSWord = Microsoft.Office.Interop.Word;
using System.Reflection;
using Spire.Pdf;

namespace PDFTOWORD
{
    /// <summary>
    /// MainWindow.xaml 的交互逻辑
    /// </summary>
    public partial class MainWindow : Window
    {
        public MainWindow()
        {
            InitializeComponent();
        }

        private async void BtnAddress_Click(object sender, RoutedEventArgs e)
        {
            var dialog = new CommonOpenFileDialog
            {
                IsFolderPicker = true
            };
            CommonFileDialogResult result = dialog.ShowDialog();
            if (result != CommonFileDialogResult.Ok)
            {
                return;
            }
            await Task.Run(() =>
            {

                //Console.WriteLine(dialog.FileName);

                string dir_Desktop = Environment.GetFolderPath(Environment.SpecialFolder.Desktop) + "/";
                string dir_result = dir_Desktop + "PDFresult/";
                Directory.CreateDirectory(dir_result);

                DirectoryInfo info = new DirectoryInfo(dialog.FileName);
                foreach (var item in info.GetFiles("*.pdf", SearchOption.AllDirectories))
                {
                    TbInfo.Dispatcher.Invoke(()=> {
                        TbInfo.Text += "Processing " + item.FullName + Environment.NewLine;
                    });
                    Console.WriteLine();
                    string dir_resultDOC = dir_result + item.Name + ".doc";
                    if (File.Exists(dir_resultDOC))
                    {
                        TbInfo.Dispatcher.Invoke(() => {
                            TbInfo.Text += "已存在" + Environment.NewLine;
                        });
                    }
                    else
                    {
                        PdfDocument doc = new PdfDocument();
                        doc.LoadFromFile(item.FullName);
                        doc.SaveToFile(dir_resultDOC, FileFormat.DOC);
                        doc.Dispose();
                        GC.Collect();
                    }

                }
                TbInfo.Dispatcher.Invoke(() => {
                    TbInfo.Text += "全部结束了." + Environment.NewLine;
                });
            });

        }

    }
}
