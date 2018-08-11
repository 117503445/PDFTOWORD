using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Configuration;
using System.Data;
using System.IO;
using System.Linq;
using System.Threading.Tasks;
using System.Windows;
using TLib.Software;

namespace PDFTOWORD
{
    /// <summary>
    /// App.xaml 的交互逻辑
    /// </summary>
    public partial class App : Application
    {
        public static string Dir_APP { get; set; } = AppDomain.CurrentDomain.BaseDirectory;
        public static string Dir_File { get; set; } = Dir_APP + "File/";
        public static string Dir_File_PDF { get; set; } = Dir_File + "PDF/";
        public static string Dir_Desktop { get; set; } = Environment.GetFolderPath(Environment.SpecialFolder.Desktop)+"/";
        public static string dir_LstPaths = Dir_File + "LstPaths.xml";
        public static BindingList<string> LstPaths { get; set; }

        public static WdSetting WdSetting = new WdSetting();

        private void Application_Startup(object sender, StartupEventArgs e)
        {
            Directory.CreateDirectory(Dir_File);
            Directory.CreateDirectory(Dir_File_PDF);

            if (File.Exists(dir_LstPaths))
            {
                LstPaths = SerializeHelper.Load<BindingList<string>>(dir_LstPaths);
            }
            else
            {
                LstPaths = new BindingList<string>();
            }
#if !DEBUG
            WPF_ExpectionHandler.HandleExpection(Current, AppDomain.CurrentDomain);
            WPF_ExpectionHandler.ExpectionCatched += (s,arg) => { MessageBox.Show("翻车了,去找齐浩天,把Logger.xml发给他"); };
#endif
        }
        private void Application_Exit(object sender, ExitEventArgs e)
        {
            SerializeHelper.Save(LstPaths, dir_LstPaths);
        }
    }
}
