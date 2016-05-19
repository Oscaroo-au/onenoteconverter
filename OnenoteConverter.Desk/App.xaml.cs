using System;
using System.Collections.Generic;
using System.Configuration;
using System.Data;
using System.IO;
using System.Linq;
using System.Threading.Tasks;
using System.Windows;
using log4net;

namespace OnenoteConverter
{
    /// <summary>
    /// Interaction logic for App.xaml
    /// </summary>
    public partial class App : Application
    {
        protected override void OnStartup(StartupEventArgs e)
        {
            base.OnStartup(e);

            var fi = new FileInfo("logger_config.xml");

            log4net.Config.XmlConfigurator.ConfigureAndWatch(fi);
        }

    }
}
