using System;
using System.Collections;
using System.Collections.Generic;
using System.ComponentModel;
using System.Configuration.Install;
using System.Linq;
using System.ServiceProcess;
using System.Threading.Tasks;

namespace Exchange_Mail_Service
{
    [RunInstaller(true)]
    public partial class ProjectInstaller : System.Configuration.Install.Installer
    {
        ServiceProcessInstaller spi;
        ServiceInstaller si;
        public ProjectInstaller()
        {
            InitializeComponent();
            spi = new ServiceProcessInstaller();
            spi.Account = ServiceAccount.LocalSystem;

            si = new ServiceInstaller();
            si.StartType = ServiceStartMode.Manual;
            si.ServiceName = "Exchange_Mail_Service";
            si.DisplayName = "Exchange_Mail_Service";
            si.Description = "Service to get the attachment from mails through Exchange Server";

            this.Installers.Add(spi);
            this.Installers.Add(si);
        }

        private void serviceInstaller1_AfterInstall(object sender, InstallEventArgs e)
        {

        }

        private void serviceProcessInstaller1_AfterInstall(object sender, InstallEventArgs e)
        {

        }
    }
}
