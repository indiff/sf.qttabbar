//    This file is part of QTTabBar, a shell extension for Microsoft
//    Windows Explorer.
//    Copyright (C) 2007-2010  Quizo, Paul Accisano
//
//    QTTabBar is free software: you can redistribute it and/or modify
//    it under the terms of the GNU General Public License as published by
//    the Free Software Foundation, either version 3 of the License, or
//    (at your option) any later version.
//
//    QTTabBar is distributed in the hope that it will be useful,
//    but WITHOUT ANY WARRANTY; without even the implied warranty of
//    MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
//    GNU General Public License for more details.
//
//    You should have received a copy of the GNU General Public License
//    along with QTTabBar.  If not, see <http://www.gnu.org/licenses/>.

using System.ComponentModel;
using System.Configuration.Install;
using System.ServiceProcess;

namespace QTTabBarService {
    internal class QTService : ServiceBase {

        private CommManager comm;

        private static void Main(string[] args) {
            Run(new QTService());
        }

        public QTService() {
            ServiceName = ServiceConst.REAL_NAME;
        }

        protected override void OnStart(string[] args) {
            if(comm != null) {
                comm.Dispose();
                comm = null;
            }
            comm = new CommManager();
        }

        protected override void OnStop() {
            if(comm != null) {
                comm.Dispose();
                comm = null;
            }
        }
    }

    [RunInstaller(true)]
    public class MyWindowsServiceInstaller : Installer {
        public MyWindowsServiceInstaller() {
            var processInstaller = new ServiceProcessInstaller();
            var serviceInstaller = new ServiceInstaller();

            //set the privileges
            processInstaller.Account = ServiceAccount.LocalSystem;

            serviceInstaller.DisplayName = ServiceConst.DISPLAY_NAME;
            serviceInstaller.StartType = ServiceStartMode.Automatic;

            //must be the same as what was set in Program's constructor
            serviceInstaller.ServiceName = ServiceConst.REAL_NAME;

            Installers.Add(processInstaller);
            Installers.Add(serviceInstaller);
        }
    }
}
