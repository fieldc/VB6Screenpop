using System;
using System.Collections;
using System.Collections.Generic;
using System.ComponentModel;
using System.Configuration.Install;
using System.Runtime.InteropServices;


namespace SSFIntegration
{
    [RunInstaller(true)]
    public partial class ComInstall : System.Configuration.Install.Installer
    {
        public ComInstall()
        {
            InitializeComponent();
        }

        public override void Install(IDictionary stateSaver)
        {
            base.Install(stateSaver);
            RegistrationServices regsrv = new RegistrationServices();
            if (!regsrv.RegisterAssembly(GetType().Assembly, AssemblyRegistrationFlags.None))
            {
                throw new InstallException("Failed to register for COM Interop.");
            }
        }

        public override void Uninstall(IDictionary savedState)
        {
            base.Uninstall(savedState);
            RegistrationServices regsrv = new RegistrationServices();
            if (!regsrv.UnregisterAssembly(GetType().Assembly))
            {
                throw new InstallException("Failed to unregister for COM Interop.");
            }
        }

    }
}
