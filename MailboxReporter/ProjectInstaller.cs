﻿using System.ComponentModel;
using System.Configuration.Install;

// ReSharper disable ClassNeverInstantiated.Global

namespace MailboxReporter
{
    [RunInstaller(true)]
    public partial class ProjectInstaller : Installer
    {
        public ProjectInstaller()
        {
            InitializeComponent();
        }
    }
}