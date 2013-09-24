// --------------------------------------------------------------------------
// <copyright file="ProjectInstaller.cs" company="Nexcel Solutions Vietnam">
//     Copyright (c) Nexcel Solutions Vietnam. All rights reserved.
// </copyright>
// --------------------------------------------------------------------------

namespace WorkbookServiceHost
{
    using System.ComponentModel;

    /// <summary>
    /// Class ProjectInstaller.
    /// </summary>
    [RunInstaller(true)]
    public partial class ProjectInstaller : System.Configuration.Install.Installer
    {
        /// <summary>
        /// Initializes a new instance of the <see cref="ProjectInstaller"/> class.
        /// </summary>
        public ProjectInstaller()
        {
            InitializeComponent();
        }
    }
}
