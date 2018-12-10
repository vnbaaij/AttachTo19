using System;
using System.Diagnostics.CodeAnalysis;
using System.Runtime.InteropServices;
using Microsoft.VisualStudio.Shell;

namespace AttachTo
{
    [PackageRegistration(UseManagedResourcesOnly = true, AllowsBackgroundLoading = true)]
    [InstalledProductRegistration("#110", "#112", "1.0", IconResourceID = 400)] // Info on this package for Help/About
    [ProvideMenuResource("Menus.ctmenu", 1)]
    [Guid(PackageGuidString)]
    [SuppressMessage("StyleCop.CSharp.DocumentationRules", "SA1650:ElementDocumentationMustBeSpelledCorrectly", Justification = "pkgdef, VS and vsixmanifest are valid VS terms")]
    [ProvideOptionPage(typeof(GeneralOptionsPage), "AttachTo", "General", 110, 120, false)]
    public sealed class AttachToPackage : AsyncPackage
    {
        /// <summary>
        /// AttachToPackage GUID string.
        /// </summary>
        public const string PackageGuidString = "4324de84-99e1-4732-bd0c-8af14184c59c";

        /// <summary>
        /// Initializes a new instance of the <see cref="AttachTo"/> class.
        /// </summary>
        public AttachToPackage()
        {
        }
    }
}
