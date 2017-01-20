using Microsoft.VisualStudio.Shell;
using System;
using System.Runtime.InteropServices;

namespace ZenCodingVS
{
    [PackageRegistration(UseManagedResourcesOnly = true)]
    [InstalledProductRegistration("#110", "#112", "1.0", IconResourceID = 400)]
    [Guid(ZenCodingPackage.PackageGuidString)]
    public sealed class ZenCodingPackage : Package
    {
        public const string PackageGuidString = "7c1ab83e-2477-4554-a538-f3204dcd63dd";

        protected override void Initialize()
        {
            base.Initialize();
        }
    }
}
