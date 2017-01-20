using Microsoft.VisualStudio.Shell;
using System;
using System.Runtime.InteropServices;

namespace ZenCodingVS
{
    [PackageRegistration(UseManagedResourcesOnly = true)]
    [InstalledProductRegistration("#110", "#112", Vsix.Version, IconResourceID = 400)]
    [Guid("7c1ab83e-2477-4554-a538-f3204dcd63dd")]
    public sealed class ZenCodingPackage : Package
    {
    }
}
