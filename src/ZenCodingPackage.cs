using System;
using System.Runtime.InteropServices;
using Microsoft.VisualStudio.Shell;

namespace ZenCodingVS
{
    [PackageRegistration(UseManagedResourcesOnly = true, AllowsBackgroundLoading = true)]
    [InstalledProductRegistration(Vsix.Name, Vsix.Description, Vsix.Version)]
    [Guid("7c1ab83e-2477-4554-a538-f3204dcd63dd")]
    public sealed class ZenCodingPackage : AsyncPackage
    {
    }
}
