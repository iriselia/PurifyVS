namespace FrenchKiwi.PurifyVS
{
    using System;
    
    /// <summary>
    /// Helper class that exposes all GUIDs used across VS Package.
    /// </summary>
    internal sealed partial class PackageGuids
    {
        public const string guidPurifyVSPkgString = "27dd9dea-6dd2-403e-929d-3ff20d896c5e";
        public const string guidPurifyVSCmdSetString = "32af8a17-bbbc-4c56-877e-fc6c6575a8cf";
        public static Guid guidPurifyVSPkg = new Guid(guidPurifyVSPkgString);
        public static Guid guidPurifyVSCmdSet = new Guid(guidPurifyVSCmdSetString);
    }
    /// <summary>
    /// Helper class that encapsulates all CommandIDs uses across VS Package.
    /// </summary>
    internal sealed partial class PackageIds
    {
        public const int cmdidMyCommand = 0x0100;
    }
}
