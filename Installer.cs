﻿using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using Microsoft.VisualStudio.ExtensionManager;
using Microsoft.VisualStudio.Settings;
using Microsoft.Win32;
using System.Globalization;

namespace Root_VSIX
{
    public static class Installer
    {
        static int Main(string[] args)
        {
            if (args.Length != 2 && args.Length != 3)
            {
                PrintUsage();
                return 1;
            }

            string version;
            if (args.Length == 3)
            {
                version = args[0];
                args = args.Skip(1).ToArray();
            }
            else
            {
                version = FindVsVersions().LastOrDefault().ToString();
                if (string.IsNullOrEmpty(version))
                    return PrintError("Cannot find any installed copies of Visual Studio.");
            }

            string vsExe = GetVersionExe(version);
            if (string.IsNullOrEmpty(vsExe) && version.All(char.IsNumber))
            {
                version += ".0";
                vsExe = GetVersionExe(version);
            }

            if (string.IsNullOrEmpty(vsExe))
            {
                Console.Error.WriteLine("Cannot find Visual Studio " + version);
                PrintVsVersions();
                return 1;
            }

            if (!File.Exists(args[1]))
                return PrintError("Cannot find VSIX file " + args[1]);

            var vsix = ExtensionManagerService.CreateInstallableExtension(args[1]);

            Console.WriteLine("Installing " + vsix.Header.Name + " version " + vsix.Header.Version + " to Visual Studio " + version + " /RootSuffix " + args[0]);

            try
            {
                Install(vsExe, vsix, args[0]);
            }
            catch (Exception ex)
            {
                return PrintError("Error: " + ex.Message);
            }

            return 0;
        }

        public static IEnumerable<decimal?> FindVsVersions()
        {
            using (var software = Registry.LocalMachine.OpenSubKey("SOFTWARE"))
            using (var ms = software.OpenSubKey("Microsoft"))
            using (var vs = ms.OpenSubKey("VisualStudio"))
                return vs.GetSubKeyNames()
                        .Select(s =>
                {
                    decimal v;
                    if (!decimal.TryParse(s, NumberStyles.Number, CultureInfo.InvariantCulture, out v))
                        return new decimal?();
                    return v;
                })
                .Where(d => d.HasValue)
                .OrderBy(d => d);
        }

        public static string GetVersionExe(string version)
        {
            return Registry.GetValue(@"HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\VisualStudio\" + version + @"\Setup\VS", "EnvironmentPath", null) as string;
        }

        public static void Install(string vsExe, IInstallableExtension vsix, string rootSuffix)
        {
            using (var esm = ExternalSettingsManager.CreateForApplication(vsExe, rootSuffix))
            {
                var ems = new ExtensionManagerService(esm);
                IInstalledExtension installedVsix = null;
                if (ems.TryGetInstalledExtension(vsix.Header.Identifier, out installedVsix))
                {
                    Console.WriteLine($"Extension {vsix.Header.Name} version {vsix.Header.Version} already installed, unistalling first.");
                    ems.Uninstall(installedVsix);
                }

                ems.Install(vsix, perMachine: false);
            }
        }

        #region Output Messages
        private static int PrintError(string message)
        {
            Console.Error.WriteLine(message);
            return 1;
        }
        private static void PrintUsage()
        {
            Console.Error.WriteLine(typeof(Installer).Assembly.GetName().Name);
            Console.Error.WriteLine("Installs local VSIX extensions to custom Visual Studio RootSuffixes");
            Console.Error.WriteLine();
            Console.Error.WriteLine("Usage:");
            Console.Error.WriteLine();
            Console.Error.WriteLine("  " + typeof(Installer).Assembly.GetName().Name + " [<VS version>] <RootSuffix> <Path to VSIX>");
            Console.Error.WriteLine();
            Console.Error.WriteLine("The Visual Studio version must be specified as the internal version number (12.0 is 2013).");
            Console.Error.WriteLine("If omitted, the extension will be installed to the latest version of Visual Studio installed on the computer.");
        }
        private static void PrintVsVersions()
        {
            Console.Error.WriteLine("Detected versions:");
            Console.Error.WriteLine(string.Join(
                Environment.NewLine,
                FindVsVersions()
                    .Where(v => !string.IsNullOrEmpty(GetVersionExe(v.ToString())))
            ));
        }
        #endregion
    }
}
