using Microsoft.VisualStudio.Tools.Applications.Deployment;
using Microsoft.VisualStudio.Tools.Applications;
using System.IO;
using System;
using static System.Net.Mime.MediaTypeNames;

namespace FileCopyPDA
{
    public class FileCopyPDA : IAddInPostDeploymentAction
    {
        public void Execute(AddInPostDeploymentActionArgs args)
        {
            string dataDirectory = @"Assets\index.html";
            string file = @"index.html";
            string sourcePath = args.AddInPath;
            Uri deploymentManifestUri = args.ManifestLocation;
            string destPath = Environment.GetFolderPath(Environment.SpecialFolder.DesktopDirectory);
            string sourceFile = System.IO.Path.Combine(sourcePath, dataDirectory);
            string destFile = System.IO.Path.Combine(destPath, file);
            File.WriteAllText(System.IO.Path.Combine(destPath, @"sourcePath.txt"), sourcePath);

            switch (args.InstallationStatus)
            {
                case AddInInstallationStatus.InitialInstall:
                case AddInInstallationStatus.Update:
                    File.Copy(sourceFile, destFile);
                    break;
                case AddInInstallationStatus.Uninstall:
                    if (File.Exists(destFile))
                    {
                        File.Delete(destFile);
                    }
                    break;
            }
        }
    }
}