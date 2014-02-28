using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace CopyShortcuts
{
    class Program
    {
        static void Main(string[] args)
        {
            // Define the source directory to copy links from
            var sourceDirectoryPath = "c:\\brads\\pictures\\wallpaper\\comics\\";
            var backupDirectoryPath = sourceDirectoryPath + "backup-links\\";
            var sourceDirectory = new DirectoryInfo(sourceDirectoryPath);

            // Make sure we have a backup folder
            Console.WriteLine("Checking for backup folder...");
            var backupLinksFolder = new DirectoryInfo(backupDirectoryPath);
            if (!backupLinksFolder.Exists)
            {
                // Backup Folder was not found, creating...
                Console.WriteLine("--> Backup folder not found.");
                Console.WriteLine(@"Creating backup folder...", backupDirectoryPath);
                backupLinksFolder.Create();
                Console.WriteLine("Created backup folder.");
            }
            else
            {
                Console.WriteLine("Found backup folder {0}", backupLinksFolder.FullName);
            }

            foreach (var shortCut in sourceDirectory.GetFiles("*.lnk"))
            {
                Console.WriteLine(@"Processing shortcut: ""{0}""", shortCut.Name);
                var targetPath = GetShortcutTargetFile(shortCut.FullName);

                // Check if the target exists
                Console.WriteLine(@"Locating target file...");
                var targetFile = new FileInfo(targetPath);
                if (targetFile.Exists)
                {
                    Console.WriteLine(@"Located target file: ""{0}"".", targetPath);
                }
                else
                {
                    Console.WriteLine(@"--> ERROR: Target file not found: ""{0}"".", targetPath);
                    continue;
                }

                // Copy target file to source directory
                Console.WriteLine("Copying target of link to source directory...");
                targetFile.CopyTo(sourceDirectory.FullName + targetFile.Name);
                Console.WriteLine("Copied sucessfully.");

                // Move the link to backup folder
                Console.WriteLine("Moving link to backup folder...");
                shortCut.MoveTo(backupLinksFolder.FullName + targetFile.Name);
                Console.WriteLine("Moved link to backup folder.");                
            }            
        }

        static string GetShortcutTargetFile(string shortcutFilename)
        {
            string pathOnly = System.IO.Path.GetDirectoryName(shortcutFilename);
            string filenameOnly = System.IO.Path.GetFileName(shortcutFilename);

            //
            ///dynamic shell = new Shell32.Shell();
            Type t = Type.GetTypeFromProgID("Shell.Application");
            dynamic shell = Activator.CreateInstance(t);

            dynamic folder = shell.NameSpace(pathOnly);
            dynamic folderItem = folder.ParseName(filenameOnly);
            if (folderItem != null)
            {
                return ((Shell32.ShellLinkObject)folderItem.GetLink).Path;
            }

            return ""; // not found, use if (File.Exists()) in the calling code
            // or remove the return and throw an exception here if you like
        }

    }
}
