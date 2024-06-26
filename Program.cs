using System;
using System.IO;
using Microsoft.Win32;
using Newtonsoft.Json;
using System.Collections.Generic;

namespace BrowserDataBackup
{
    class Program
    {
        static void Main(string[] args)
        {
            BackupEdgeFavorites();
            BackupChromeFavorites();
            BackupFirefoxFavorites();
            // BackupPasswords();
            // BackupWordOutlookTextbausteine();

            Console.WriteLine("Backup completed.");
        }

        static void BackupEdgeFavorites()
        {
            string edgeFavoritesPath = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),
                @"Microsoft\Edge\User Data\Default\Bookmarks");
            string backupPath = @"C:\Backup\Edge\Favorites.json";

            Directory.CreateDirectory(@"C:\Backup\Edge\");   

            if (File.Exists(edgeFavoritesPath))
            {
                File.Copy(edgeFavoritesPath, backupPath, true);
                Console.WriteLine("Edge favorites backed up.");
            }
            else
            {
                Console.WriteLine("Edge favorites not found.");
            }
        }

        static void BackupChromeFavorites()
        {
            string chromeFavoritesPath = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),
                @"Google\Chrome\User Data\Default\Bookmarks");
            string backupPath = @"C:\Backup\Chrome\Favorites.json";

            Directory.CreateDirectory(@"C:\Backup\Chrome\");

            if (File.Exists(chromeFavoritesPath))
            {
                File.Copy(chromeFavoritesPath, backupPath, true);
                Console.WriteLine("Chrome favorites backed up.");
            }
            else
            {
                Console.WriteLine("Chrome favorites not found.");
            }
        }

        static void BackupFirefoxFavorites()
        {
            string firefoxProfilesPath = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData),
                @"Mozilla\Firefox\Profiles");
            string[] directories = Directory.GetDirectories(firefoxProfilesPath, "*.default-release");

            if (directories.Length > 0)
            {
                string firefoxBookmarksPath = Path.Combine(directories[0], "places.sqlite");
                string backupPath = @"C:\Backup\Firefox\places.sqlite";

                Directory.CreateDirectory(@"C:\Backup\Firefox\");

                if (File.Exists(firefoxBookmarksPath))
                {
                    File.Copy(firefoxBookmarksPath, backupPath, true);
                    Console.WriteLine("Firefox favorites backed up.");
                }
                else
                {
                    Console.WriteLine("Firefox favorites not found.");
                }
            }
            else
            {
                Console.WriteLine("Firefox profile not found.");
            }
        }
    }
}