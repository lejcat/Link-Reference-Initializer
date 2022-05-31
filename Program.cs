using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace LinkReferenceInit
{
    class Program
    {
        static string LogPath { get; set; }

        [STAThread]
        static void Main(string[] args)
        {
            Console.WriteLine("<Link Reference Initializer>");

            LogPath = System.Reflection.Assembly.GetExecutingAssembly().GetName().Name;
            LogPath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, LogPath);
            LogPath = Path.ChangeExtension(LogPath, ".log");

            //ChangeLinkTarget(@"C:\Users\Finaldata\Desktop\DFOCS.lnk", "NewPath");

            string[] curLinks = Directory.GetFiles(AppDomain.CurrentDomain.BaseDirectory, "*.lnk", SearchOption.TopDirectoryOnly);
            string[] curFolders = Directory.GetDirectories(AppDomain.CurrentDomain.BaseDirectory, "*", SearchOption.TopDirectoryOnly);

            foreach (string lnk in curLinks)
            {
                string linkPath = GetShortcutLinkPath(lnk);
                if (string.IsNullOrEmpty(linkPath) == false)
                {
                    linkPath = linkPath.Replace("/", "\\");

                    foreach (string folder in curFolders)
                    {
                        string folderName = Path.GetFileName(folder);
                        if (string.IsNullOrEmpty(folderName) == false)
                        {
                            folderName = "\\" + folderName + "\\";
                            int index = linkPath.IndexOf(folderName);
                            if (index > -1)
                            {
                                linkPath = linkPath.Remove(0, index + 1);
                                linkPath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, linkPath);

                                if (Directory.Exists(linkPath) || File.Exists(linkPath))
                                {
                                    ChangeLinkToCurrent(lnk, linkPath);

                                    if (Path.GetFileName(lnk).StartsWith("[Invalid] "))
                                    {
                                        string filepath = Path.GetFileName(lnk).Remove(0, 10);
                                        filepath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, filepath);

                                        File.Move(lnk, filepath);
                                    }
                                }
                                else if(File.Exists(lnk))
                                {
                                    Console.WriteLine("Link not found : " + lnk);
                                    SaveErrorLog("Link not found > " + lnk);

                                    string filepath = "[Invalid] " + Path.GetFileName(lnk);
                                    filepath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, filepath);

                                    File.Move(lnk, filepath);
                                }

                                break;
                            }
                        }
                    }
                }
            }

            Console.WriteLine("***** Initialize Ended *****");

            Console.ReadKey();
        }

        static void ChangeLinkToCurrent(string shortcutFullPath, string newTarget)
        {
            // Load the shortcut.
            Shell32.Shell shell = new Shell32.Shell();
            Shell32.Folder folder = shell.NameSpace(Path.GetDirectoryName(shortcutFullPath));
            Shell32.FolderItem folderItem = folder.Items().Item(Path.GetFileName(shortcutFullPath));
            Shell32.ShellLinkObject currentLink = (Shell32.ShellLinkObject)folderItem.GetLink;

            // Assign the new path here. This value is not read-only.
            currentLink.Path = newTarget;

            // Save the link to commit the changes.
            currentLink.Save();
        }
        static string GetShortcutLinkPath(string shortcutFullPath)
        {
            string result = string.Empty;

            try
            {
                // Load the shortcut.
                Shell32.Shell shell = new Shell32.Shell();
                Shell32.Folder folder = shell.NameSpace(Path.GetDirectoryName(shortcutFullPath));
                Shell32.FolderItem folderItem = folder.Items().Item(Path.GetFileName(shortcutFullPath));
                Shell32.ShellLinkObject currentLink = (Shell32.ShellLinkObject)folderItem.GetLink;

                result = currentLink.Path;
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.Message);
                SaveErrorLog(ex.Message);
            }

            return result;
        }

        static void SaveErrorLog(string message)
        {
            if (!string.IsNullOrEmpty(LogPath))
            {
                try
                {
                    string text = DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss") + " : " + message.Replace("\n", " ");
                    using (StreamWriter logWriter = File.AppendText(LogPath))
                    {
                        logWriter.WriteLine(text);
                    }
                }
                catch
                { }
            }
        }
    }
}
