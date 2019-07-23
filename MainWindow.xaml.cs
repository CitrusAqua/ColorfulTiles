using System;
using System.Collections.Generic;
using System.Windows;
using System.IO;
using System.Drawing;
using System.Security;
using System.Security.Permissions;
using System.Threading;
using System.Windows.Documents;
using System.Windows.Media;

using IWshRuntimeLibrary;



namespace ColorfulTiles
{

    public partial class MainWindow : Window
    {
        private WshShell shell = new WshShell();

        private string pathForAllUsers = "C:\\ProgramData\\Microsoft\\Windows\\Start Menu\\Programs";
        private string pathForCurentUser = System.Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData) + "\\Microsoft\\Windows\\Start Menu\\Programs";
        private string tmpPathAllUsers = ".\\tmp\\allUsers";
        private string tmpPathCurUser = ".\\tmp\\curUser";
        private string XMLList = ".\\XMLList.txt";

        private List<string> allLnks = new List<string>();
        private List<string> allLnkTargets = new List<string>();
        private List<string> generatedXMLs = new List<string>();

        private delegate void taskDelegate();


        //=====
        // Helper functions to print info onto richtextbox

        private void out2box(string text)
        {
            OutputBox.Dispatcher.BeginInvoke(new Action(() =>
            {
                OutputBox.AppendText(text);
                OutputBox.ScrollToEnd();
            }));
        }

        private void out2box(string text, string color)
        {
            OutputBox.Dispatcher.BeginInvoke(new Action(() =>
            {
                BrushConverter bc = new BrushConverter();
                TextRange tr = new TextRange(OutputBox.Document.ContentEnd, OutputBox.Document.ContentEnd);
                tr.Text = text;
                tr.ApplyPropertyValue(TextElement.ForegroundProperty, bc.ConvertFromString(color));
                OutputBox.ScrollToEnd();
            }));
        }


        //=====
        // Helper functions to copy and remove files / folders 

        private void CopyAll(string sourceDir, string targetDir)
        {
            Directory.CreateDirectory(targetDir);
            foreach (string file in Directory.GetFiles(sourceDir))
            {
                string targetPath = Path.Combine(targetDir, Path.GetFileName(file));
                if (System.IO.File.Exists(targetPath))
                {
                    out2box("Already exists: " + targetPath + "\n", "red");
                    continue;
                }
                System.IO.File.Copy(file, targetPath);
            }
            foreach (string directory in Directory.GetDirectories(sourceDir))
                CopyAll(directory, Path.Combine(targetDir, Path.GetFileName(directory)));
        }

        private void RemoveAll(string sourceDir)
        {
            System.IO.DirectoryInfo di = new DirectoryInfo(sourceDir);
            foreach (FileInfo file in di.GetFiles())
            {
                try
                {
                    file.Delete();
                }
                catch (Exception e)
                {
                    out2box("Cannot delete: " + file.FullName + "\n", "red");
                }
            }
            foreach (DirectoryInfo dir in di.GetDirectories())
            {
                try
                {
                    dir.Delete(true);
                }
                catch (Exception e)
                {
                    out2box("Cannot delete: " + dir.FullName + "\n", "red");
                }
            }
        }



        public MainWindow()
        {
            InitializeComponent();
            if (System.IO.File.Exists(XMLList))
                generatedXMLs = TextListConverter.ReadFile(XMLList);
        }


        private void DisableAllButtons()
        {
            gen_button.IsEnabled = false;
            rem_button.IsEnabled = false;
            move_button.IsEnabled = false;
            restore_button.IsEnabled = false;
        }

        private void FinishCallback(IAsyncResult asyncResult)
        {
            out2box("Done\n");
            Dispatcher.Invoke(() =>
            {
                gen_button.IsEnabled = true;
                rem_button.IsEnabled = true;
                move_button.IsEnabled = true;
                restore_button.IsEnabled = true;
            });
        }


        
        private void Gen_Clicked(object sender, RoutedEventArgs e)
        {
            DisableAllButtons();
            taskDelegate t = GenXMLs;
            IAsyncResult asyncResult = t.BeginInvoke(FinishCallback, t);
        }
        
        private void Rem_Clicked(object sender, RoutedEventArgs e)
        {
            DisableAllButtons();
            taskDelegate t = RemoveXMLs;
            IAsyncResult asyncResult = t.BeginInvoke(FinishCallback, t);
        }
        
        private void Move_Clicked(object sender, RoutedEventArgs e)
        {
            DisableAllButtons();
            taskDelegate t = MoveShortcuts;
            IAsyncResult asyncResult = t.BeginInvoke(FinishCallback, t);
        }
        
        private void Restore_Clicked(object sender, RoutedEventArgs e)
        {
            DisableAllButtons();
            taskDelegate t = RestoreShortcuts;
            IAsyncResult asyncResult = t.BeginInvoke(FinishCallback, t);
        }


        //=====
        // Generate configure files for all apps in start menu

        private void GenXMLs()
        {
            out2box("Start scanning all start menu items...\n");
            allLnks.Clear();
            allLnkTargets.Clear();
            ScanLnks(new DirectoryInfo(pathForAllUsers));
            ScanLnks(new DirectoryInfo(pathForCurentUser));

            out2box("Start generating colors...\n");

            foreach (string targetFile in allLnkTargets)
            {
                if (!System.IO.File.Exists(targetFile))
                    continue;

                Icon icon = System.Drawing.Icon.ExtractAssociatedIcon(targetFile);
                Bitmap bitmap = icon.ToBitmap();
                int h = bitmap.Height;
                int w = bitmap.Width;
                int pixelCount = h * w;
                int meanR = 0, meanG = 0, meanB = 0;
                for (int i = 0; i < h; i++)
                {
                    for (int j = 0; j < w; j++)
                    {
                        meanR += bitmap.GetPixel(i, j).R;
                        meanG += bitmap.GetPixel(i, j).G;
                        meanB += bitmap.GetPixel(i, j).B;
                    }
                }
                meanR /= pixelCount;
                meanG /= pixelCount;
                meanB /= pixelCount;
                string hex = meanR.ToString("X2") + meanG.ToString("X2") + meanB.ToString("X2");

                double grayLevel = (0.299 * meanR + 0.587 * meanG + 0.114 * meanB) / 255.0;
                string fgColor = grayLevel > 0.5 ? "dark" : "light";

                WriteStyleXML(targetFile, fgColor, hex);
                
            }

            out2box("Writting XML list...\n");

            TextListConverter.AppendFile(generatedXMLs, XMLList);
        }


        private void RemoveXMLs()
        {
            out2box("Removing created XML files\n");
            List<string> removed = new List<string>();
            foreach (string xmlfile in generatedXMLs)
            {
                try
                {
                    System.IO.File.Delete(xmlfile);
                    removed.Add(xmlfile);
                }
                catch (Exception e)
                {
                    out2box("Cannot delete: " + xmlfile + "\n", "red");
                }
            }
            foreach (string xmlfile in removed)
                generatedXMLs.Remove(xmlfile);
            TextListConverter.WriteFile(generatedXMLs, XMLList);
        }


        private void MoveShortcuts()
        {
            out2box("Moving shortcuts...\n");

            CopyAll(pathForAllUsers, tmpPathAllUsers);
            CopyAll(pathForCurentUser, tmpPathCurUser);

            RemoveAll(pathForAllUsers);
            RemoveAll(pathForCurentUser);
        }


        private void RestoreShortcuts()
        {
            out2box("Restoring shortcuts...\n");

            CopyAll(tmpPathAllUsers, pathForAllUsers);
            CopyAll(tmpPathCurUser, pathForCurentUser);

            RemoveAll(tmpPathAllUsers);
            RemoveAll(tmpPathCurUser);
        }


        //=====
        // Scan the start menu for app list.

        private void ScanLnks(FileSystemInfo info)
        {
            if (!info.Exists) return;
            DirectoryInfo dir = info as DirectoryInfo;

            if (dir == null) return;
            FileSystemInfo[] files = dir.GetFileSystemInfos();

            for (int i = 0; i < files.Length; i++)
            {
                FileInfo file = files[i] as FileInfo;
                if (file != null && file.FullName.EndsWith(".lnk"))
                {
                    IWshShortcut targetIWsh = (IWshShortcut)shell.CreateShortcut(file.FullName);
                    if (string.IsNullOrEmpty(targetIWsh.TargetPath))
                        continue;
                    allLnks.Add(file.FullName);
                    allLnkTargets.Add(targetIWsh.TargetPath);
                }
                else
                    ScanLnks(files[i]);
            }
        }


        //=====
        // Generate XML file for a specific app.

        private void WriteStyleXML(string targetFile, string fgColor, string bgColor)
        {
            string xmlPath = Path.GetDirectoryName(targetFile);
            string xmlFile = Path.GetFileNameWithoutExtension(targetFile) + ".VisualElementsManifest.xml";
            string xmlFull = Path.Combine(xmlPath, xmlFile);

            PermissionSet permissionSet = new PermissionSet(PermissionState.None);
            FileIOPermission writePermission = new FileIOPermission(FileIOPermissionAccess.AllAccess, xmlFull);
            //FileIOPermissionAccess.AllAccess
            permissionSet.AddPermission(writePermission);
            if (!permissionSet.IsSubsetOf(AppDomain.CurrentDomain.PermissionSet))
            {
                out2box("Permission denied: " + xmlFull + "\n", "red");
                return;
            }

            if (System.IO.File.Exists(xmlFull))
            {
                out2box("Already exists: " + xmlFull + "\n", "blue");
                return;
            }

            try
            {
                FileStream fs = new FileStream(xmlFull, FileMode.OpenOrCreate, FileAccess.Write);
                StreamWriter sw = new StreamWriter(fs);
                sw.Flush();
                sw.BaseStream.Seek(0, SeekOrigin.Begin);

                sw.WriteLine("<Application>");
                sw.WriteLine("  <VisualElements");
                sw.WriteLine("      ShowNameOnSquare150x150Logo='on'");
                sw.WriteLine("      ForegroundText='" + fgColor + "'");
                sw.WriteLine("      BackgroundColor='#" + bgColor + "'/>");
                sw.WriteLine("</Application>");

                sw.Flush();
                sw.Close();
                fs.Close();

                out2box("Success: " + xmlFull + "\n", "green");

                generatedXMLs.Add(xmlFull);
            }
            catch (UnauthorizedAccessException ex)
            {
                out2box("Permission denied: " + xmlFull + "\n", "red");
            }

            

        }

    }
}
