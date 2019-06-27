using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Navigation;
using System.Windows.Shapes;
using System.IO;
using System.Drawing;
using IWshRuntimeLibrary;


namespace ColorfulTiles
{
    public partial class MainWindow : Window
    {
        private WshShell shell = new WshShell();

        private List<string> allLnks = new List<string>();
        private List<string> allLnkTargets = new List<string>();
        private bool lnkLoaded = false;

        public MainWindow()
        {
            InitializeComponent();
        }

        private void Start_Clicked(object sender, RoutedEventArgs e)
        {
            if (!lnkLoaded)
            {
                string pathForAllUsers = "C:\\ProgramData\\Microsoft\\Windows\\Start Menu\\Programs";
                string pathForCurentUser = System.Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData) + "\\Microsoft\\Windows\\Start Menu\\Programs";
                ScanLnks(new DirectoryInfo(pathForAllUsers));
                ScanLnks(new DirectoryInfo(pathForCurentUser));
                TextListConverter.WriteFile(allLnks, "LnkPathes.txt");
                TextListConverter.WriteFile(allLnkTargets, "LnkTargets.txt");
                lnkLoaded = true;
            }

            foreach (string targetFile in allLnkTargets)
            {
                Icon icon = System.Drawing.Icon.ExtractAssociatedIcon(targetFile);
                Bitmap bitmap = icon.ToBitmap();
                UInt64 h = (UInt64)bitmap.Height;
                UInt64 w = (UInt64)bitmap.Width;
                UInt64 pixelCount = h * w;
                UInt64 meanR = 0, meanG = 0, meanB = 0;
                for (UInt64 i = 0; i < h; i++)
                {
                    for (UInt64 j = 0; j < w; j++)
                    {
                        meanR += bitmap.GetPixel(i, j).R;
                        meanG += bitmap.GetPixel(i, j).G;
                        meanB += bitmap.GetPixel(i, j).B;
                    }
                }
                meanR /= pixelCount;
                meanG /= pixelCount;
                meanB /= pixelCount;
                string hex = myColor.R.ToString("X2") + myColor.G.ToString("X2") + myColor.B.ToString("X2");
            }
            


        }


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


        private void WriteStyleXML(string targetFile, string fgColor, string bgColor)
        {
            string xmlFile = System.IO.Path.GetFileNameWithoutExtension(targetFile) + ".VisualElementsManifest.xml";
            FileStream fs = new FileStream(xmlFile, FileMode.OpenOrCreate, FileAccess.Write);
            StreamWriter sw = new StreamWriter(fs);
            sw.Flush();
            sw.BaseStream.Seek(0, SeekOrigin.Begin);

            sw.WriteLine("<Application xmlns:xsi='http://www.w3.org/2001/XMLSchema-instance'>");
            sw.WriteLine("  <VisualElements");
            sw.WriteLine("      ShowNameOnSquare150x150Logo='on'");
            sw.WriteLine("      ShowNameOnSquare150x150Logo='on'");
            sw.WriteLine("      ForegroundText='" + fgColor + "'");
            sw.WriteLine("      BackgroundColor='" + bgColor + "'/>");
            sw.WriteLine("</Application>");

            sw.Flush();
            sw.Close();
            fs.Close();
        }


    }
}
