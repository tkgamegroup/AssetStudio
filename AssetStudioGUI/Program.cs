using System;
using System.IO;
using System.Collections.Generic;
using System.Diagnostics;
using System.Windows.Forms;
using AssetStudio;

namespace AssetStudioGUI
{
    static class Program
    {
        static string FormatFileSize(float len)
        {
            string[] sizes = { "B", "KB", "MB", "GB", "TB" };
            int order = 0;
            while (len >= 1024F && order < sizes.Length - 1)
            {
                order++;
                len = len / 1024F;
            }

            return string.Format("{0:0.##} {1}", len, sizes[order]);
        }

        /// <summary>
        ///  The main entry point for the application.
        /// </summary>
        [STAThread]
        static void Main()
        {
            var args = Environment.GetCommandLineArgs();
            if (args.Length > 1)
            {
                if (args[1] == "redundancy_analyze")
                {
                    if (args.Length == 4)
                    {
                        if (Directory.Exists(args[2]) && Directory.Exists(args[3]))
                        {
                            Studio.assetsManager.doLog = true;
                            Studio.assetsManager.LoadFolder(args[2]);

                            var assetItems = new List<AssetItem>();
                            Studio.BuildAssetData(assetItems);

                            var miscFile = new StreamWriter(args[3] + "/Asset Bundles Misc.txt");
                            miscFile.WriteLine(string.Format("Count: {0}", Studio.assetsManager.assetBundlesCount));
                            miscFile.WriteLine(string.Format("Total Size: {0}", FormatFileSize(Studio.assetsManager.assetBundlesTotalSize)));
                            miscFile.Close();

                            var assetsFile = new StreamWriter(args[3] + "/Assets List.txt");
                            assetsFile.WriteLine("PathID\tName\tContainer\tAsset Bundle\tType\tSize");
                            string ab_name = "";
                            foreach (var a in assetItems)
                            {
                                if (a.Type == ClassIDType.AssetBundleManifest)
                                {
                                    continue;
                                }
                                if (a.Type == ClassIDType.AssetBundle)
                                {
                                    ab_name = a.Text;
                                    continue;
                                }
                                assetsFile.WriteLine(string.Format("{0}\t{1}\t{2}\t{3}\t{4}\t{5}",
                                   a.m_PathID, a.Text, a.Container, ab_name, a.TypeString, FormatFileSize(a.FullSize)));
                            }
                            assetsFile.Close();

                            var redundancies = new Dictionary<long, Tuple<int, string, int>>();
                            foreach (var a in assetItems)
                            {
                                if (a.Type == ClassIDType.AssetBundleManifest)
                                {
                                    continue;
                                }
                                Tuple<int, string, int> r;
                                if (redundancies.TryGetValue(a.m_PathID, out r))
                                {
                                    Debug.Assert(r.Item2 == a.TypeString && r.Item3 == a.FullSize);
                                    redundancies[a.m_PathID] = Tuple.Create(r.Item1 + 1, a.TypeString, (int)a.FullSize);
                                }
                                else
                                {
                                    redundancies.Add(a.m_PathID, Tuple.Create(1, a.TypeString, (int)a.FullSize));
                                }
                            }

                            var redundanciesFile = new StreamWriter(args[3] + "/Redundancies List.txt");
                            redundanciesFile.WriteLine("PathID\tCount\tType\tSize\tTotal Size");
                            foreach (var r in redundancies)
                            {
                                if (r.Key != 1 && r.Value.Item1 > 1)
                                {
                                    redundanciesFile.WriteLine(string.Format("{0}\t{1}\t{2}\t{3}\t{4}", r.Key, r.Value.Item1, r.Value.Item2, FormatFileSize(r.Value.Item3), FormatFileSize(r.Value.Item3 * r.Value.Item1)));
                                }
                            }
                            redundanciesFile.Close();
                        }
                    }
                    return;
                }
            }

#if !NETFRAMEWORK
            Application.SetHighDpiMode(HighDpiMode.SystemAware);
#endif
            Application.EnableVisualStyles();
            Application.SetCompatibleTextRenderingDefault(false);
            Application.Run(new AssetStudioGUIForm());
        }
    }
}
