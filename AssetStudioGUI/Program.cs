using System;
using System.IO;
using System.Collections.Generic;
using System.Diagnostics;
using System.Reflection;
using System.Windows.Forms;
using AssetStudio;
using Excel = Microsoft.Office.Interop.Excel;

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

        const string ExcelFileSizeFmt = "[<1048576]0.00,\" KB\";[<1073741824]0.00,,\" MB\";0.00,,,\" GB\"";

        class AssetIdentifier
        {
            public string type;
            public string name;
            public int size;
            public int count = 1;

            public AssetIdentifier(string _type, string _name, int _size)
            {
                type = _type;
                name = _name;
                size = _size;
            }
        }

        class StringLongPair
        {
            public string a;
            public long b;

            public StringLongPair(string _a, long _b)
            {
                a = _a;
                b = _b;
            }
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
                if (args[1] == "analyze")
                {
                    if (args.Length == 4)
                    {
                        if (Directory.Exists(args[2]) && Directory.Exists(args[3]))
                        {
                            Studio.assetsManager.doLog = true;
                            Studio.assetsManager.LoadFolder(args[2]);

                            var assetItems = new List<AssetItem>();
                            Studio.BuildAssetData(assetItems);

                            var redundancies = new Dictionary<long, List<AssetIdentifier>>();
                            foreach (var a in assetItems)
                            {
                                if (a.Type == ClassIDType.AssetBundle || a.Type == ClassIDType.AssetBundleManifest ||
                                    a.Type == ClassIDType.GameObject || 
                                    a.Type == ClassIDType.MeshFilter || a.Type == ClassIDType.MeshRenderer || a.Type == ClassIDType.SkinnedMeshRenderer ||
                                    a.Type == ClassIDType.Transform ||
                                    a.Type == ClassIDType.BoxCollider || a.Type == ClassIDType.MeshCollider ||
                                    a.Type == ClassIDType.Animator || a.Type == ClassIDType.AnimatorController ||
                                    a.Type == ClassIDType.ParticleSystemRenderer ||
                                    a.Type == ClassIDType.Light || a.Type == ClassIDType.Camera ||
                                    a.Type == ClassIDType.MonoBehaviour || a.Type == ClassIDType.MonoScript)
                                {
                                    continue;
                                }
                                if (redundancies.TryGetValue(a.m_PathID, out var list))
                                {
                                    bool found = false;
                                    foreach (var i in list)
                                    {
                                        if (i.type == a.TypeString && i.size == a.FullSize && i.name == a.Text)
                                        {
                                            i.count++;
                                            found = true;
                                        }
                                    }
                                    if (!found)
                                    {
                                        list.Add(new AssetIdentifier(a.TypeString, a.Text, (int)a.FullSize));
                                    }
                                }
                                else
                                {
                                    var new_list = new List<AssetIdentifier>();
                                    new_list.Add(new AssetIdentifier(a.TypeString, a.Text, (int)a.FullSize));
                                    redundancies.Add(a.m_PathID, new_list);
                                }
                            }

                            var topSizes = new List<StringLongPair>();
                            foreach (var a in assetItems)
                            {
                                bool new_type = true;
                                foreach (var i in topSizes)
                                {
                                    if (i.a == a.TypeString)
                                    {
                                        i.b += a.FullSize;
                                        new_type = false;
                                        break;
                                    }
                                }
                                if (new_type)
                                {
                                    topSizes.Add(new StringLongPair(a.TypeString, a.FullSize));
                                }
                            }
                            topSizes.Sort(delegate (StringLongPair a, StringLongPair b)
                            {
                                if (a.b == b.b) return 0;
                                else if (a.b < b.b) return 1;
                                else return -1;
                            });

                            {
                                var oXL = new Excel.Application();
                                oXL.Visible = true;

                                var oWB = oXL.Workbooks.Add(Missing.Value);
                                var oSheetGeneral = (Excel.Worksheet)oWB.ActiveSheet;
                                oSheetGeneral.Name = "一般";
                                oSheetGeneral.Cells[1, 1] = "包数量";
                                oSheetGeneral.Cells[2, 1] = Studio.assetsManager.assetBundlesCount.ToString();
                                oSheetGeneral.Cells[1, 2] = "总大小";
                                oSheetGeneral.Cells[2, 2] = Studio.assetsManager.assetBundlesTotalSize.ToString();

                                for (int i = 0; i < 10; i++)
                                {
                                    oSheetGeneral.Cells[1, 3 + i] = topSizes[i].a;
                                    oSheetGeneral.Cells[2, 3 + i] = topSizes[i].b.ToString();
                                }

                                oSheetGeneral.Range["A1", "L1"].EntireColumn.AutoFit();
                                oSheetGeneral.Range["B2", "L2"].NumberFormat = ExcelFileSizeFmt;
                            }

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
                                   a.m_PathID, a.Text, a.Container, ab_name, a.TypeString, a.FullSize));
                            }
                            assetsFile.Close();

                            var redundanciesFile = new StreamWriter(args[3] + "/Redundancies List.txt");
                            redundanciesFile.WriteLine("PathID\tCount\tType\tSize\tTotal Size");
                            foreach (var r in redundancies)
                            {
                                foreach (var i in r.Value)
                                {
                                    if (i.count > 1)
                                    {
                                        redundanciesFile.WriteLine(string.Format("{0}\t{1}\t{2}\t{3}\t{4}", r.Key, i.count, i.type, i.size, i.size * i.count));
                                    }
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
