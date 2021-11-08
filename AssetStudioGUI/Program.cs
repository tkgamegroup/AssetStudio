using System;
using System.IO;
using System.Collections.Generic;
using System.Windows.Forms;
using System.Threading;
using AssetStudio;
using Excel = Microsoft.Office.Interop.Excel;

namespace AssetStudioGUI
{
    static class Program
    {
        class RedundantAsset
        {
            public string type;
            public string name;
            public int size;
            public int count = 1;
            public string pathID;

            public RedundantAsset(string type, string name, int size, string pathID)
            {
                this.type = type;
                this.name = name;
                this.size = size;
                this.pathID = pathID;
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
        class ConsoleProgress : IProgress
        {
            public void Reset(string task)
            {
                Console.WriteLine();
                Console.WriteLine(task);
            }

            public void Report(int current, int total)
            {
                Console.Write("\r{0}/{1}", current, total);
            }
        }

        [STAThread]
        static void Main()
        {
            var args = Environment.GetCommandLineArgs();
            if (args.Length == 4 && args[1] == "analyze")
            {
                string inputDir = args[2];
                string savePath = args[3];

                if (!Directory.Exists(inputDir))
                    return;

                Progress.Default = new ConsoleProgress();

                var oXL = new Excel.Application();

                try
                {
                    var oWB = oXL.Workbooks.Open(Directory.GetCurrentDirectory() + "\\template.xlsm");

                    Studio.assetsManager.LoadFolder(inputDir);

                    var assetItems = new List<AssetItem>();
                    Studio.BuildAssetData(assetItems);

                    var redundancies_map = new Dictionary<long, List<RedundantAsset>>();
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
                        if (redundancies_map.TryGetValue(a.m_PathID, out var list))
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
                                list.Add(new RedundantAsset(a.TypeString, a.Text, (int)a.FullSize, a.m_PathID.ToString()));
                            }
                        }
                        else
                        {
                            var new_list = new List<RedundantAsset>();
                            new_list.Add(new RedundantAsset(a.TypeString, a.Text, (int)a.FullSize, a.m_PathID.ToString()));
                            redundancies_map.Add(a.m_PathID, new_list);
                        }
                    }
                    var redundancies = new List<RedundantAsset>();
                    foreach (var r in redundancies_map)
                    {
                        foreach (var ai in r.Value)
                        {
                            if (ai.count > 1)
                            {
                                redundancies.Add(ai);
                            }
                        }
                    }

                    {
                        var oSheet = (Excel.Worksheet)oWB.Sheets["General"];
                        oSheet.Cells[1, 2] = Studio.assetsManager.assetBundleInfos.Count.ToString();
                        oSheet.Cells[2, 2] = Studio.assetsManager.assetBundlesTotalSize.ToString();

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

                        for (int i = 0; i < 10; i++)
                        {
                            if (i < topSizes.Count)
                            {
                                oSheet.Cells[4 + i, 1] = topSizes[i].a;
                                oSheet.Cells[4 + i, 2] = topSizes[i].b.ToString();
                            }
                            else
                            {
                                oSheet.Cells[4 + i, 1] = "NULL_" + i.ToString();
                                oSheet.Cells[4 + i, 2] = "0";
                            }
                        }

                        topSizes.Clear();

                        foreach (var a in redundancies)
                        {
                            bool new_type = true;
                            foreach (var i in topSizes)
                            {
                                if (i.a == a.type)
                                {
                                    i.b += a.size * (a.count - 1);
                                    new_type = false;
                                    break;
                                }
                            }
                            if (new_type)
                            {
                                topSizes.Add(new StringLongPair(a.type, a.size * (a.count - 1)));
                            }
                        }
                        topSizes.Sort(delegate (StringLongPair a, StringLongPair b)
                        {
                            if (a.b == b.b) return 0;
                            else if (a.b < b.b) return 1;
                            else return -1;
                        });

                        for (int i = 0; i < 10; i++)
                        {
                            if (i < topSizes.Count)
                            {
                                oSheet.Cells[15 + i, 1] = topSizes[i].a;
                                oSheet.Cells[15 + i, 2] = topSizes[i].b.ToString();
                            }
                            else
                            {
                                oSheet.Cells[15 + i, 1] = "NULL_" + i.ToString();
                                oSheet.Cells[15 + i, 2] = "0";
                            }
                        }

                        var chartObjects = (Excel.ChartObjects)oSheet.ChartObjects();
                        var chart1 = chartObjects.Add(200, 10, 400, 200).Chart;
                        chart1.ChartType = Excel.XlChartType.xlPie;
                        chart1.SetSourceData(oSheet.Range["A4", "B13"]);
                        var chart2 = chartObjects.Add(200, 250, 400, 200).Chart;
                        chart2.ChartType = Excel.XlChartType.xlPie;
                        chart2.SetSourceData(oSheet.Range["A15", "B24"]);

                        var header = (Excel.Range)oSheet.Rows[1];
                        header.EntireColumn.AutoFit();
                    }

                    {
                        var oSheet = (Excel.Worksheet)oWB.Sheets["Assets"];

                        string ab_name = "";
                        int i = 2;
                        int j = 0;
                        string[,] staging = new string[128, 6];
                        Progress.Reset("Writing assets into excel");
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
                            staging[j, 0] = a.Text;
                            staging[j, 1] = a.Container;
                            staging[j, 2] = ab_name;
                            staging[j, 3] = a.TypeString;
                            staging[j, 4] = a.FullSize.ToString();
                            staging[j, 5] = a.m_PathID.ToString();
                            j++;
                            if (j >= 128)
                            {
                                oSheet.Range["A" + i, "F" + (i + j - 1)].Value2 = staging;
                                i += j;
                                j = 0;
                                Progress.Report(i - 1, assetItems.Count);
                            }
                        }
                        if (j > 0)
                        {
                            oSheet.Range["A" + i, "F" + (i + j - 1)].Value2 = staging;
                            i += j;
                            j = 0;
                            Progress.Report(i - 1, assetItems.Count);
                        }

                        var header = (Excel.Range)oSheet.Rows[1];
                        header.EntireColumn.AutoFit();
                    }

                    {
                        var oSheet = (Excel.Worksheet)oWB.Sheets["redundancies"];

                        int i = 2;
                        int j = 0;
                        string[,] staging = new string[128, 6];
                        Progress.Reset("Writing redundancies into excel");
                        foreach (var ai in redundancies)
                        {
                            staging[j, 0] = ai.name;
                            staging[j, 1] = ai.type;
                            staging[j, 2] = ai.count.ToString();
                            staging[j, 3] = ai.size.ToString();
                            staging[j, 4] = (ai.size * (ai.count - 1)).ToString();
                            staging[j, 5] = ai.pathID;
                            j++;
                            if (j >= 128)
                            {
                                oSheet.Range["A" + i, "F" + (i + j - 1)].Value2 = staging;
                                i += j;
                                j = 0;
                                Progress.Report(i - 1, redundancies.Count);
                            }
                        }
                        if (j > 0)
                        {
                            oSheet.Range["A" + i, "F" + (i + j - 1)].Value2 = staging;
                            i += j;
                            j = 0;
                            Progress.Report(i - 1, redundancies.Count);
                        }

                        var header = (Excel.Range)oSheet.Rows[1];
                        header.EntireColumn.AutoFit();
                    }

                    {
                        var oSheet = (Excel.Worksheet)oWB.Sheets["AssetBundles"];

                        foreach (var abi in Studio.assetsManager.assetBundleInfos)
                        {
                            string fn_manifest = abi.path + ".manifest";
                            if (!File.Exists(fn_manifest)) continue;

                            StreamReader manifest = new StreamReader(fn_manifest);
                            string line;
                            int tag = 0; // reading nothing, reading assets and reading dependencies
                            while ((line = manifest.ReadLine()) != null)
                            {
                                if (line == "Assets:")
                                {
                                    tag = 1;
                                }
                                else if (line == "Dependencies:")
                                {
                                    tag = 2;
                                }
                                else if (line[0] == '-')
                                {
                                    if (tag == 1) abi.assetsCount++;
                                    else if (tag == 2) abi.dependenciesCount++;
                                }
                            }
                        }

                        int i = 2;
                        int j = 0;
                        string[,] staging = new string[128, 6];
                        Progress.Reset("Writing assetbundles into excel");
                        foreach (var abi in Studio.assetsManager.assetBundleInfos)
                        {
                            staging[j, 0] = abi.path;
                            staging[j, 1] = abi.assetsCount.ToString();
                            staging[j, 2] = abi.dependenciesCount.ToString();
                            j++;
                            if (j >= 128)
                            {
                                oSheet.Range["A" + i, "C" + (i + j - 1)].Value2 = staging;
                                i += j;
                                j = 0;
                                Progress.Report(i - 1, redundancies.Count);
                            }
                        }
                        if (j > 0)
                        {
                            oSheet.Range["A" + i, "C" + (i + j - 1)].Value2 = staging;
                            i += j;
                            j = 0;
                            Progress.Report(i - 1, redundancies.Count);
                        }

                        var header = (Excel.Range)oSheet.Rows[1];
                        header.EntireColumn.AutoFit();
                    }

                    oWB.SaveAs(savePath);
                }
                catch (Exception e)
                {
                    MessageBox.Show("Remeber to save your excel!!!\nThe exception is:\n" + e.Message, "Exception occured!!", MessageBoxButtons.OK, MessageBoxIcon.Warning, MessageBoxDefaultButton.Button1, MessageBoxOptions.DefaultDesktopOnly);
                }

                oXL.Visible = true;
                return;
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
