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
                        var oSheet = (Excel.Worksheet)oWB.Sheets["General"];
                        oSheet.Cells[2, 1] = Studio.assetsManager.assetBundlesCount.ToString();
                        oSheet.Cells[2, 2] = Studio.assetsManager.assetBundlesTotalSize.ToString();

                        for (int i = 0; i < 10; i++)
                        {
                            if (i < topSizes.Count)
                            {
                                oSheet.Cells[1, 3 + i] = topSizes[i].a;
                                oSheet.Cells[2, 3 + i] = topSizes[i].b.ToString();
                            }
                            else
                            {
                                oSheet.Cells[1, 3 + i] = "NULL_" + i.ToString();
                                oSheet.Cells[2, 3 + i] = "0";
                            }
                        }

                        var chartObjects = (Excel.ChartObjects)oSheet.ChartObjects();
                        var chart = chartObjects.Add(10, 40, 400, 400).Chart;
                        chart.ChartType = Excel.XlChartType.xlPie;
                        chart.SetSourceData(oSheet.Range["C1", "L2"]);

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
                            if (Progress.stopTask) return;

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
                        var oSheet = (Excel.Worksheet)oWB.Sheets["Redundancies"];

                        int i = 2;
                        int j = 0;
                        string[,] staging = new string[128, 6];
                        Progress.Reset("Writing redundancies into excel");
                        int p = 0;
                        foreach (var r in redundancies)
                        {
                            foreach (var ai in r.Value)
                            {
                                if (Progress.stopTask) return;

                                if (ai.count > 1)
                                {
                                    staging[j, 0] = ai.name;
                                    staging[j, 1] = ai.type;
                                    staging[j, 2] = ai.count.ToString();
                                    staging[j, 3] = ai.size.ToString();
                                    staging[j, 4] = (ai.size * ai.count).ToString();
                                    staging[j, 5] = r.Key.ToString();
                                    j++;
                                    if (j >= 128)
                                    {
                                        oSheet.Range["A" + i, "F" + (i + j - 1)].Value2 = staging;
                                        i += j;
                                        j = 0;
                                    }
                                }
                            }
                            p++;
                            Progress.Report(p, redundancies.Count);
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
