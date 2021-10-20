using System;
using System.IO;
using System.Collections.Generic;
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

        [STAThread]
        static void Main()
        {
            var args = Environment.GetCommandLineArgs();
            if (args.Length > 1)
            {
                if (args[1] == "analyze")
                {
                    string inputDir = "";
                    string savePath = "";
                    {
                        var fbd = new FolderBrowserDialog();
                        fbd.Description = "Select the AssetBundles Directory";
                        fbd.ShowNewFolderButton = false;
                        if (fbd.ShowDialog() == DialogResult.OK)
                        {
                            inputDir = fbd.SelectedPath;
                        }
                    }
                    {
                        var sfd = new SaveFileDialog();
                        sfd.Filter = "excel|*.xlsx";
                        sfd.RestoreDirectory = true;
                        if (sfd.ShowDialog() == DialogResult.OK)
                        {
                            savePath = sfd.FileName;
                        }
                    }

                    if (Directory.Exists(inputDir))
                    {
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

                        try
                        {
                            var oXL = new Excel.Application();

                            var oWB = oXL.Workbooks.Add(Missing.Value);

                            {
                                var oSheet = (Excel.Worksheet)oWB.ActiveSheet;
                                oSheet.Name = "General";
                                oSheet.Cells[1, 1] = "Bundles Count";
                                oSheet.Cells[2, 1] = Studio.assetsManager.assetBundlesCount.ToString();
                                oSheet.Cells[1, 2] = "Bundles Total Size";
                                oSheet.Cells[2, 2] = Studio.assetsManager.assetBundlesTotalSize.ToString();

                                for (int i = 0; i < 10; i++)
                                {
                                    oSheet.Cells[1, 3 + i] = topSizes[i].a;
                                    oSheet.Cells[2, 3 + i] = topSizes[i].b.ToString();
                                }

                                var chartObjects = (Excel.ChartObjects)oSheet.ChartObjects();
                                var chart = chartObjects.Add(10, 40, 400, 400).Chart;
                                chart.ChartType = Excel.XlChartType.xlPie;
                                chart.SetSourceData(oSheet.Range["C1", "L2"]);

                                var header = (Excel.Range)oSheet.Rows[1];
                                header.Interior.Color = System.Drawing.Color.FromArgb(84, 130, 53);
                                header.Font.Color = System.Drawing.Color.FromArgb(255, 255, 255);
                                header.Font.Bold = true;

                                var row2 = (Excel.Range)oSheet.Rows[2];
                                row2.NumberFormat = ExcelFileSizeFmt;

                                header.EntireColumn.AutoFit();
                            }

                            {
                                var oSheet = (Excel.Worksheet)oWB.Sheets.Add();
                                oSheet.Name = "Assets";
                                oSheet.Cells[1, 1] = "Name";
                                oSheet.Cells[1, 2] = "Container";
                                oSheet.Cells[1, 3] = "Bundle";
                                oSheet.Cells[1, 4] = "Type";
                                oSheet.Cells[1, 5] = "Size";
                                oSheet.Cells[1, 6] = "PathID";

                                var header = (Excel.Range)oSheet.Rows[1];
                                header.AutoFilter2(1, Type.Missing, Excel.XlAutoFilterOperator.xlAnd, Type.Missing, true);
                                header.Interior.Color = System.Drawing.Color.FromArgb(84, 130, 53);
                                header.Font.Color = System.Drawing.Color.FromArgb(255, 255, 255);
                                header.Font.Bold = true;

                                oSheet.Range["A2"].EntireColumn.NumberFormat = "@";
                                oSheet.Range["E2"].EntireColumn.NumberFormat = ExcelFileSizeFmt;
                                oSheet.Range["F2"].EntireColumn.NumberFormat = "@";

                                string ab_name = "";
                                int i = 2;
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
                                    oSheet.Cells[i, 1] = a.Text;
                                    oSheet.Cells[i, 2] = a.Container;
                                    oSheet.Cells[i, 3] = ab_name;
                                    oSheet.Cells[i, 4] = a.TypeString;
                                    oSheet.Cells[i, 5] = a.FullSize.ToString();
                                    oSheet.Cells[i, 6] = a.m_PathID.ToString();
                                    i++;
                                }

                                Excel.FormatCondition format = oSheet.UsedRange.FormatConditions.Add(Excel.XlFormatConditionType.xlExpression, Excel.XlFormatConditionOperator.xlEqual, "=MOD(ROW(),2)=0");
                                format.Interior.Color = System.Drawing.Color.FromArgb(226, 239, 218);

                                header.EntireColumn.AutoFit();

                                oSheet.Select();
                                var row2 = (Excel.Range)oSheet.Rows[2];
                                row2.Select();
                                oXL.ActiveWindow.FreezePanes = true;
                            }

                            {
                                var oSheet = (Excel.Worksheet)oWB.Sheets.Add();
                                oSheet.Name = "Redundancies";
                                oSheet.Cells[1, 1] = "Name";
                                oSheet.Cells[1, 2] = "Type";
                                oSheet.Cells[1, 3] = "Count";
                                oSheet.Cells[1, 4] = "Size";
                                oSheet.Cells[1, 5] = "Total Size";
                                oSheet.Cells[1, 6] = "PathID";

                                var header = (Excel.Range)oSheet.Rows[1];
                                header.AutoFilter2(1, Type.Missing, Excel.XlAutoFilterOperator.xlAnd, Type.Missing, true);
                                header.Interior.Color = System.Drawing.Color.FromArgb(84, 130, 53);
                                header.Font.Color = System.Drawing.Color.FromArgb(255, 255, 255);
                                header.Font.Bold = true;

                                oSheet.Range["A2"].EntireColumn.NumberFormat = "@";
                                oSheet.Range["D2"].EntireColumn.NumberFormat = ExcelFileSizeFmt;
                                oSheet.Range["E2"].EntireColumn.NumberFormat = ExcelFileSizeFmt;
                                oSheet.Range["F2"].EntireColumn.NumberFormat = "@";

                                int i = 2;
                                foreach (var r in redundancies)
                                {
                                    foreach (var ai in r.Value)
                                    {
                                        if (ai.count > 1)
                                        {
                                            oSheet.Cells[i, 1] = ai.name;
                                            oSheet.Cells[i, 2] = ai.type;
                                            oSheet.Cells[i, 3] = ai.count.ToString();
                                            oSheet.Cells[i, 4] = ai.size;
                                            oSheet.Cells[i, 5] = (ai.size * ai.count).ToString();
                                            oSheet.Cells[i, 6] = r.Key.ToString();
                                            i++;
                                        }
                                    }
                                }

                                Excel.FormatCondition format = oSheet.UsedRange.FormatConditions.Add(Excel.XlFormatConditionType.xlExpression, Excel.XlFormatConditionOperator.xlEqual, "=MOD(ROW(),2)=0");
                                format.Interior.Color = System.Drawing.Color.FromArgb(226, 239, 218);

                                header.EntireColumn.AutoFit();

                                oSheet.Select();
                                var row2 = (Excel.Range)oSheet.Rows[2];
                                row2.Select();
                                oXL.ActiveWindow.FreezePanes = true;
                            }

                            oXL.Visible = true;
                            oWB.SaveAs2(savePath);
                        }
                        catch (Exception e)
                        {
                            MessageBox.Show("Remeber to save your excel!!!\nThe exception is:\n" + e.Message, "Exception occured!!", MessageBoxButtons.OK, MessageBoxIcon.Warning, MessageBoxDefaultButton.Button1, MessageBoxOptions.DefaultDesktopOnly);
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
