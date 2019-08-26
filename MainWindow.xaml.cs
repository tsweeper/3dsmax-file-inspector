using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Media;
using Ookii.Dialogs.Wpf;
using OpenMcdf;

namespace _3DSMaxFileVersion
{
    /// <summary>
    ///     Interaction logic for MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {
        public string TargetPath;

        public MainWindow()
        {
            InitializeComponent();
        }

        private void CheckBox1_OnClick(object sender, RoutedEventArgs e)
        {
            if (Directory.Exists(TargetPath))
                //GetFiles(TargetPath);
                RefreshListView();
        }

        private void BtnPath_Click(object sender, RoutedEventArgs e)
        {
            var dialog = new VistaFolderBrowserDialog();
            if (!dialog.ShowDialog(this).GetValueOrDefault()) return;
            TargetPath = dialog.SelectedPath;
            LblPath.Text = "Checking: " + TargetPath;

            //GetFiles(dialog.SelectedPath);
            RefreshListView();
        }

        private void GetFiles(string directory)
        {
            // old method
            //var files = CheckBox1.IsChecked.HasValue && CheckBox1.IsChecked.Value
            //? Directory.GetFiles(TargetPath, "*.max", SearchOption.AllDirectories)
            //: Directory.GetFiles(TargetPath, "*.max", SearchOption.TopDirectoryOnly);
            //? Directory.EnumerateFiles(TargetPath, "*.max", SearchOption.AllDirectories)
            //: Directory.EnumerateFiles(TargetPath, "*.max", SearchOption.TopDirectoryOnly);

            //var files = new List<string>();
            //var includeSubdir = CheckBox1.IsChecked != null && CheckBox1.IsChecked.Value;
            //CheckBox1.IsEnabled = false;
            //BtnPath.IsEnabled = false;

            // Create an enumeration of the files we will want to process that simply accumulates these values...
            long total = 0;
            var findFile = new CSharpTest.Net.IO.FindFile(directory, "*.max", true, true, true) {RaiseOnAccessDenied = false};
            findFile.FileFound +=
                (o, e) =>
                {
                    if (!e.IsDirectory)
                    {
                        Interlocked.Increment(ref total);
                    }
                };

            // Start a high-priority thread to perform the accumulation
            var t = new Thread(findFile.Find)
            {
                IsBackground = true,
                Priority = ThreadPriority.AboveNormal,
                Name = "Enumerate Files"
            };
            t.Start();

            // Allow the accumulator thread to get a head-start on us
            do { Thread.Sleep(100); }
            while (total < 100 && t.IsAlive);

            // Now we can process the files normally and update a percentage
            long count = 0, percentage = 0;
            var task = new CSharpTest.Net.IO.FindFile(directory, "*.max", true, true, true) {RaiseOnAccessDenied = false};
            task.FileFound +=
                (o, e) =>
                {
                    if (!e.IsDirectory)
                    {
                        //ProcessFile(e.FullPath);
                        Console.WriteLine($@"found {e.FullPath}");
                        // Update the percentage complete...
                        var progress = ++count * 100 / Interlocked.Read(ref total);
                        if (progress > percentage && progress <= 100)
                        {
                            percentage = progress;
                            Console.WriteLine($@"{percentage}% complete.");
                        }
                    }
                };

            task.Find();
            // probably task is completes
            //CheckBox1.IsEnabled = true;
            //BtnPath.IsEnabled = true;
            //LblPath.Text = "Checking: " + TargetPath;
            //Console.WriteLine("Enumerated {0:n0} files", total);
        }

        private void RefreshListView()
        {
            var files = CheckBox1.IsChecked.HasValue && CheckBox1.IsChecked.Value
            //? Directory.GetFiles(TargetPath, "*.max", SearchOption.AllDirectories)
            //: Directory.GetFiles(TargetPath, "*.max", SearchOption.TopDirectoryOnly);
            ? Directory.EnumerateFiles(TargetPath, "*.max", SearchOption.AllDirectories)
            : Directory.EnumerateFiles(TargetPath, "*.max", SearchOption.TopDirectoryOnly);

            var items = new List<LvItem>();
            foreach (var file in files)
            {
                var docInfo = ParseDocInfo(file);

                var warningInfo = (ValidateObjectsName(docInfo) > 0) ? "\u26A0" : "";
                var tip = (ValidateObjectsName(docInfo) > 0) ? "May have ALC scripts" : "";

                var verInfo = GetVersionInfo(docInfo.Values.ToArray()[0]);
                if (!verInfo.TryGetValue(0, out var verMax)) verMax = "???";
                if (!verInfo.TryGetValue(1, out var verSave)) verSave = "???";

                var item = new LvItem
                {
                    About = warningInfo,
                    AboutTooltip = tip,
                    FileName = Path.GetFileNameWithoutExtension(file),
                    FilePath = file,
                    Version = verMax,
                    SaveAsVersion = verSave,
                    SumInfo = null,
                    DocInfo = docInfo,
                };

                items.Add(item);
            }
            ListView1.ItemsSource = items;

            var view = (CollectionView) CollectionViewSource.GetDefaultView(ListView1.ItemsSource);
            view?.SortDescriptions.Add(new SortDescription("About", ListSortDirection.Descending));
            view?.SortDescriptions.Add(new SortDescription("FileName", ListSortDirection.Ascending));
        }

        private int ValidateObjectsName(Dictionary<string, Dictionary<int, string>> docInfo)
        {
            if (docInfo.Count <= 0) return 0;
            var objInfo = docInfo.ElementAt(4).Value.Values.ToArray();
            var alcNodes = new[]
            {
                "×þ×ü", "¡¡×ý×û", "¡¡¡¡", "¡¡¡¡¡¡", "¡¡¡¡¡¡¡¡", "¡¡¡¡¡¡¡¡¡¡¡¡", "¡¡¡¡¡¡¡¡¡¡¡¡¡¡", "Rectangles135", "×ú×ú", "×þ×ú", "×t×ü", "??×y×?", "　", ""
            };
            var alcHashSet = new HashSet<string>(alcNodes);
            return objInfo.Count(alcHashSet.Contains);
        }

        private static Dictionary<int, string> GetVersionInfo(Dictionary<int, string> obj)
        {
            var returnValue = new Dictionary<int, string>();

            var patternList = new Dictionary<string, string>
            {
                {"AppVerEN", "3ds Max Version"},
                {"FileVerEN", "Saved As Version"},
                {"AppVerCN", "3ds Max 版本"},
                {"FileVerCN", "另存为版本"},
                {"AppVerJA", "3ds Max バージョン"},
                {"FileVerJA", "バージョンとして保存"}
            };

            // search version info
            foreach (var k in obj)
            foreach (var sPattern in patternList)
            {
                if (!Regex.IsMatch(k.Value, sPattern.Value, RegexOptions.IgnoreCase)) continue;
                switch (sPattern.Key)
                {
                    case "AppVerEN":
                    case "AppVerCN":
                    case "AppVerJA":
                        var verA = k.Value.Split(' ').Last();
                        returnValue.Add(0, MaxString(verA));
                        returnValue.Add(1, "-");
                        break;
                    case "FileVerEN":
                    case "FileVerCN":
                    case "FileVerJA":
                        var verB = k.Value.Split(' ').Last();
                        returnValue[1] = MaxString(verB);
                        break;
                }
            }

            return returnValue;
        }

        public Dictionary<string, Dictionary<int, string>> ParseDocInfo(string file)
        {
            var headParts = new Dictionary<string, int>();
            var valueParts = new Dictionary<int, string>();

            var cf = new CompoundFile(file);
            var docInfo = cf.RootStorage.TryGetStream("\u0005DocumentSummaryInformation");
            if (docInfo != null)
            {
                var streamData = docInfo.GetData(); // get byte[] of document information
                var sec1Offset = BitConverter.ToInt32(streamData, 0x2c);
                var sec2Offset = BitConverter.ToInt32(streamData, 0x40); // End of Header address 0x44

                // get properties - get number of properties such as codepage(0x01), docParts(0x0d), headingPairs(0x0c), etc...
                var sec1Cnt = BitConverter.ToInt32(streamData, sec1Offset + 4);
                var docProperties = new Dictionary<int, int>();
                for (var i = 0; i < sec1Cnt; i++)
                {
                    var pid = BitConverter.ToInt32(streamData, sec1Offset + 8 + 8 * i); // 0x4c byte[4]  property id
                    var pOffset =
                        BitConverter.ToInt32(streamData, sec1Offset + 8 + 8 * i + 4); // 0x50 byte[4]  property offset
                    docProperties.Add(pid, pOffset);
                }

                // get CodePage property   wType[2] padding[2] value[2] unused[2]   = total 8 bytes
                const int codePagePid = 0x01;
                docProperties.TryGetValue(codePagePid, out var codePageOffset); // 0x44 + Offset
                var codePageValue =
                    BitConverter.ToUInt16(streamData, sec1Offset + codePageOffset + 4); // 1200 = utf-16, 65001 = utf-8

                // get HeadingPair property
                const int headPairPid = 0x0c; // BitConverter.ToInt32(sd, sec1Offset + 24); // 0x0c
                docProperties.TryGetValue(headPairPid,
                    out var headPairOffset); // BitConverter.ToInt32(sd, sec1Offset + 28);
                var headPairElements = headPairOffset > 0
                    ? BitConverter.ToInt32(streamData, sec1Offset + headPairOffset + 4)
                    : 0;
                var headPairArrayOffset = headPairOffset > 0 ? sec1Offset + headPairOffset + 8 : 0;

                // get DocParts property
                const int docPartsPid = 0x0d; // BitConverter.ToInt32(sd, sec1Offset + 32); // 0x0d
                docProperties.TryGetValue(docPartsPid,
                    out var docPartsOffset); // BitConverter.ToInt32(sd, sec1Offset + 36);
                var docPartsElements = docPartsOffset > 0
                    ? BitConverter.ToInt32(streamData, sec1Offset + docPartsOffset + 4)
                    : 0;
                var docPartsArrayOffset =
                    docPartsOffset > 0
                        ? sec1Offset + docPartsOffset + 8
                        : 0; // skip byte[2] wType, byte[2] padding, byte[4] cElements

                // start parsing heading pairs
                try
                {
                    // parse heading pairs
                    var buf = streamData.Skip(headPairArrayOffset).Take(streamData.Length - headPairArrayOffset)
                        .ToArray(); // feel lazy to calc, putting all
                    var pos = 0;
                    for (var i = 0;
                        i < headPairElements / 2;
                        i++) // cElements is always twice the number of elements in rgHeadingPairs because each VtHeadingPair (section 2.3.3.1.13) element contains two values, headingString and headerParts.
                    {
                        // VtHeadingPairElement
                        // headingString
                        var stringType = BitConverter.ToInt16(buf, pos);
                        pos += 2;
                        var headStrPadding = BitConverter.ToInt16(buf, pos);
                        pos += 2;
                        var cch = BitConverter.ToInt32(buf, pos);
                        if (cch % 4 != 0) cch += cch % 4; // treat string paddings
                        pos += 4;
                        var strValue = Encoding.GetEncoding(codePageValue).GetString(buf, pos, cch)
                            .TrimEnd(char.MinValue);
                        pos += cch;
                        // headerParts
                        var wType = BitConverter.ToInt16(buf, pos);
                        pos += 2;
                        var headPrtPadding = BitConverter.ToInt16(buf, pos);
                        pos += 2;
                        var partsValue = BitConverter.ToInt32(buf, pos);
                        pos += 4;
                        headParts.Add(strValue, partsValue);
                    }
                }
                catch (Exception ex)
                {
                    Console.WriteLine(ex.ToString());
                }

                // start parsing document information 
                try
                {
                    // parse document information
                    var buf = streamData.Skip(docPartsArrayOffset).Take(sec2Offset - docPartsArrayOffset).ToArray();
                    var pos = 0;
                    for (var i = 0; i < docPartsElements; i++) // i < docPartsElements; // for complete parsing
                    {
                        var cch = BitConverter.ToInt32(buf, pos);
                        if (cch % 4 != 0) cch += cch % 4; // treat string paddings
                        pos += 4;
                        var value = Encoding.GetEncoding(codePageValue).GetString(buf, pos, cch).TrimEnd(char.MinValue);
                        pos += cch;
                        valueParts.Add(i, value);
                    }
                }
                catch (Exception ex)
                {
                    Console.WriteLine(ex.ToString());
                }
            }
            cf.Close();

            var returnValue = new Dictionary<string, Dictionary<int, string>>();
            var idx = 0;
            foreach (var entry in headParts)
            {
                var temp = new Dictionary<int, string>();

                for (var i = idx; i < idx + entry.Value; i++)
                    temp.Add(valueParts.Keys.ToArray()[i], valueParts.Values.ToArray()[i]);

                returnValue.Add(entry.Key, temp);
                idx += entry.Value;
            }
            return returnValue;
        }

        public static string MaxString(string ver)
        {
            var maxVer = ver;
            switch (ver)
            {
                case "1.0":
                    maxVer = "1";
                    break;
                case "2.0":
                    maxVer = "2";
                    break;
                case "3.0":
                    maxVer = "3";
                    break;
                case "4.0":
                    maxVer = "4";
                    break;
                case "5.0":
                    maxVer = "5";
                    break;
                case "6.0":
                    maxVer = "6";
                    break;
                case "7.0":
                    maxVer = "7";
                    break;
                case "8.0":
                    maxVer = "8";
                    break;
                case "9.0":
                case "9.00":
                    maxVer = "9";
                    break;
                case "10.0":
                case "10.00":
                    maxVer = "2008";
                    break;
                case "11.0":
                case "11.00":
                    maxVer = "2009";
                    break;
                case "12.0":
                case "12.00":
                    maxVer = "2010";
                    break;
                case "13.00":
                    maxVer = "2011";
                    break;
                case "14.00":
                    maxVer = "2012";
                    break;
                case "15.00":
                    maxVer = "2013";
                    break;
                case "16.00":
                    maxVer = "2014";
                    break;
                case "17.00":
                    maxVer = "2015";
                    break;
                case "18.00":
                    maxVer = "2016";
                    break;
                case "19.00":
                    maxVer = "2017";
                    break;
                case "20.00":
                    maxVer = "2018";
                    break;
                case "21.00":
                    maxVer = "2019";
                    break;
                case "22.00":
                    maxVer = "2020";
                    break;
            }
            return maxVer;
        }

        private void BtnOpen_OnClick(object sender, RoutedEventArgs e)
        {
            var item = GetAncestorOfType<ListViewItem>(sender as Button);
            if (item.Content is LvItem filePath)
            {
                if (filePath.FilePath != null)
                {
                    Process.Start(Path.GetDirectoryName(filePath.FilePath) ?? throw new InvalidOperationException());
                }
            }
        }

        private void BtnInfo_OnClick(object sender, RoutedEventArgs e)
        {
            if (!(GetAncestorOfType<ListViewItem>(sender as Button).Content is LvItem item)) return;
            var modalWindow = new ModalWindow
            {
                Title = item.FileName + ".max",
                Owner = Application.Current.MainWindow,
                WindowStartupLocation = WindowStartupLocation.CenterOwner
            };

            var tb = modalWindow.MdlTxtBlck;
            tb.Text = string.Empty;

            foreach (var entry in item.DocInfo)
            {
                tb.Inlines.Add(new Run(entry.Key + "\n")
                {
                    FontWeight = FontWeights.Bold,
                    Foreground = Brushes.DarkSlateGray
                });

                foreach (var innerEntry in entry.Value)
                    tb.Inlines.Add("  " + innerEntry.Value + "\n");
            }
            modalWindow.ShowDialog();
        }

        public T GetAncestorOfType<T>(FrameworkElement child) where T : FrameworkElement
        {
            var parent = VisualTreeHelper.GetParent(child);
            if (parent != null && !(parent is T))
                return GetAncestorOfType<T>((FrameworkElement) parent);
            return (T) parent;
        }

        private class LvItem
        {
            public string About { get; set; }
            public string AboutTooltip { get; set; }
            public string FileName { get; set; }
            public string FilePath { get; set; }
            public string Version { get; set; }
            public string SaveAsVersion { get; set; }

            public Dictionary<string, string> SumInfo { get; set; }
            public Dictionary<string, Dictionary<int, string>> DocInfo { get; set; }
        }
    }
}