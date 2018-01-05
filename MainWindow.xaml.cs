using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Windows;
using Ookii.Dialogs.Wpf;
using OpenMcdf;

namespace _3DSMaxFileVersion
{
    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {
        public MainWindow()
        {
            InitializeComponent();
        }

        public string Path;

        private void BtnPath_Click(object sender, RoutedEventArgs e)
        {
            var dialog = new VistaFolderBrowserDialog();
            if (!dialog.ShowDialog(this).GetValueOrDefault()) return;
            Path = dialog.SelectedPath;
            LblPath.Text = Path;

            ListView1.Items.Clear();
            var files = (CheckBox1.IsChecked.HasValue && CheckBox1.IsChecked.Value)
                ? Directory.GetFiles(Path, "*.max", SearchOption.AllDirectories)
                : Directory.GetFiles(Path, "*.max", SearchOption.TopDirectoryOnly);

            foreach (var filename in files)
            {
                ListView1.Items.Add(new LvItem { FileName = System.IO.Path.GetFileNameWithoutExtension(filename), Version = "", SaveAsVersion = "" });
            }
        }

        private class LvItem
        {
            public string FileName { get; set; }
            public string Version { get; set; }
            public string SaveAsVersion { get; set; }
        }

        private void BtnScan_OnClick(object sender, RoutedEventArgs e)
        {
            ListView1.Items.Clear();
            var files = (CheckBox1.IsChecked.HasValue && CheckBox1.IsChecked.Value)
                ? Directory.GetFiles(Path, "*.max", SearchOption.AllDirectories)
                : Directory.GetFiles(Path, "*.max", SearchOption.TopDirectoryOnly);

            foreach (var file in files)
            {
                // OpenMCDF is a 100% managed .net component that allows client applications to manipulate COM structured storage files, also known as Microsoft Compound Document Format files.
                // Structured storage file is a container which include multiple streams of information.
                //
                //
                // 3ds Max's stream item names
                // Scene
                // Config
                // ClassData
                // DllDirectory
                // SaveConfigData
                // VideoPostQueue
                // ClassDirectory3
                // FileAssetMetaData3	
                // \u0005SummaryInformation
                // ScriptedCustAttribDefs	
                // \u0005DocumentSummaryInformation

                var cf = new CompoundFile(file);
                //Action<CFItem> va = delegate (CFItem item)
                //{
                //    Console.WriteLine("Name: {0, 30}\tIsRoot {1,5}\tIsStorage {2,5}\tIsStream {3,5}", item.Name, item.IsRoot, item.IsStorage, item.IsStream);
                //};
                //cf.RootStorage.VisitEntries(va, true);

                //CFStream docInfo = cf.RootStorage.TryGetStream("Config");
                var docInfo = cf.RootStorage.TryGetStream("\u0005DocumentSummaryInformation");

                if (docInfo != null)
                {
                    byte[] sd = docInfo.GetData();

                    // DocumentSummaryInfoStream header start
                    // 0x00 byte[2]  byteOrder        0xFFFE
                    // 0x02 byte[2]  version          0x0000
                    // 0x04 byte[1]  OSMajorVersion
                    // 0x05 byte[1]  OSMinorVersion
                    // 0x06 byte[2]  OSType
                    // 0x08 byte[10] applicationClsid
                    // 0x18 byte[4]  cSections        0x00000002
                    // 0x1c byte[10] FMTID_DocSummaryInformation 
                    var sec1Guid = new Guid(sd.Skip(0x1c).Take(0x10).ToArray()); // 02 D5 CD D5 9C 2E 1B 10 93 97 08 00 2B 2C F9 AE
                    // 0x2c byte[4]  sectionOffset
                    var sec1Offset = BitConverter.ToInt32(sd, 0x2c);
                    // 0x30 byte[10] FMTID_UserDefinedProperties 
                    var sec2Guid = new Guid(sd.Skip(0x30).Take(0x10).ToArray()); // 05 D5 CD D5 9C 2E 1B 10 93 97 08 00 2B 2C F9 AE
                    // 0x40 byte[4]  sectionOffset
                    var sec2Offset = BitConverter.ToInt32(sd, 0x40); // End of Header 0x44

                    // FMTID_DocSummaryInformation property set start
                    // 0x44 byte[4]  cbSection      total length of first section
                    var sec1Len = BitConverter.ToInt32(sd, sec1Offset);
                    // 0x48 byte[4]  cProps         number of properties
                    var sec1Cnt = BitConverter.ToInt32(sd, sec1Offset + 4); // 0x04; as far as allways 4 properties - codepage(0x01), unknown?(0x80000000), docParts(0x0d), headingPairs(0x0c)

                    var docProperties = new Dictionary<int, int>();
                    for (var i = 0; i < sec1Cnt; i++)
                    {
                        // 0x4c byte[4]  property id
                        // 0x50 byte[4]  property offset
                        var pid = BitConverter.ToInt32(sd, sec1Offset + 8 + 8 * i);
                        var poffset = BitConverter.ToInt32(sd, sec1Offset + 8 + 8 * i + 4);
                        docProperties.Add(pid, poffset);
                        
                    }
                    // CodePage Property   wType[2] padding[2] value[2] unused[2]   = total 8 bytes
                    var codePagePid = 0x01; // BitConverter.ToInt32(sd, sec1Offset +  8); // 0x1
                    docProperties.TryGetValue(codePagePid, out var codePageOffset); // BitConverter.ToInt32(sd, sec1Offset + 12); // 0x44 + Offset
                    var codePageType = BitConverter.ToInt16(sd, sec1Offset + codePageOffset);
                    var codePageValue = BitConverter.ToUInt16(sd, sec1Offset + codePageOffset + 4); // 1200 = utf-16, 65001 = utf-8
                    //
                    var unknownPid = 0x80000000;  // BitConverter.ToInt32(sd, sec1Offset + 16); // 0x80000000
                    var unknownOffset = BitConverter.ToInt32(sd, sec1Offset + 20);
                    var UnknownArrayOffset = sec1Offset + unknownOffset;
                    //
                    var headPairPid = 0x0c; // BitConverter.ToInt32(sd, sec1Offset + 24); // 0x0c
                    docProperties.TryGetValue(headPairPid, out var headPairOffset); // BitConverter.ToInt32(sd, sec1Offset + 28);
                    var headPairElements = 0;
                    var headPairArrayOffset = 0;
                    if (headPairOffset > 0)
                    {
                        headPairElements = BitConverter.ToInt32(sd, sec1Offset + headPairOffset + 4);
                        headPairArrayOffset = sec1Offset + headPairOffset + 8;
                    }
                    //
                    var docPartsPID = 0x0d; // BitConverter.ToInt32(sd, sec1Offset + 32); // 0x0d
                    docProperties.TryGetValue(docPartsPID, out var docPartsOffset); // BitConverter.ToInt32(sd, sec1Offset + 36);
                    var docPartsElements = 0;
                    var docPartsArrayOffset = 0;
                    if (docPartsOffset > 0)
                    {
                        docPartsElements = BitConverter.ToInt32(sd, sec1Offset + docPartsOffset + 4);
                        docPartsArrayOffset = sec1Offset + docPartsOffset + 8; // skip byte[2] wType, byte[2] padding, byte[4] cElements
                    }

                    // parsing test - headingPair property header and the first element
                    //var wType                  = BitConverter.ToInt16(sd, sec1Offset + headPairOffset + 0); //new byte[2];
                    //var padding                = BitConverter.ToInt16(sd, sec1Offset + headPairOffset + 2); //new byte[2];
                    //var headingPairElments     = BitConverter.ToInt32(sd, sec1Offset + headPairOffset + 4); //new byte[4];
                    //var headingStringType      = BitConverter.ToInt16(sd, sec1Offset + headPairOffset + 8); //new byte[2];
                    //var headingStringPadding   = BitConverter.ToInt16(sd, sec1Offset + headPairOffset + 10); //new byte[2];
                    //var headingStringValueCch  = BitConverter.ToInt32(sd, sec1Offset + headPairOffset + 12); //new byte[4];
                    //var headingStringArray     = sd.Skip(sec1Offset + headPairOffset + 16).Take(headingStringValueCch).ToArray();
                    //var headingString          = Encoding.GetEncoding(CodePageValue).GetString(headingStringArray).TrimEnd(Char.MinValue); //new byte[headingStringValueCch] //GetEncoding(932) Japanese?
                    //var headerPartsWType       = BitConverter.ToInt16(sd, sec1Offset + headPairOffset + 16 + headingStringValueCch); //new byte[2];
                    //var headerPartsPadding     = BitConverter.ToInt16(sd, sec1Offset + headPairOffset + 18 + headingStringValueCch); //new byte[2];
                    //var headerPartsValue       = BitConverter.ToInt32(sd, sec1Offset + headPairOffset + 22 + headingStringValueCch); //new byte[4];

                    int section2Length = BitConverter.ToInt32(sd, sec2Offset);
                    int section2Count = BitConverter.ToInt32(sd, sec2Offset + 0x04);

                    Console.WriteLine("\nName: {0}\tStreamLength: 0x{1:X4}({1,4})\tHeader Length: 0x{2:X4}({2,4})", System.IO.Path.GetFileName(file), sd.Length, 0x44);
                    //Console.WriteLine("Fingerprint: {0:x8}", BitConverter.ToString(sd, 0x00, 0x08).Replace("-", " "));
                    //Console.WriteLine("FMTID_DocSummaryInformation GUID1 {0}\tOffset1 0x{1:X4}({1,4})\tLength1 0x{2:X4}({2,4})\tCount1 0x{3:X4}({3,4})", sec1Guid, sec1Offset, sec1Len, sec1Cnt);
                    //Console.WriteLine("Fingerprint: {0:x8}", BitConverter.ToString(sd, sec1Offset, 0x08).Replace("-", " "));
                    Console.WriteLine("CodePage     (0x1): 0x{0:x8}({0,2})\t0x{1:x4}({1,4})\tCodePageType: 0x{2:x2}({2,2})\tCodePageValue: 0x{3:x2}({3,2})", codePagePid, codePageOffset, codePageType, codePageValue);
                    //Console.WriteLine("Unknown             0x{0:x8}(  )\t0x{1:x4}({1,4})\t0x{2:x4}({2,4})", UnknownPID, UnknownOffset, UnknownArrayOffset);
                    //Console.WriteLine("HeadingPairs (0xC): 0x{0:x8}({0,2})\t0x{1:x4}({1,4})\tType: 0x{2:x4}({2,4})", headPairPID, headPairOffset, headPairArrayOffset);
                    //Console.WriteLine("Number of Elements: 0x{0:x4}({0,4}) @ 0x{1:x4}({1,4})", headingPairElments / 2, sec1Offset + headPairOffset + 4);
                    Console.WriteLine("DocumentParts(0xD): 0x{0:x8}({0,2})\t0x{1:x4}({1,4})\t0x{2:x4}({2,4})", docPartsPID, docPartsOffset, docPartsArrayOffset);
                    //Console.WriteLine("FMTID_UserDefinedProperties GUID2 {0}\tOffset2 0x{1:X4}({1,4})\tLength2 0x{2:X4}({2,4})\tCount2 0x{3:X4}({3,4})", sec2Guid, sec2Offset, section2Length, section2Count);
                    //Console.WriteLine("HeadingString: Type {0:x2}, Pad {1:x2}, Length {2:x4}, String {3} @ {4:x4}",
                    //    headingStringType, headingStringPadding, headingStringValueCch, headingString, sec1Offset + headPairOffset + 16);
                    //Console.WriteLine("HeaderParts: Type {0:x2}, Pad {1:x2}, Count {2:x4}",
                    //    headerPartsWType, headerPartsPadding, headerPartsValue);


                    // start parsing document information 
                    try
                    {
                        byte[] buf = sd.Skip(docPartsArrayOffset).Take(sec2Offset - docPartsArrayOffset).ToArray();
                        var pos = 0;
                        var lpstrElement = new List<KeyValuePair<int, string>>();
                        for (var i = 0; i < docPartsElements; i++) // i < docPartsElements; // for complete parsing
                        {
                            var cch = BitConverter.ToInt32(buf, pos);
                            if (cch % 4 != 0) cch += cch % 4;
                            var value = Encoding.GetEncoding(codePageValue).GetString(buf, pos + 4, cch).TrimEnd(Char.MinValue);
                            lpstrElement.Add(new KeyValuePair<int, string>(i, value));
                            pos += 4 + cch;
                        }

                        Console.WriteLine("docPart successfully parsed!!!");

                        var patternList = new Dictionary<string, string>
                        {
                            { "AppVerEN", "3ds Max Version" },
                            { "FileVerEN", "Saved As Version" },
                            { "AppVerCN", "3ds Max 版本" },
                            { "FileVerCN", "另存为版本" },
                            { "AppVerJA", "3ds Max バージョン" },
                            { "FileVerJA", "バージョンとして保存" }
                        };

                        var verA = "-";
                        var verB = "-";

                        foreach (var sPattern in patternList)
                        {
                            foreach (var k in lpstrElement)
                            {
                                if (Regex.IsMatch(k.Value, sPattern.Value, RegexOptions.IgnoreCase))
                                {
                                    if (sPattern.Key == "AppVerEN" || sPattern.Key == "AppVerCN" || sPattern.Key == "AppVerJA")
                                    {
                                        verA = k.Value.Split(' ').Last();
                                        Console.WriteLine("3ds Max Version: {0}", verA);
                                    }
                                    else if (sPattern.Key == "FileVerEN" || sPattern.Key == "FileVerCN" || sPattern.Key == "FileVerJA")
                                    {
                                        verB = k.Value.Split(' ').Last();
                                        Console.WriteLine("Saved As Version: {0}", verA);
                                    }
                                }
                            }
                        }
                        ListView1.Items.Add(new LvItem { FileName = System.IO.Path.GetFileNameWithoutExtension(file), Version = MaxString(verA), SaveAsVersion = MaxString(verB) });
                    }
                    catch (Exception ex)
                    {
                        Console.WriteLine(ex.ToString());
                    }
                }
                else
                {
                    Console.WriteLine("Can't get steam data!!!!");
                };
                cf.Close();
            }
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
            }
            return maxVer;
        }
    }
}
