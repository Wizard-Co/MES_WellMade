using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Controls.Primitives;
using System.Windows.Data;
using System.Windows.Media;

namespace WizMes_WellMade
{
    /// <summary>
    /// Represents a control that displays hierarchical data in a tree structure
    /// that has items that can expand and collapse.
    /// </summary>
    public class TSCLIB_DLL
    {
        [DllImport("TSCLIB.dll", EntryPoint = "about")]
        public static extern int about();

        [DllImport("TSCLIB.dll", EntryPoint = "openport")]
        public static extern int openport(string printername);

        [DllImport("TSCLIB.dll", EntryPoint = "barcode")]
        public static extern int barcode(string x, string y, string type,
                    string height, string readable, string rotation,
                    string narrow, string wide, string code);

        [DllImport("TSCLIB.dll", EntryPoint = "clearbuffer")]
        public static extern int clearbuffer();

        [DllImport("TSCLIB.dll", EntryPoint = "closeport")]
        public static extern int closeport();

        [DllImport("TSCLIB.dll", EntryPoint = "downloadpcx")]
        public static extern int downloadpcx(string filename, string image_name);

        [DllImport("TSCLIB.dll", EntryPoint = "formfeed")]
        public static extern int formfeed();

        [DllImport("TSCLIB.dll", EntryPoint = "nobackfeed")]
        public static extern int nobackfeed();

        [DllImport("TSCLIB.dll", EntryPoint = "printerfont")]
        public static extern int printerfont(string x, string y, string fonttype,
                        string rotation, string xmul, string ymul,
                        string text);

        [DllImport("TSCLIB.dll", EntryPoint = "printlabel")]
        public static extern int printlabel(string set, string copy);

        [DllImport("TSCLIB.dll", EntryPoint = "sendcommand")]
        public static extern int sendcommand(string printercommand);

        [DllImport("TSCLIB.dll", EntryPoint = "setup")]
        public static extern int setup(string width, string height,
                  string speed, string density,
                  string sensor, string vertical,
                  string offset);

        [DllImport("TSCLIB.dll", EntryPoint = "windowsfont")]
        public static extern int windowsfont(int x, int y, int fontheight,
                        int rotation, int fontstyle, int fontunderline,
                        string szFaceName, string content);
    }

    public class EnumItem
    {
        public static int IO_DATA = 0;
        public static int IO_BARCODE = 1;
        public static int IO_TEXT = 2;
        public static int IO_LINE = 3;
        public static int IO_RECT = 4;
        public static int IO_DIAMOND = 5;
        public static int IO_CIRCLE = 6;
        public static int IO_IMAGE = 7;
        public static int IO_QRcode = 8;
        public static int IO_BOX = 9;
    }

    public class mt_Tag_CodeView
    {
        public string TagID { get; set; }
        public string Tag { get; set; }
        public int Width { get; set; }
        public int Height { get; set; }
        public int DefHeight { get; set; }

        public int DefBaseY { get; set; }
        public int DefBaseX1 { get; set; }
        public int DefBaseX2 { get; set; }
        public int DefBaseX3 { get; set; }
        public int DefGapY { get; set; }

        public int DefGapX1 { get; set; }
        public int DefGapX2 { get; set; }
        public int DefLength { get; set; }
        public int DefHCount { get; set; }
        public int DefBarClss { get; set; }

        public int Gap { get; set; }
        public string Direct { get; set; }
    }

    public class mt_TagSub_CodeView
    {
        public string Name { get; set; }
        public int Type { get; set; }
        public int Align { get; set; }
        public int x { get; set; }
        public int y { get; set; }

        public int Font { get; set; }
        public int Length { get; set; }
        public int HMulti { get; set; }
        public int VMulti { get; set; }
        public int Relation { get; set; }

        public int Rotation { get; set; }
        public int Space { get; set; }
        public int PrevItem { get; set; }
        public int BarType { get; set; }
        public int BarHeight { get; set; }

        public int FigureWidth { get; set; }
        public int FigureHeight { get; set; }
        public int Thickness { get; set; }
        public string ImageFile { get; set; }
        public int Width { get; set; }

        public int Height { get; set; }
        public int Visible { get; set; }
        public string FontName { get; set; }
        public string FontStyle { get; set; }
        public string FontUnderLine { get; set; }

        public string Text { get; set; }
        public int BuyerArticle { get; internal set; }
    }
}
