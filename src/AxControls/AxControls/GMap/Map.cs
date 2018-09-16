using System.Windows.Forms;

using System.Drawing;
using System;
using System.Globalization;
using System.Runtime.InteropServices;

using System.Collections.Generic;
using AxControls;
using GMap.NET.WindowsForms;

namespace AxControls
{
    [ComVisible(true)]
    [ClassInterface(ClassInterfaceType.None)]
    public class Map : GMapControl
    {
        //public new GMapProvider MapProvider;

        public Map()
        {
            InitializeComponent();
        }

        private void InitializeComponent()
        {
            this.SuspendLayout();
            // 
            // Map
            // 
            this.BackColor = System.Drawing.SystemColors.Control;
            this.EmptyTileColor = System.Drawing.Color.LightSkyBlue;
            this.Name = "Map";
            this.Size = new System.Drawing.Size(279, 288);
            this.ResumeLayout(false);
        }

        //#if DEBUG
        //      private int counter;
        //      readonly Font DebugFont = new Font(FontFamily.GenericSansSerif, 14, FontStyle.Regular);
        //      readonly Font DebugFontSmall = new Font(FontFamily.GenericSansSerif, 12, FontStyle.Bold);
        //      DateTime start;
        //      DateTime end;
        //      int delta;

        //      protected override void OnPaint(PaintEventArgs e)
        //      {
        //         start = DateTime.Now;

        //         base.OnPaint(e);

        //         end = DateTime.Now;
        //         delta = (int)(end - start).TotalMilliseconds;
        //      }

        //      /// <summary>
        //      /// any custom drawing here
        //      /// </summary>
        //      /// <param name="drawingContext"></param>
        //      protected override void OnPaintOverlays(System.Drawing.Graphics g)
        //      {
        //         base.OnPaintOverlays(g);

        //         g.DrawString(string.Format(CultureInfo.InvariantCulture, "{0:0.0}", Zoom) + "z, " + MapProvider + ", refresh: " + counter++ + ", load: " + ElapsedMilliseconds + "ms, render: " + delta + "ms", DebugFont, Brushes.Blue, DebugFont.Height, DebugFont.Height + 20);

        //         //g.DrawString(ViewAreaPixel.Location.ToString(), DebugFontSmall, Brushes.Blue, DebugFontSmall.Height, DebugFontSmall.Height);

        //         //string lb = ViewAreaPixel.LeftBottom.ToString();
        //         //var lbs = g.MeasureString(lb, DebugFontSmall);
        //         //g.DrawString(lb, DebugFontSmall, Brushes.Blue, DebugFontSmall.Height, Height - DebugFontSmall.Height * 2);

        //         //string rb = ViewAreaPixel.RightBottom.ToString();
        //         //var rbs = g.MeasureString(rb, DebugFontSmall);
        //         //g.DrawString(rb, DebugFontSmall, Brushes.Blue, Width - rbs.Width - DebugFontSmall.Height, Height - DebugFontSmall.Height * 2);

        //         //string rt = ViewAreaPixel.RightTop.ToString();
        //         //var rts = g.MeasureString(rb, DebugFontSmall);
        //         //g.DrawString(rt, DebugFontSmall, Brushes.Blue, Width - rts.Width - DebugFontSmall.Height, DebugFontSmall.Height);
        //      }     
        //#endif
    }
}
