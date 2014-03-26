using System;
using System.IO;
using System.Collections.Generic;
using System.Drawing;
using System.Drawing.Drawing2D;
using System.Linq;
using System.Threading;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Runtime.InteropServices;
using Microsoft.Win32;

using System.Diagnostics;

namespace Curator
{
    class WallpaperChanger
    {
        private string _path;
        static Random randGen;

        public WallpaperChanger()
        {
            _path = null;
            randGen = new Random();
        }

        public string path { get { return _path; } set { _path = value; } }

        public void ChangeWallpaper()
        {
            if (path != null)
            {
                List<string> images = new List<string>();
                var filters = new String[] { "jpg", "jpeg", "png", "gif", "tiff", "bmp" };
                foreach (var filter in filters)
                {
                    images.AddRange(System.IO.Directory.GetFiles(path, String.Format("*.{0}", filter), System.IO.SearchOption.AllDirectories));
                }

                string fileName = images[randGen.Next(0, images.Count)];

                // Temporary resizing method... essentially just preserving quality with Fit

                float width = Screen.PrimaryScreen.Bounds.Width;
                float height = Screen.PrimaryScreen.Bounds.Height;
                var brush = new SolidBrush(Color.Black);

                Bitmap image = new Bitmap(fileName);
                Bitmap bmp = new Bitmap((int)width, (int)height);
                float scale = Math.Min(width / image.Width, height / image.Height);
                var graph = Graphics.FromImage(bmp);

                graph.InterpolationMode = InterpolationMode.High;
                graph.CompositingQuality = CompositingQuality.HighQuality;
                graph.SmoothingMode = SmoothingMode.AntiAlias;

                var scaleWidth = (int)(image.Width * scale);
                var scaleHeight = (int)(image.Height * scale);

                graph.FillRectangle(brush, new RectangleF(0, 0, width, height));
                graph.DrawImage(image, new Rectangle(((int)width - scaleWidth) / 2, ((int)height - scaleHeight) / 2, scaleWidth, scaleHeight));
                graph.Dispose();

                if (!Directory.Exists(Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData), @"Curator\temp")))
                {
                    Directory.CreateDirectory(Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData), @"Curator\temp"));
                }

                try
                {
                    bmp.Save(Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData), @"Curator\temp\temp.bmp"), System.Drawing.Imaging.ImageFormat.Bmp);
                }
                catch (Exception e)
                {
                    Debug.WriteLine("DEBUG::ChangeWallpaper()::Error attempting to save temporary image::" + e.Message);
                }

                // Clean up clean up everybody everywhere
                image.Dispose();
                bmp.Dispose();

                IntPtr result = IntPtr.Zero;
                WinAPI.SendMessageTimeout(WinAPI.FindWindow("Progman", null), 0x52c, IntPtr.Zero, IntPtr.Zero, 0, 500, out result);

                ThreadStart threadStarter = () =>
                {
                    WinAPI.IActiveDesktop _activeDesktop = WinAPI.ActiveDesktopWrapper.GetActiveDesktop();
                    _activeDesktop.SetWallpaper(Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData), @"Curator\temp\temp.bmp"), 0);
                    _activeDesktop.ApplyChanges(WinAPI.AD_Apply.ALL | WinAPI.AD_Apply.FORCE);

                    Marshal.ReleaseComObject(_activeDesktop);
                };
                Thread thread = new Thread(threadStarter);
                thread.SetApartmentState(ApartmentState.STA); //Set the thread to STA (REQUIRED!!!!)
                thread.Start();
                thread.Join(2000);
            }
        }
    }
}
