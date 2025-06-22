using System;
using System.Drawing;
using System.IO;
using StlThumbnailExtension;
using System.Diagnostics;

namespace StlPreviewTestApp
{
    class Program
    {
        static void Main(string[] args)
        {
            if (args.Length < 2)
            {
                Console.WriteLine("Usage: StlPreviewTestApp <input.stl> <output.png> [size]");
                return;
            }

            string stlFile = args[0];
            string pngFile = args[1];
            int size = 256;
            if (args.Length > 2)
                int.TryParse(args[2], out size);

            if (!File.Exists(stlFile))
            {
                Console.WriteLine("Error: File not found: " + stlFile);
                return;
            }

            try
            {
                // --- Use new direct streaming Z-buffer+GDI renderer ---
                Bitmap bmp;
                var sw = Stopwatch.StartNew();
                using (var fs = File.OpenRead(stlFile))
                {
                    bmp = StlRenderer.RenderDirectlyFromStream(fs, size, size);
                }
                sw.Stop();

                bmp.Save(pngFile, System.Drawing.Imaging.ImageFormat.Png);
                bmp.Dispose();

                Console.WriteLine($"New direct streaming Z-buffer+GDI preview generated: {pngFile} in {sw.ElapsedMilliseconds} ms");
            }
            catch (Exception ex)
            {
                Console.WriteLine("Error generating preview: " + ex.Message);
            }
        }
    }
}