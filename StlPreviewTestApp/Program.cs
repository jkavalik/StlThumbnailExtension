using System;
using System.Drawing;
using System.IO;
using StlThumbnailExtension;

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
                using (var fs = File.OpenRead(stlFile))
                {
                    var model = StlModel.Load(fs);
                    using (var bmp = StlRenderer.RenderToBitmap(model, size, size))
                    {
                        bmp.Save(pngFile, System.Drawing.Imaging.ImageFormat.Png);
                    }
                }
                Console.WriteLine("Preview generated: " + pngFile);
            }
            catch (Exception ex)
            {
                Console.WriteLine("Error generating preview: " + ex.Message);
            }
        }
    }
}