using System;
using System.Drawing;
using System.Drawing.Imaging;
using System.Drawing.Text;
using System.IO;
using System.Runtime.InteropServices;
using SharpShell.Attributes;
using SharpShell.SharpThumbnailHandler;
using System.Collections.Generic;
using System.Globalization;
using System.Text;

namespace StlThumbnailExtension
{
    [ComVisible(true)]
    [COMServerAssociation(AssociationType.FileExtension, ".stl")]
    public class StlThumbnailHandler : SharpThumbnailHandler
    {
        protected override Bitmap GetThumbnailImage(uint width)
        {
            try
            {
                using (var stream = new BufferedStream(SelectedItemStream))
                {
                    return StlRenderer.RenderDirectlyFromStream(stream, (int)width, (int)width);
                }
            }
            catch (Exception exception)
            {
                LogError("An exception occurred generating the thumbnail.", exception);
                return null;
            }
        }
    }

    public class StlModel
    {
        public struct Vec3
        {
            public float X, Y, Z;
            public Vec3(float x, float y, float z) { X = x; Y = y; Z = z; }
            public static Vec3 operator +(Vec3 a, Vec3 b) => new Vec3(a.X + b.X, a.Y + b.Y, a.Z + b.Z);
            public static Vec3 operator /(Vec3 a, float f) => new Vec3(a.X / f, a.Y / f, a.Z / f);
        }
        public struct Triangle
        {
            public Vec3 Normal;
            public Vec3 V1, V2, V3;
        }
    }

    public static class StlRenderer
    {
        private const int maxTriangles = 500000;
        private struct ProjectedPoint
        {
            public PointF Point;
            public float Z;
            public ProjectedPoint(PointF point, float z)
            {
                Point = point;
                Z = z;
            }
        }

        /// <summary>
        /// Render STL directly from stream, in a single pass, without storing the mesh in memory.
        /// Works for both binary and ASCII STL. 
        /// </summary>
        public static Bitmap RenderDirectlyFromStream(Stream stream, int width, int height)
        {
            // --- First pass: scan bounds (min/max) ---
            // We need to scan the mesh once for bounds, then again for rendering.
            // (Alternatively, estimate bounds or use a 2-pass approach. Here, we use 2 pass for simplicity.)
            StlBounds bounds = ScanStlBounds(stream);

            // --- Setup projection math ---
            var center = new StlModel.Vec3(
                (bounds.MinX + bounds.MaxX) / 2,
                (bounds.MinY + bounds.MaxY) / 2,
                (bounds.MinZ + bounds.MaxZ) / 2
            );
            float scale = Math.Min(width, height) * 0.8f / Math.Max(
                bounds.MaxX - bounds.MinX,
                Math.Max(bounds.MaxY - bounds.MinY, bounds.MaxZ - bounds.MinZ)
            );
            float angle = (float)(Math.PI / 6);
            float cosA = (float)Math.Cos(angle), sinA = (float)Math.Sin(angle);
            float halfA = angle / 2f;
            float cosH = (float)Math.Cos(halfA), sinH = (float)Math.Sin(halfA);

            Func<StlModel.Vec3, ProjectedPoint> projectWithViewZ = v =>
            {
                float x = v.X - center.X, y = v.Y - center.Y, z = v.Z - center.Z;
                float x1 = x * cosA - z * sinA;
                float z1 = x * sinA + z * cosA;
                float y1 = y * cosH - z1 * sinH;
                float z2 = y * sinH + z1 * cosH;
                return new ProjectedPoint(
                    new PointF(
                        width / 2 + x1 * scale,
                        height / 2 - y1 * scale
                    ),
                    z2
                );
            };

            // --- Second pass: render with z-buffer, no mesh storage ---
            var bmp = new Bitmap(width, height, PixelFormat.Format32bppArgb);
            var zbuf = new float[width, height];
            for (int x = 0; x < width; x++)
                for (int y = 0; y < height; y++)
                    zbuf[x, y] = float.MinValue;

            Color fillColor = Color.LightGray;
            Color lineColor = Color.Gray;

            using (var g = Graphics.FromImage(bmp))
            {
                g.Clear(Color.White);
                g.SmoothingMode = System.Drawing.Drawing2D.SmoothingMode.AntiAlias;

                stream.Seek(0, SeekOrigin.Begin);
                if (IsBinaryStl(stream, out uint triCount))
                {
                    // --- Binary STL render pass ---
                    using (var br = new BinaryReader(stream, System.Text.Encoding.ASCII, leaveOpen: true))
                    {
                        stream.Seek(84, SeekOrigin.Begin); // skip header & count
                        int maxTrianglesToRender = maxTriangles;
                        for (uint i = 0; i < triCount && i < maxTrianglesToRender; i++)
                        {
                            // Read one triangle
                            br.ReadSingle(); br.ReadSingle(); br.ReadSingle(); // normal, not used
                            var v1 = new StlModel.Vec3(br.ReadSingle(), br.ReadSingle(), br.ReadSingle());
                            var v2 = new StlModel.Vec3(br.ReadSingle(), br.ReadSingle(), br.ReadSingle());
                            var v3 = new StlModel.Vec3(br.ReadSingle(), br.ReadSingle(), br.ReadSingle());
                            br.ReadUInt16(); // attr

                            RenderTriangleGdiZbuf(projectWithViewZ, v1, v2, v3, g, zbuf, width, height, fillColor, lineColor);
                        }
                    }
                }
                else
                {
                    // --- ASCII STL render pass ---
                    stream.Seek(0, SeekOrigin.Begin);
                    using (var reader = new StreamReader(stream))
                    {
                        string line;
                        StlModel.Vec3 n = new StlModel.Vec3();
                        int maxTrianglesToRender = maxTriangles;
                        int triangleCount = 0;
                        while ((line = reader.ReadLine()) != null && triangleCount < maxTrianglesToRender)
                        {
                            line = line.Trim();
                            if (line.StartsWith("facet normal"))
                            {
                                // process triangle
                                triangleCount++;
                                var parts = line.Split(new[] { ' ' }, StringSplitOptions.RemoveEmptyEntries);
                                n = new StlModel.Vec3(
                                    float.Parse(parts[2], CultureInfo.InvariantCulture),
                                    float.Parse(parts[3], CultureInfo.InvariantCulture),
                                    float.Parse(parts[4], CultureInfo.InvariantCulture)
                                );
                            }
                            else if (line.StartsWith("vertex"))
                            {
                                var v1 = ParseVertex(line);
                                var v2 = ParseVertex(reader.ReadLine());
                                var v3 = ParseVertex(reader.ReadLine());
                                RenderTriangleGdiZbuf(projectWithViewZ, v1, v2, v3, g, zbuf, width, height, fillColor, lineColor);
                            }
                        }
                    }
                }
            }

            return bmp;

            // --- Helpers ---
            StlModel.Vec3 ParseVertex(string vertexLine)
            {
                var parts = vertexLine.Trim().Split(new[] { ' ' }, StringSplitOptions.RemoveEmptyEntries);
                return new StlModel.Vec3(
                    float.Parse(parts[1], CultureInfo.InvariantCulture),
                    float.Parse(parts[2], CultureInfo.InvariantCulture),
                    float.Parse(parts[3], CultureInfo.InvariantCulture)
                );
            }
        }

        /// <summary>
        /// Helper: Draw triangle if any new pixel would be visible, and update z-buffer.
        /// </summary>
        private static void RenderTriangleGdiZbuf(
            Func<StlModel.Vec3, ProjectedPoint> projectWithViewZ,
            StlModel.Vec3 v1, StlModel.Vec3 v2, StlModel.Vec3 v3,
            Graphics g, float[,] zbuf, int width, int height,
            Color fillColor, Color lineColor)
        {
            var p1 = projectWithViewZ(v1);
            var p2 = projectWithViewZ(v2);
            var p3 = projectWithViewZ(v3);

            float minX = Math.Min(p1.Point.X, Math.Min(p2.Point.X, p3.Point.X));
            float maxX = Math.Max(p1.Point.X, Math.Max(p2.Point.X, p3.Point.X));
            float minY = Math.Min(p1.Point.Y, Math.Min(p2.Point.Y, p3.Point.Y));
            float maxY = Math.Max(p1.Point.Y, Math.Max(p2.Point.Y, p3.Point.Y));

            // Fast skip: If triangle is completely outside image, skip processing
            if (maxX < 0 || minX >= width || maxY < 0 || minY >= height)
                return;

            int ix0 = Math.Max(0, (int)Math.Floor(minX));
            int ix1 = Math.Min(width - 1, (int)Math.Ceiling(maxX));
            int iy0 = Math.Max(0, (int)Math.Floor(minY));
            int iy1 = Math.Min(height - 1, (int)Math.Ceiling(maxY));

            float x1 = p1.Point.X, y1 = p1.Point.Y, z1 = p1.Z;
            float x2 = p2.Point.X, y2 = p2.Point.Y, z2 = p2.Z;
            float x3 = p3.Point.X, y3 = p3.Point.Y, z3 = p3.Z;
            float denom = (y2 - y3) * (x1 - x3) + (x3 - x2) * (y1 - y3);
            if (Math.Abs(denom) < 1e-6) return;

            bool visible = false;
            int maskW = ix1 - ix0 + 1;
            int maskH = iy1 - iy0 + 1;
            if (maskW <= 0 || maskH <= 0)
                return; // No pixels to process, skip this triangle.

            bool[,] mask = new bool[maskW, maskH];

            for (int y = iy0; y <= iy1; y++)
            {
                for (int x = ix0; x <= ix1; x++)
                {
                    if (x < 0 || x >= width || y < 0 || y >= height) continue;
                    float px = x + 0.5f;
                    float py = y + 0.5f;
                    float a = ((y2 - y3) * (px - x3) + (x3 - x2) * (py - y3)) / denom;
                    float b = ((y3 - y1) * (px - x3) + (x1 - x3) * (py - y3)) / denom;
                    float c = 1 - a - b;
                    if (a < 0 || b < 0 || c < 0) continue;
                    float z = a * z1 + b * z2 + c * z3;
                    if (z > zbuf[x, y])
                    {
                        mask[x - ix0, y - iy0] = true;
                        visible = true;
                    }
                }
            }

            if (visible)
            {
                g.FillPolygon(new SolidBrush(fillColor), new[] { p1.Point, p2.Point, p3.Point });
                g.DrawPolygon(new Pen(lineColor, 1f), new[] { p1.Point, p2.Point, p3.Point });

                for (int y = iy0; y <= iy1; y++)
                {
                    for (int x = ix0; x <= ix1; x++)
                    {
                        if (!mask[x - ix0, y - iy0]) continue;
                        float px = x + 0.5f;
                        float py = y + 0.5f;
                        float a = ((y2 - y3) * (px - x3) + (x3 - x2) * (py - y3)) / denom;
                        float b = ((y3 - y1) * (px - x3) + (x1 - x3) * (py - y3)) / denom;
                        float c = 1 - a - b;
                        float z = a * z1 + b * z2 + c * z3;
                        zbuf[x, y] = z;
                    }
                }
            }
        }

        /// <summary>
        /// Scan STL file for min/max bounds. Returns struct of min/max for X, Y, Z.
        /// </summary>
        private struct StlBounds
        {
            public float MinX, MinY, MinZ, MaxX, MaxY, MaxZ;
        }

        private static StlBounds ScanStlBounds(Stream stream)
        {
            stream.Seek(0, SeekOrigin.Begin);
            StlBounds bounds = new StlBounds
            {
                MinX = float.MaxValue,
                MinY = float.MaxValue,
                MinZ = float.MaxValue,
                MaxX = float.MinValue,
                MaxY = float.MinValue,
                MaxZ = float.MinValue
            };

            if (IsBinaryStl(stream, out uint triCount))
            {
                // Binary STL
                using (var br = new BinaryReader(stream, System.Text.Encoding.ASCII, leaveOpen: true))
                {
                    stream.Seek(84, SeekOrigin.Begin);
                    int maxTrianglesForBounds = maxTriangles;
                    for (uint i = 0; i < triCount && i < maxTrianglesForBounds; i++)
                    {
                        br.ReadSingle(); br.ReadSingle(); br.ReadSingle(); // normal
                        var vs = new StlModel.Vec3[3];
                        for (int v = 0; v < 3; v++)
                        {
                            vs[v] = new StlModel.Vec3(br.ReadSingle(), br.ReadSingle(), br.ReadSingle());
                            bounds.MinX = Math.Min(bounds.MinX, vs[v].X); bounds.MaxX = Math.Max(bounds.MaxX, vs[v].X);
                            bounds.MinY = Math.Min(bounds.MinY, vs[v].Y); bounds.MaxY = Math.Max(bounds.MaxY, vs[v].Y);
                            bounds.MinZ = Math.Min(bounds.MinZ, vs[v].Z); bounds.MaxZ = Math.Max(bounds.MaxZ, vs[v].Z);
                        }
                        br.ReadUInt16();
                    }
                }
            }
            else
            {
                // ASCII STL
                stream.Seek(0, SeekOrigin.Begin);
                using (var reader = new StreamReader(stream, Encoding.ASCII, detectEncodingFromByteOrderMarks: true, bufferSize: 4096, leaveOpen: true))
                {
                    string line;
                    int maxVerticesForBounds = maxTriangles * 3, vertexCount = 0;
                    while ((line = reader.ReadLine()) != null && vertexCount < maxVerticesForBounds)
                    {
                        line = line.Trim();
                        if (line.StartsWith("vertex"))
                        {
                            // update bounds
                            vertexCount++;
                            var parts = line.Split(new[] { ' ' }, StringSplitOptions.RemoveEmptyEntries);
                            float x = float.Parse(parts[1], CultureInfo.InvariantCulture);
                            float y = float.Parse(parts[2], CultureInfo.InvariantCulture);
                            float z = float.Parse(parts[3], CultureInfo.InvariantCulture);
                            bounds.MinX = Math.Min(bounds.MinX, x); bounds.MaxX = Math.Max(bounds.MaxX, x);
                            bounds.MinY = Math.Min(bounds.MinY, y); bounds.MaxY = Math.Max(bounds.MaxY, y);
                            bounds.MinZ = Math.Min(bounds.MinZ, z); bounds.MaxZ = Math.Max(bounds.MaxZ, z);
                        }
                    }
                }
            }
            return bounds;
        }

        /// <summary>
        /// Determines if an STL stream is binary. Returns triangle count if so.
        /// </summary>
        private static bool IsBinaryStl(Stream stream, out uint triCount)
        {
            stream.Seek(0, SeekOrigin.Begin);
            using (var br = new BinaryReader(stream, System.Text.Encoding.ASCII, leaveOpen: true))
            {
                if (stream.Length < 84)
                {
                    triCount = 0;
                    return false;
                }
                stream.Seek(80, SeekOrigin.Begin);
                triCount = br.ReadUInt32();
                var expectedLength = 84 + triCount * 50;
                return stream.Length == expectedLength;
            }
        }
    }
}