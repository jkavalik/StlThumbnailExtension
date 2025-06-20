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
using System.Linq;

namespace StlThumbnailExtension
{
    [ComVisible(true)]
    [COMServerAssociation(AssociationType.FileExtension, ".stl")]
    public class StlThumbnailHandler : SharpThumbnailHandler
    {
        protected override Bitmap GetThumbnailImage(uint width)
        {
            //  Create a stream reader for the selected item stream
            try
            {
                using (var stream = new BufferedStream(SelectedItemStream))
                {
                    var model = StlModel.Load(stream);
                    return StlRenderer.RenderToBitmap(model, (int)width, (int)width);
                }
            }
            catch (Exception exception)
            {
                //  Log the exception and return null for failure (no thumbnail to show)
                LogError("An exception occurred generating the thumbnail.", exception);
                return null;
            }
        }
    }

    public class StlModel
    {
        public List<Triangle> Triangles = new List<Triangle>();

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

        public static StlModel Load(Stream stream)
        {
            using (var br = new BinaryReader(stream, System.Text.Encoding.ASCII, leaveOpen: true))
            {
                stream.Seek(0, SeekOrigin.Begin);
                var header = br.ReadBytes(80);
                var count = br.ReadUInt32();
                var expectedLength = 84 + count * 50;
                if (stream.Length == expectedLength)
                {
                    var model = new StlModel();
                    for (uint i = 0; i < count; i++)
                    {
                        var normal = new Vec3(br.ReadSingle(), br.ReadSingle(), br.ReadSingle());
                        var v1 = new Vec3(br.ReadSingle(), br.ReadSingle(), br.ReadSingle());
                        var v2 = new Vec3(br.ReadSingle(), br.ReadSingle(), br.ReadSingle());
                        var v3 = new Vec3(br.ReadSingle(), br.ReadSingle(), br.ReadSingle());
                        br.ReadUInt16();
                        model.Triangles.Add(new Triangle { Normal = normal, V1 = v1, V2 = v2, V3 = v3 });
                    }
                    return model;
                }
            }

            stream.Seek(0, SeekOrigin.Begin);
            using (var reader = new StreamReader(stream))
            {
                var model = new StlModel();
                string line;
                StlModel.Vec3 n = new StlModel.Vec3();
                while ((line = reader.ReadLine()) != null)
                {
                    line = line.Trim();
                    if (line.StartsWith("facet normal"))
                    {
                        var parts = line.Split(new[] { ' ' }, StringSplitOptions.RemoveEmptyEntries);
                        n = new StlModel.Vec3(
                            float.Parse(parts[2], CultureInfo.InvariantCulture),
                            float.Parse(parts[3], CultureInfo.InvariantCulture),
                            float.Parse(parts[4], CultureInfo.InvariantCulture)
                        );
                    }
                    else if (line.StartsWith("vertex"))
                    {
                        var v = ParseVertex(line);
                        var v2 = ParseVertex(reader.ReadLine());
                        var v3 = ParseVertex(reader.ReadLine());
                        model.Triangles.Add(new Triangle { Normal = n, V1 = v, V2 = v2, V3 = v3 });
                    }
                }
                return model;

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
        }
    }

    public static class StlRenderer
    {
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
        // Helper struct to hold projected triangles and their average Z for sorting
        private struct ProjectedTriangle
        {
            public PointF P1, P2, P3;
            public float AvgZ;
        }

        public static Bitmap RenderToBitmap(StlModel model, int width, int height)
        {
            float minX = float.MaxValue, minY = float.MaxValue, minZ = float.MaxValue;
            float maxX = float.MinValue, maxY = float.MinValue, maxZ = float.MinValue;
            foreach (var tri in model.Triangles)
            {
                foreach (var v in new[] { tri.V1, tri.V2, tri.V3 })
                {
                    if (v.X < minX) minX = v.X; if (v.X > maxX) maxX = v.X;
                    if (v.Y < minY) minY = v.Y; if (v.Y > maxY) maxY = v.Y;
                    if (v.Z < minZ) minZ = v.Z; if (v.Z > maxZ) maxZ = v.Z;
                }
            }
            var center = new StlModel.Vec3((minX + maxX) / 2, (minY + maxY) / 2, (minZ + maxZ) / 2);
            float scale = Math.Min(width, height) * 0.8f / Math.Max(maxX - minX, Math.Max(maxY - minY, maxZ - minZ));

            float angle = (float)(Math.PI / 6);

            Func<StlModel.Vec3, ProjectedPoint> projectWithViewZ = v =>
            {
                float x = v.X - center.X, y = v.Y - center.Y, z = v.Z - center.Z;
                float x1 = x * (float)Math.Cos(angle) - z * (float)Math.Sin(angle);
                float z1 = x * (float)Math.Sin(angle) + z * (float)Math.Cos(angle);
                float y1 = y * (float)Math.Cos(angle / 2) - z1 * (float)Math.Sin(angle / 2);
                float z2 = y * (float)Math.Sin(angle / 2) + z1 * (float)Math.Cos(angle / 2);
                return new ProjectedPoint(
                    new PointF(
                        width / 2 + x1 * scale,
                        height / 2 - y1 * scale
                    ),
                    z2
                );
            };

            var projectedTriangles = new List<ProjectedTriangle>(model.Triangles.Count);

            foreach (var tri in model.Triangles)
            {
                var pr1 = projectWithViewZ(tri.V1);
                var pr2 = projectWithViewZ(tri.V2);
                var pr3 = projectWithViewZ(tri.V3);
                float avgZ = (pr1.Z + pr2.Z + pr3.Z) / 3f;
                projectedTriangles.Add(new ProjectedTriangle
                {
                    P1 = pr1.Point,
                    P2 = pr2.Point,
                    P3 = pr3.Point,
                    AvgZ = avgZ
                });
            }

            // Sort triangles back-to-front (lowest Z first)
            projectedTriangles.Sort((a, b) => a.AvgZ.CompareTo(b.AvgZ));

            var bmp = new Bitmap(width, height, PixelFormat.Format32bppArgb);
            using (var g = Graphics.FromImage(bmp))
            {
                g.Clear(Color.White);
                g.SmoothingMode = System.Drawing.Drawing2D.SmoothingMode.AntiAlias;
                var brush = new SolidBrush(Color.LightGray);
                var pen = new Pen(Color.Gray, 1f);

                foreach (var tri in projectedTriangles)
                {
                    g.FillPolygon(brush, new[] { tri.P1, tri.P2, tri.P3 });
                    g.DrawPolygon(pen, new[] { tri.P1, tri.P2, tri.P3 });
                }
            }
            return bmp;
        }
    }
}