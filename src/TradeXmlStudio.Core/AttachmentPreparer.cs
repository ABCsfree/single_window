using System.IO;
using System.Windows;
using System.Windows.Media;
using System.Windows.Media.Imaging;

namespace TradeXmlStudio.Core;

internal static class AttachmentPreparer
{
    // Local safety budget, not a confirmed platform limit. Leave 64 KiB for
    // XML metadata and the import client's signature below the suspected 3 MiB cap.
    internal const int XmlBudget = 3 * 1024 * 1024 - 64 * 1024;
    private const int PayloadBudget = (XmlBudget - 16 * 1024) / 4 * 3;

    internal static EdocSource Prepare(EdocSource source, TradeXmlOptions options)
    {
        if (Thread.CurrentThread.GetApartmentState() == ApartmentState.STA)
            return PrepareCore(source, options);
        EdocSource? result = null;
        Exception? failure = null;
        var worker = new Thread(() =>
        {
            try { result = PrepareCore(source, options); }
            catch (Exception ex) { failure = ex; }
        });
        worker.SetApartmentState(ApartmentState.STA);
        worker.Start();
        worker.Join();
        if (failure is not null)
            System.Runtime.ExceptionServices.ExceptionDispatchInfo.Capture(failure).Throw();
        return result!;
    }

    private static EdocSource PrepareCore(EdocSource source, TradeXmlOptions options)
    {
        var bytes = File.ReadAllBytes(source.FullPath);
        var limit = Math.Min(options.MaxImageBytes, PayloadBudget);
        var isPhoto = TradeXmlOptions.PhotoExtensions.Contains(
            Path.GetExtension(source.FullPath), StringComparer.OrdinalIgnoreCase);
        if (!isPhoto)
        {
            if (bytes.Length == 0 || bytes.Length > limit)
                throw Error(source, "附件为空或超过大小目标；非图片附件请手动压缩。");
            return source with { PreparedContent = bytes };
        }

        try
        {
            using var input = new MemoryStream(bytes);
            var decoder = BitmapDecoder.Create(input, BitmapCreateOptions.PreservePixelFormat, BitmapCacheOption.OnLoad);
            var type = decoder switch
            {
                JpegBitmapDecoder => "jpg",
                PngBitmapDecoder => "png",
                BmpBitmapDecoder => "bmp",
                _ => throw Error(source, "仅支持真实格式为 JPEG、PNG 或 BMP 的照片。")
            };
            if (decoder.Frames.Count != 1)
                throw Error(source, "不支持多帧图片，请先手动转换。");
            var frame = decoder.Frames[0];
            // Force pixel decoding even for files that will pass through unchanged.
            var stride = checked((frame.PixelWidth * frame.Format.BitsPerPixel + 7) / 8);
            frame.CopyPixels(new byte[checked(stride * frame.PixelHeight)], stride, 0);
            if (bytes.Length <= limit)
            {
                var extension = Path.GetExtension(source.FileName).ToLowerInvariant();
                var matches = extension == "." + type || (type == "jpg" && extension == ".jpeg");
                return source with
                {
                    FileName = matches ? source.FileName : Path.ChangeExtension(source.FileName, type),
                    AttTypeCode = matches && extension == ".jpeg" ? "jpeg" : type,
                    PreparedContent = bytes
                };
            }

            // Render over white to flatten transparency. Apply EXIF orientation
            // before encoding because the output intentionally drops metadata.
            BitmapSource oriented = frame;
            var orientation = GetOrientation(frame);
            var transform = orientation switch
            {
                2 => new Matrix(-1, 0, 0, 1, 0, 0),
                3 => new Matrix(-1, 0, 0, -1, 0, 0),
                4 => new Matrix(1, 0, 0, -1, 0, 0),
                5 => new Matrix(0, 1, 1, 0, 0, 0),
                6 => new Matrix(0, 1, -1, 0, 0, 0),
                7 => new Matrix(0, -1, -1, 0, 0, 0),
                8 => new Matrix(0, -1, 1, 0, 0, 0),
                _ => Matrix.Identity
            };
            if (!transform.IsIdentity)
                oriented = new TransformedBitmap(frame, new MatrixTransform(transform));
            var visual = new DrawingVisual();
            using (var drawing = visual.RenderOpen())
            {
                var bounds = new Rect(0, 0, oriented.PixelWidth, oriented.PixelHeight);
                drawing.DrawRectangle(Brushes.White, null, bounds);
                drawing.DrawImage(oriented, bounds);
            }
            var flattened = new RenderTargetBitmap(oriented.PixelWidth, oriented.PixelHeight, 96, 96, PixelFormats.Pbgra32);
            flattened.Render(visual);
            foreach (var quality in new[] { 95, 90, 85, 80 })
            {
                var encoder = new JpegBitmapEncoder { QualityLevel = quality };
                encoder.Frames.Add(BitmapFrame.Create(flattened));
                using var output = new MemoryStream();
                encoder.Save(output);
                if (output.Length <= limit)
                    return source with
                    {
                        FileName = Path.ChangeExtension(source.FileName, "jpg"),
                        AttTypeCode = "jpg",
                        PreparedContent = output.ToArray()
                    };
            }
            throw Error(source, "JPEG 质量降至 80 后仍超过大小目标，请手动处理；未缩小图片分辨率。");
        }
        catch (Exception ex) when (ex is not XmlGenerationException)
        {
            throw Error(source, $"图片读取或转换失败：{ex.Message}");
        }
    }

    private static int GetOrientation(BitmapFrame frame)
    {
        if (frame.Metadata is not BitmapMetadata metadata) return 1;
        foreach (var query in new[] { "/app1/ifd/{ushort=274}", "/ifd/{ushort=274}" })
            if (metadata.ContainsQuery(query) && metadata.GetQuery(query) is ushort value)
                return value;
        return 1;
    }

    private static XmlGenerationException Error(EdocSource source, string message) =>
        new([$"{source.FileName}：{message}"]);
}
