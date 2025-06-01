using SkiaSharp;
using Svg.Skia;

namespace ClosedXML.Report.XLCustom.Functions;

/// <summary>
/// Helper class providing image processing functionality
/// </summary>
public static class ImageHelper
{
    private static readonly HttpClient _httpClient = new HttpClient { Timeout = TimeSpan.FromSeconds(15) };
    private const int DEFAULT_SVG_WIDTH = 300;
    private const int DEFAULT_SVG_HEIGHT = 300;
    private const int MAX_IMAGE_DIMENSION = 2000;

    /// <summary>
    /// Gets image from file path or downloads it from URL
    /// </summary>
    public static string? GetImageFromPathOrUrl(string pathOrUrl)
    {
        if (string.IsNullOrEmpty(pathOrUrl))
            return null;

        if (Uri.TryCreate(pathOrUrl, UriKind.Absolute, out Uri? uri) &&
            (uri.Scheme == Uri.UriSchemeHttp || uri.Scheme == Uri.UriSchemeHttps))
        {
            return DownloadImage(uri);
        }

        return File.Exists(pathOrUrl) ? pathOrUrl : null;
    }

    /// <summary>
    /// Downloads an image from a URL
    /// </summary>
    private static string? DownloadImage(Uri uri)
    {
        try
        {
            // Create temp file with appropriate extension
            string extension = Path.GetExtension(uri.AbsolutePath);
            if (string.IsNullOrEmpty(extension))
                extension = ".tmp";

            string tempFile = Path.Combine(
                Path.GetTempPath(),
                $"xlimg_{Guid.NewGuid()}{extension}");

            // Download image
            byte[] imageBytes;
            string? contentType = null;

            using (var response = _httpClient.GetAsync(uri).GetAwaiter().GetResult())
            {
                if (!response.IsSuccessStatusCode)
                {
                    Log.Debug($"HTTP error downloading image: {response.StatusCode}");
                    return null;
                }

                contentType = response.Content.Headers.ContentType?.MediaType;
                imageBytes = response.Content.ReadAsByteArrayAsync().GetAwaiter().GetResult();
            }

            // Handle SVG images
            if (contentType == "image/svg+xml" || extension.ToLowerInvariant() == ".svg")
            {
                string svgFile = tempFile;
                string pngFile = Path.ChangeExtension(tempFile, ".png");

                // Save SVG file
                File.WriteAllBytes(svgFile, imageBytes);

                // Convert to PNG
                if (ConvertSvgToPng(svgFile, pngFile))
                {
                    // Clean up temporary SVG file
                    try { File.Delete(svgFile); } catch { }
                    return pngFile;
                }

                // If conversion fails, clean up and return null
                try { File.Delete(svgFile); } catch { }
                return null;
            }

            // Save regular image
            File.WriteAllBytes(tempFile, imageBytes);

            // Verify file was created
            if (!File.Exists(tempFile) || new FileInfo(tempFile).Length == 0)
            {
                Log.Debug("Downloaded file is empty or not created");
                return null;
            }

            return tempFile;
        }
        catch (Exception ex)
        {
            Log.Debug($"Image download error: {ex.Message}");
            return null;
        }
    }

    /// <summary>
    /// Converts SVG to PNG using SkiaSharp
    /// </summary>
    public static bool ConvertSvgToPng(string svgFilePath, string pngFilePath)
    {
        try
        {
            using (var svg = new SKSvg())
            {
                // Load SVG
                svg.Load(svgFilePath);

                if (svg.Picture == null)
                {
                    Log.Debug("Failed to load SVG - Picture is null");
                    return false;
                }

                // Get and validate dimensions
                SKRect bounds = svg.Picture.CullRect;
                int width = (int)bounds.Width;
                int height = (int)bounds.Height;

                // Use default dimensions if invalid
                if (width <= 0 || height <= 0)
                {
                    width = DEFAULT_SVG_WIDTH;
                    height = DEFAULT_SVG_HEIGHT;
                }

                // Limit dimensions for safety
                width = Math.Min(width, MAX_IMAGE_DIMENSION);
                height = Math.Min(height, MAX_IMAGE_DIMENSION);

                // Create PNG from SVG
                using (var surface = SKSurface.Create(new SKImageInfo(width, height)))
                {
                    if (surface == null)
                    {
                        Log.Debug("Failed to create surface for SVG rendering");
                        return false;
                    }

                    var canvas = surface.Canvas;
                    canvas.Clear(SKColors.Transparent);
                    canvas.DrawPicture(svg.Picture);
                    canvas.Flush();

                    // Save to file
                    using (var image = surface.Snapshot())
                    using (var data = image.Encode(SKEncodedImageFormat.Png, 100))
                    using (var stream = File.OpenWrite(pngFilePath))
                    {
                        data.SaveTo(stream);
                    }

                    return File.Exists(pngFilePath) && new FileInfo(pngFilePath).Length > 0;
                }
            }
        }
        catch (Exception ex)
        {
            Log.Debug($"SVG conversion error: {ex.Message}");
            return false;
        }
    }
}