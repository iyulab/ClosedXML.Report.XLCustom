namespace ClosedXML.Report.XLCustom.Functions;

/// <summary>
/// Provides image-related function handlers with direct tag-based processing
/// </summary>
public static class ImageFunction
{
    // Regular expression for image markers
    private static readonly Regex ImageMarkerRegex = new Regex(@"\[IMG:([^\]]+)\]", RegexOptions.Compiled);

    // Constants for pixel conversion
    private const double COLUMN_WIDTH_TO_PIXEL_RATIO = 7.0; // Approximate: 1 column width unit ≈ 7 pixels
    private const double ROW_HEIGHT_TO_PIXEL_RATIO = 1.33; // Approximate: 1 point ≈ 1.33 pixels

    /// <summary>
    /// Registers all image-related functions to the specified registry
    /// </summary>
    public static void RegisterImageFunctions(FunctionRegistry registry)
    {
        if (registry == null)
            throw new ArgumentNullException(nameof(registry));

        registry.Register("image", HandleImageFunction);
        Log.Debug("Registered image function");
    }

    /// <summary>
    /// Handles the 'image' function by placing an image marker in the cell
    /// </summary>
    private static void HandleImageFunction(IXLCell cell, object value, string[] parameters)
    {
        if (cell == null || value == null)
        {
            Log.Debug("Image function called with null cell or value");
            return;
        }

        var pathOrUrl = value.ToString();
        if (string.IsNullOrEmpty(pathOrUrl))
        {
            Log.Debug("Image function called with empty path/URL");
            return;
        }

        try
        {
            // Get image from path or URL using ImageHelper
            var localPath = ImageHelper.GetImageFromPathOrUrl(pathOrUrl);

            if (string.IsNullOrEmpty(localPath))
            {
                Log.Debug($"Failed to get image from path or URL: {pathOrUrl}");
                cell.Value = $"Error: Unable to load image";
                cell.Style.Font.FontColor = XLColor.Red;
                return;
            }

            // Create marker and set it in the cell
            string marker = $"[IMG:{localPath}]";
            cell.Value = marker;
            Log.Debug($"Set image marker in cell {cell.Address} with path {localPath}");
        }
        catch (Exception ex)
        {
            Log.Debug($"Error in image function: {ex.Message}");
            cell.Value = $"Error: {ex.Message}";
            cell.Style.Font.FontColor = XLColor.Red;
        }
    }

    /// <summary>
    /// Process all image markers in the workbook by scanning all worksheets and cells
    /// </summary>
    public static void ProcessAllImageMarkers(IXLWorkbook workbook)
    {
        if (workbook == null)
        {
            Log.Debug("Cannot process images: workbook is null");
            return;
        }

        Log.Debug("Processing image markers in all worksheets");
        int processedCount = 0;

        // Iterate through all worksheets
        foreach (var worksheet in workbook.Worksheets)
        {
            try
            {
                Log.Debug($"Scanning worksheet '{worksheet.Name}' for image markers");
                processedCount += ProcessImagesInWorksheet(worksheet);
            }
            catch (ObjectDisposedException)
            {
                Log.Debug($"Worksheet already disposed, skipping");
            }
            catch (Exception ex)
            {
                Log.Debug($"Error scanning worksheet: {ex.Message}");
            }
        }

        Log.Debug($"Processed {processedCount} images total across all worksheets");
    }

    /// <summary>
    /// Process all image markers in a single worksheet
    /// </summary>
    public static int ProcessImagesInWorksheet(IXLWorksheet worksheet)
    {
        if (worksheet == null)
        {
            Log.Debug("Cannot process images: worksheet is null");
            return 0;
        }

        int processedCount = 0;

        try
        {
            // Get the used range (if any)
            var usedRange = worksheet.RangeUsed();
            if (usedRange == null)
            {
                Log.Debug($"Worksheet '{worksheet.Name}' has no used range");
                return 0;
            }

            // Process each cell in the used range
            foreach (var cell in usedRange.CellsUsed())
            {
                try
                {
                    if (!cell.Value.IsText)
                        continue;

                    string cellValue = cell.Value.ToString();
                    var match = ImageMarkerRegex.Match(cellValue);

                    if (!match.Success)
                        continue;

                    string imagePath = match.Groups[1].Value;

                    if (string.IsNullOrEmpty(imagePath))
                    {
                        Log.Debug($"Empty image path in cell {cell.Address}");
                        continue;
                    }

                    if (!File.Exists(imagePath))
                    {
                        Log.Debug($"Image file not found: {imagePath}");
                        cell.Value = "Image Not Found";
                        cell.Style.Font.FontColor = XLColor.Red;
                        continue;
                    }

                    // Process the image
                    if (ProcessImageInCell(worksheet, cell, imagePath))
                    {
                        processedCount++;
                    }
                }
                catch (Exception ex)
                {
                    Log.Debug($"Error processing cell: {ex.Message}");
                }
            }

            Log.Debug($"Processed {processedCount} images in worksheet '{worksheet.Name}'");
            return processedCount;
        }
        catch (ObjectDisposedException)
        {
            Log.Debug("Worksheet is disposed, cannot process images");
            return 0;
        }
        catch (Exception ex)
        {
            Log.Debug($"Error processing images in worksheet: {ex.Message}");
            return 0;
        }
    }

    /// <summary>
    /// Process a single image in a cell, maintaining aspect ratio and centering it
    /// </summary>
    private static bool ProcessImageInCell(IXLWorksheet worksheet, IXLCell cell, string imagePath)
    {
        try
        {
            // Clear the marker text
            cell.Value = "";

            try
            {
                // Get cell dimensions and position
                var cellDimensions = GetCellDimensions(cell);
                int cellWidth = cellDimensions.Width;
                int cellHeight = cellDimensions.Height;

                // Get image dimensions
                var imageDimensions = GetImageDimensions(imagePath);
                int imageWidth = imageDimensions.Width;
                int imageHeight = imageDimensions.Height;

                if (imageWidth <= 0 || imageHeight <= 0)
                {
                    Log.Debug($"Invalid image dimensions: {imageWidth}x{imageHeight}");
                    cell.Value = "Invalid Image";
                    cell.Style.Font.FontColor = XLColor.Red;
                    return false;
                }

                // Calculate size to fit in cell while preserving aspect ratio
                var finalSize = CalculateSizeToFit(imageWidth, imageHeight, cellWidth, cellHeight);
                int finalWidth = finalSize.Width;
                int finalHeight = finalSize.Height;

                // Calculate center position offsets
                int offsetX = (cellWidth - finalWidth) / 2;
                int offsetY = (cellHeight - finalHeight) / 2;

                // Add and position the picture
                var picture = worksheet.AddPicture(imagePath);
                picture.MoveTo(cell); // First move to the cell

                // Set the size while maintaining aspect ratio
                picture.WithSize(finalWidth, finalHeight);

                // Apply offset for centering within the cell
                picture.MoveTo(picture.TopLeftCell, offsetX, offsetY);

                Log.Debug($"Added image at {cell.Address} from {imagePath} with size {finalWidth}x{finalHeight} and offsets {offsetX},{offsetY}");
                return true;
            }
            catch (Exception ex)
            {
                Log.Debug($"Error adding picture: {ex.Message}");
                cell.Value = "Image Error";
                cell.Style.Font.FontColor = XLColor.Red;
                return false;
            }
        }
        catch (Exception ex)
        {
            Log.Debug($"Failed to process image in cell: {ex.Message}");
            return false;
        }
    }

    /// <summary>
    /// Get the dimensions of a cell in pixels, considering merged cells
    /// </summary>
    private static (int Width, int Height) GetCellDimensions(IXLCell cell)
    {
        try
        {
            if (cell == null)
                return (0, 0);

            var worksheet = cell.Worksheet;
            if (worksheet == null)
                return (0, 0);

            int cellWidth, cellHeight;

            // Check if cell is merged
            if (cell.IsMerged())
            {
                var mergedRange = cell.MergedRange();

                // Sum width of all columns in merged range
                cellWidth = 0;
                for (int col = mergedRange.FirstColumn().ColumnNumber(); col <= mergedRange.LastColumn().ColumnNumber(); col++)
                {
                    cellWidth += ConvertColumnWidthToPixels(worksheet.Column(col).Width);
                }

                // Sum height of all rows in merged range
                cellHeight = 0;
                for (int row = mergedRange.FirstRow().RowNumber(); row <= mergedRange.LastRow().RowNumber(); row++)
                {
                    cellHeight += ConvertRowHeightToPixels(worksheet.Row(row).Height);
                }
            }
            else
            {
                // Get dimensions of a single cell
                cellWidth = ConvertColumnWidthToPixels(worksheet.Column(cell.Address.ColumnNumber).Width);
                cellHeight = ConvertRowHeightToPixels(worksheet.Row(cell.Address.RowNumber).Height);
            }

            // Ensure minimum dimensions
            return (Math.Max(10, cellWidth), Math.Max(10, cellHeight));
        }
        catch (Exception ex)
        {
            Log.Debug($"Error getting cell dimensions: {ex.Message}");
            return (100, 100); // Default fallback size
        }
    }

    /// <summary>
    /// Get the dimensions of an image in pixels using ImageSharp (cross-platform compatible)
    /// </summary>
    private static (int Width, int Height) GetImageDimensions(string imagePath)
    {
        try
        {
            if (string.IsNullOrEmpty(imagePath) || !File.Exists(imagePath))
                return (0, 0);

            try
            {
                using var image = SixLabors.ImageSharp.Image.Load(imagePath);
                return (image.Width, image.Height);
            }
            catch (Exception ex)
            {
                Log.Debug($"Error getting image dimensions with ImageSharp: {ex.Message}");
                return (100, 100); // Default fallback size
            }
        }
        catch (Exception ex)
        {
            Log.Debug($"Error getting image dimensions: {ex.Message}");
            return (100, 100); // Default fallback size
        }
    }

    /// <summary>
    /// Calculate the size to fit an image in a cell while preserving aspect ratio
    /// </summary>
    private static (int Width, int Height) CalculateSizeToFit(int imageWidth, int imageHeight, int cellWidth, int cellHeight)
    {
        if (imageWidth <= 0 || imageHeight <= 0 || cellWidth <= 0 || cellHeight <= 0)
            return (0, 0);

        double imageRatio = (double)imageWidth / imageHeight;
        double cellRatio = (double)cellWidth / cellHeight;

        int finalWidth, finalHeight;

        if (imageRatio > cellRatio)
        {
            // Image is wider than cell (relative to height)
            // Scale based on width
            finalWidth = cellWidth;
            finalHeight = (int)(finalWidth / imageRatio);

            // Ensure image fits within cell height
            if (finalHeight > cellHeight)
            {
                finalHeight = cellHeight;
                finalWidth = (int)(finalHeight * imageRatio);
            }
        }
        else
        {
            // Image is taller than cell (relative to width)
            // Scale based on height
            finalHeight = cellHeight;
            finalWidth = (int)(finalHeight * imageRatio);

            // Ensure image fits within cell width
            if (finalWidth > cellWidth)
            {
                finalWidth = cellWidth;
                finalHeight = (int)(finalWidth / imageRatio);
            }
        }

        // Apply a small margin for visual appeal (95% of available space)
        finalWidth = (int)(finalWidth * 0.95);
        finalHeight = (int)(finalHeight * 0.95);

        // Ensure minimum dimensions
        finalWidth = Math.Max(1, finalWidth);
        finalHeight = Math.Max(1, finalHeight);

        return (finalWidth, finalHeight);
    }

    /// <summary>
    /// Convert Excel column width to pixels
    /// </summary>
    private static int ConvertColumnWidthToPixels(double columnWidth)
    {
        return (int)(columnWidth * COLUMN_WIDTH_TO_PIXEL_RATIO);
    }

    /// <summary>
    /// Convert Excel row height to pixels
    /// </summary>
    private static int ConvertRowHeightToPixels(double rowHeight)
    {
        return (int)(rowHeight * ROW_HEIGHT_TO_PIXEL_RATIO);
    }
}