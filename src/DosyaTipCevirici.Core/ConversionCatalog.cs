namespace DosyaTipCevirici.Core;

public enum FileFamily { Document, Spreadsheet, Presentation, Image, Pdf }

public sealed record FileFormat(string Extension, FileFamily Family, bool CanBeTarget = true)
{
    public string Label => Extension.ToUpperInvariant();
}

public static class ConversionCatalog
{
    private static readonly FileFormat[] Formats =
    [
        new("doc", FileFamily.Document),
        new("docx", FileFamily.Document),
        new("docm", FileFamily.Document),
        new("odt", FileFamily.Document),
        new("rtf", FileFamily.Document),
        new("txt", FileFamily.Document),
        new("html", FileFamily.Document),
        new("htm", FileFamily.Document),
        new("xls", FileFamily.Spreadsheet),
        new("xlsx", FileFamily.Spreadsheet),
        new("xlsm", FileFamily.Spreadsheet),
        new("xlsb", FileFamily.Spreadsheet),
        new("ods", FileFamily.Spreadsheet),
        new("csv", FileFamily.Spreadsheet),
        new("ppt", FileFamily.Presentation),
        new("pptx", FileFamily.Presentation),
        new("pptm", FileFamily.Presentation),
        new("pps", FileFamily.Presentation),
        new("ppsx", FileFamily.Presentation),
        new("odp", FileFamily.Presentation),
        new("png", FileFamily.Image),
        new("jpg", FileFamily.Image),
        new("jpeg", FileFamily.Image),
        new("bmp", FileFamily.Image),
        new("gif", FileFamily.Image),
        new("tif", FileFamily.Image),
        new("tiff", FileFamily.Image),
        new("heic", FileFamily.Image, false),
        new("heif", FileFamily.Image, false),
        new("pdf", FileFamily.Pdf)
    ];

    public static IReadOnlyList<FileFormat> All => Formats;

    public static FileFormat? Find(string pathOrExtension)
    {
        var extension = Path.GetExtension(pathOrExtension).TrimStart('.').ToLowerInvariant();
        return Formats.FirstOrDefault(format => format.Extension == extension);
    }

    public static IReadOnlyList<FileFormat> TargetsFor(FileFormat source)
    {
        if (source.Family == FileFamily.Pdf)
            return Formats.Where(format => format.CanBeTarget && format.Family == FileFamily.Document).ToArray();

        if (source.Family == FileFamily.Image)
            return Formats.Where(format => format.CanBeTarget && format.Family == FileFamily.Image && format.Extension != source.Extension).ToArray();

        return Formats.Where(format =>
            format.CanBeTarget &&
            format.Extension != source.Extension &&
            (format.Family == source.Family || format.Family == FileFamily.Pdf)).ToArray();
    }

    public static bool CanConvert(FileFormat source, FileFormat target) =>
        TargetsFor(source).Any(format => format.Extension == target.Extension);
}
