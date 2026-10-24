using System.Drawing;
using System.Drawing.Imaging;
using System.IO;
using System.Runtime.InteropServices;
using System.Text;
using System.Windows.Media.Imaging;
using DosyaTipCevirici.Core;

namespace DosyaTipCevirici.Infrastructure;

public sealed class OfficeConversionEngine : IConversionEngine
{
    public Task ConvertAsync(string sourcePath, string targetPath, CancellationToken cancellationToken)
    {
        var completion = new TaskCompletionSource(TaskCreationOptions.RunContinuationsAsynchronously);
        var thread = new Thread(() =>
        {
            try
            {
                cancellationToken.ThrowIfCancellationRequested();
                var source = ConversionCatalog.Find(sourcePath)!;
                if (source.Family == FileFamily.Image)
                    ConvertImage(sourcePath, targetPath);
                else if (source.Family == FileFamily.Pdf || source.Family == FileFamily.Document)
                    ConvertDocument(sourcePath, targetPath, cancellationToken);
                else if (source.Family == FileFamily.Spreadsheet)
                    ConvertExcel(sourcePath, targetPath, cancellationToken);
                else
                    ConvertPowerPoint(sourcePath, targetPath, cancellationToken);
                completion.TrySetResult();
            }
            catch (Exception exception)
            {
                try { if (File.Exists(targetPath)) File.Delete(targetPath); } catch { /* Birincil hatayı koru. */ }
                completion.TrySetException(exception);
            }
        }) { IsBackground = true, Name = "Office conversion" };
        thread.SetApartmentState(ApartmentState.STA);
        thread.Start();
        return completion.Task;
    }

    public static bool IsOfficeAvailable(string programId) => Type.GetTypeFromProgID(programId) is not null;

    private static dynamic StartOffice(string programId)
    {
        var type = Type.GetTypeFromProgID(programId)
            ?? throw new InvalidOperationException($"{programId} bulunamadı. İlgili Microsoft Office uygulamasını kurun.");
        return Activator.CreateInstance(type)
            ?? throw new InvalidOperationException($"{programId} başlatılamadı.");
    }

    private static void ConvertDocument(string sourcePath, string targetPath, CancellationToken token)
    {
        var srcExt = Path.GetExtension(sourcePath).ToLowerInvariant();
        var tgtExt = Path.GetExtension(targetPath).ToLowerInvariant();

        // 1. Doğrudan C# dönüşümleri (Office gerektirmez)
        if (srcExt is ".md" or ".markdown")
        {
            if (tgtExt is ".html" or ".htm")
            {
                var content = File.ReadAllText(sourcePath, Encoding.UTF8);
                var title = Path.GetFileNameWithoutExtension(sourcePath);
                File.WriteAllText(targetPath, MarkdownConverter.ToHtml(content, title), Encoding.UTF8);
                return;
            }
            if (tgtExt is ".txt")
            {
                var content = File.ReadAllText(sourcePath, Encoding.UTF8);
                File.WriteAllText(targetPath, MarkdownConverter.ToPlainText(content), Encoding.UTF8);
                return;
            }
            if (tgtExt is ".md" or ".markdown")
            {
                File.Copy(sourcePath, targetPath, overwrite: true);
                return;
            }
        }

        if (tgtExt is ".md" or ".markdown")
        {
            if (srcExt is ".html" or ".htm")
            {
                var html = File.ReadAllText(sourcePath, Encoding.UTF8);
                File.WriteAllText(targetPath, MarkdownConverter.FromHtml(html), Encoding.UTF8);
                return;
            }
            if (srcExt is ".txt" or ".log")
            {
                File.Copy(sourcePath, targetPath, overwrite: true);
                return;
            }
        }

        if (srcExt is ".log" && tgtExt is ".txt")
        {
            File.Copy(sourcePath, targetPath, overwrite: true);
            return;
        }

        // 2. Markdown -> Word/PDF: Önce zengin HTML'e dönüştür, Word ile açıp istenen biçimde kaydet
        if (srcExt is ".md" or ".markdown")
        {
            var tempHtml = Path.Combine(Path.GetTempPath(), $"conv_{Guid.NewGuid():N}.html");
            try
            {
                var md = File.ReadAllText(sourcePath, Encoding.UTF8);
                var title = Path.GetFileNameWithoutExtension(sourcePath);
                File.WriteAllText(tempHtml, MarkdownConverter.ToHtml(md, title), Encoding.UTF8);
                ConvertWordViaOffice(tempHtml, targetPath, token);
            }
            finally
            {
                try { if (File.Exists(tempHtml)) File.Delete(tempHtml); } catch { /* Temizlik */ }
            }
            return;
        }

        // 3. Word/PDF -> Markdown: Word ile filtrelenmiş HTML'e aktar, ardından Markdown'a dönüştür
        if (tgtExt is ".md" or ".markdown")
        {
            var tempHtml = Path.Combine(Path.GetTempPath(), $"conv_{Guid.NewGuid():N}.html");
            var tempFilesDir = Path.Combine(Path.GetTempPath(), Path.GetFileNameWithoutExtension(tempHtml) + "_files");
            try
            {
                ConvertWordViaOffice(sourcePath, tempHtml, token);
                var html = File.ReadAllText(tempHtml, Encoding.UTF8);
                var md = MarkdownConverter.FromHtml(html);
                File.WriteAllText(targetPath, md, Encoding.UTF8);
            }
            finally
            {
                try { if (File.Exists(tempHtml)) File.Delete(tempHtml); } catch { /* Temizlik */ }
                try { if (Directory.Exists(tempFilesDir)) Directory.Delete(tempFilesDir, recursive: true); } catch { /* Temizlik */ }
            }
            return;
        }

        // 4. Standart Word belgeleri ve PDF dönüşümleri
        ConvertWordViaOffice(sourcePath, targetPath, token);
    }

    private static void ConvertWordViaOffice(string sourcePath, string targetPath, CancellationToken token)
    {
        dynamic? application = null;
        dynamic? document = null;
        try
        {
            application = StartOffice("Word.Application");
            application.Visible = false;
            application.DisplayAlerts = 0;
            application.AutomationSecurity = 3;
            token.ThrowIfCancellationRequested();
            document = application.Documents.Open(sourcePath, ReadOnly: true, AddToRecentFiles: false,
                ConfirmConversions: false, NoEncodingDialog: true);
            token.ThrowIfCancellationRequested();
            var extension = Path.GetExtension(targetPath).ToLowerInvariant();
            if (extension == ".pdf") document.ExportAsFixedFormat(targetPath, 17);
            else if (extension is ".txt" or ".log") document.SaveAs2(targetPath, 2, Encoding: 65001);
            else document.SaveAs2(targetPath, extension switch
            {
                ".doc" => 0, ".docx" => 16, ".docm" => 13,
                ".odt" => 23, ".rtf" => 6, ".html" or ".htm" => 10,
                _ => throw new NotSupportedException(extension)
            });
        }
        finally
        {
            if (document is not null) { try { document.Close(0); } finally { Release(document); } }
            if (application is not null) { try { application.Quit(0); } finally { Release(application); } }
        }
    }

    private static void ConvertExcel(string sourcePath, string targetPath, CancellationToken token)
    {
        var srcExt = Path.GetExtension(sourcePath).ToLowerInvariant();
        var tgtExt = Path.GetExtension(targetPath).ToLowerInvariant();

        // CSV <-> TSV için doğrudan güvenilir dönüşüm (bölgesel ayar sorunlarından etkilenmez)
        if ((srcExt is ".csv" && tgtExt is ".tsv") || (srcExt is ".tsv" && tgtExt is ".csv"))
        {
            ConvertDelimitedDirect(sourcePath, targetPath, srcExt is ".csv" ? ',' : '\t', tgtExt is ".csv" ? ',' : '\t');
            return;
        }

        dynamic? application = null;
        dynamic? workbook = null;
        try
        {
            application = StartOffice("Excel.Application");
            application.Visible = false;
            application.DisplayAlerts = false;
            application.AutomationSecurity = 3;
            token.ThrowIfCancellationRequested();
            workbook = application.Workbooks.Open(sourcePath, UpdateLinks: 0, ReadOnly: true);
            token.ThrowIfCancellationRequested();
            var extension = Path.GetExtension(targetPath).ToLowerInvariant();
            if (extension == ".pdf") workbook.ExportAsFixedFormat(0, targetPath);
            else workbook.SaveAs(targetPath, extension switch
            {
                ".xls" => 56, ".xlsx" => 51, ".xlsm" => 52,
                ".xlsb" => 50, ".ods" => 60, ".csv" => 62,
                ".tsv" => 42,
                _ => throw new NotSupportedException(extension)
            });
        }
        finally
        {
            if (workbook is not null) { try { workbook.Close(false); } finally { Release(workbook); } }
            if (application is not null) { try { application.Quit(); } finally { Release(application); } }
        }
    }

    private static void ConvertDelimitedDirect(string sourcePath, string targetPath, char fromSep, char toSep)
    {
        using var reader = new StreamReader(sourcePath, Encoding.UTF8, true);
        using var writer = new StreamWriter(targetPath, false, Encoding.UTF8);
        string? line;
        while ((line = reader.ReadLine()) is not null)
        {
            var fields = ParseDelimitedLine(line, fromSep);
            writer.WriteLine(string.Join(toSep, fields.Select(f => FormatDelimitedField(f, toSep))));
        }
    }

    private static List<string> ParseDelimitedLine(string line, char sep)
    {
        var list = new List<string>();
        var sb = new StringBuilder();
        var inQuotes = false;
        for (var i = 0; i < line.Length; i++)
        {
            var c = line[i];
            if (c == '"')
            {
                if (inQuotes && i + 1 < line.Length && line[i + 1] == '"')
                {
                    sb.Append('"');
                    i++;
                }
                else
                {
                    inQuotes = !inQuotes;
                }
            }
            else if (c == sep && !inQuotes)
            {
                list.Add(sb.ToString());
                sb.Clear();
            }
            else
            {
                sb.Append(c);
            }
        }
        list.Add(sb.ToString());
        return list;
    }

    private static string FormatDelimitedField(string field, char sep)
    {
        if (field.Contains(sep) || field.Contains('"') || field.Contains('\n') || field.Contains('\r'))
            return "\"" + field.Replace("\"", "\"\"") + "\"";
        return field;
    }

    private static void ConvertPowerPoint(string sourcePath, string targetPath, CancellationToken token)
    {
        dynamic? application = null;
        dynamic? presentation = null;
        try
        {
            application = StartOffice("PowerPoint.Application");
            application.AutomationSecurity = 3;
            token.ThrowIfCancellationRequested();
            presentation = application.Presentations.Open(sourcePath, ReadOnly: -1, Untitled: 0, WithWindow: 0);
            token.ThrowIfCancellationRequested();
            var extension = Path.GetExtension(targetPath).ToLowerInvariant();
            presentation.SaveAs(targetPath, extension switch
            {
                ".ppt" => 1, ".pptx" => 24, ".pptm" => 25,
                ".pps" => 7, ".ppsx" => 28, ".odp" => 35, ".pdf" => 32,
                _ => throw new NotSupportedException(extension)
            });
        }
        finally
        {
            if (presentation is not null) { try { presentation.Close(); } finally { Release(presentation); } }
            if (application is not null) { try { application.Quit(); } finally { Release(application); } }
        }
    }

    private static void ConvertImage(string sourcePath, string targetPath)
    {
        var extension = Path.GetExtension(sourcePath).ToLowerInvariant();
        using var source = extension is ".heic" or ".heif"
            ? DecodeHeif(sourcePath)
            : Image.FromFile(sourcePath);
        using var bitmap = new Bitmap(source);
        var targetExt = Path.GetExtension(targetPath).ToLowerInvariant();
        if (targetExt == ".ico")
        {
            SaveAsIcon(bitmap, targetPath);
            return;
        }
        var format = targetExt switch
        {
            ".png" => ImageFormat.Png,
            ".jpg" or ".jpeg" => ImageFormat.Jpeg,
            ".bmp" => ImageFormat.Bmp,
            ".gif" => ImageFormat.Gif,
            ".tif" or ".tiff" => ImageFormat.Tiff,
            _ => throw new NotSupportedException("Görsel hedef biçimi desteklenmiyor.")
        };
        bitmap.Save(targetPath, format);
    }

    private static void SaveAsIcon(Bitmap bitmap, string targetPath)
    {
        using var pngStream = new MemoryStream();
        bitmap.Save(pngStream, ImageFormat.Png);
        var pngBytes = pngStream.ToArray();

        using var fileStream = File.Create(targetPath);
        using var writer = new BinaryWriter(fileStream);
        writer.Write((ushort)0); // Reserved
        writer.Write((ushort)1); // Type: 1 = icon
        writer.Write((ushort)1); // Count: 1 image

        writer.Write((byte)(bitmap.Width >= 256 ? 0 : bitmap.Width));
        writer.Write((byte)(bitmap.Height >= 256 ? 0 : bitmap.Height));
        writer.Write((byte)0); // Color palette
        writer.Write((byte)0); // Reserved
        writer.Write((ushort)1); // Color planes
        writer.Write((ushort)32); // Bits per pixel
        writer.Write((uint)pngBytes.Length); // Image data size
        writer.Write((uint)22); // Offset to image data

        writer.Write(pngBytes);
    }

    private static Image DecodeHeif(string sourcePath)
    {
        try
        {
            var decoder = BitmapDecoder.Create(new Uri(sourcePath), BitmapCreateOptions.None, BitmapCacheOption.OnLoad);
            if (decoder.Frames.Count == 0) throw new InvalidDataException("Dosyada görüntü bulunamadı.");
            var encoder = new PngBitmapEncoder();
            encoder.Frames.Add(BitmapFrame.Create(decoder.Frames[0]));
            using var stream = new MemoryStream();
            encoder.Save(stream);
            stream.Position = 0;
            using var decoded = Image.FromStream(stream);
            return new Bitmap(decoded);
        }
        catch (Exception exception) when (exception is COMException or NotSupportedException or FileFormatException)
        {
            throw new InvalidOperationException(
                "HEIC/HEIF dosyası açılamadı. Windows HEIF Görüntü Uzantıları ve gerekiyorsa HEVC codec'i kurulu olmalıdır; dosya da bozuk olabilir.",
                exception);
        }
    }

    private static void Release(object comObject)
    {
        if (Marshal.IsComObject(comObject)) Marshal.FinalReleaseComObject(comObject);
    }
}
