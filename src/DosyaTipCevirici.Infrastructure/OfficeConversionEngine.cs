using System.Drawing;
using System.Drawing.Imaging;
using System.IO;
using System.Runtime.InteropServices;
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
                    ConvertWord(sourcePath, targetPath, cancellationToken);
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

    private static void ConvertWord(string sourcePath, string targetPath, CancellationToken token)
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
            else if (extension == ".txt") document.SaveAs2(targetPath, 2, Encoding: 65001);
            else document.SaveAs2(targetPath, extension switch
            {
                ".doc" => 0, ".docx" => 16, ".docm" => 13,
                ".odt" => 23, ".rtf" => 6, ".html" or ".htm" => 8,
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
                _ => throw new NotSupportedException(extension)
            });
        }
        finally
        {
            if (workbook is not null) { try { workbook.Close(false); } finally { Release(workbook); } }
            if (application is not null) { try { application.Quit(); } finally { Release(application); } }
        }
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
        var format = Path.GetExtension(targetPath).ToLowerInvariant() switch
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
