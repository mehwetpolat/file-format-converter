namespace DosyaTipCevirici.Core;

public interface IConversionEngine
{
    Task ConvertAsync(string sourcePath, string targetPath, CancellationToken cancellationToken);
}

public sealed class ConversionService(IConversionEngine engine)
{
    public async Task<string> ConvertAsync(string sourcePath, string outputDirectory, string targetExtension,
        CancellationToken cancellationToken = default)
    {
        ArgumentException.ThrowIfNullOrWhiteSpace(sourcePath);
        ArgumentException.ThrowIfNullOrWhiteSpace(outputDirectory);
        if (!File.Exists(sourcePath)) throw new FileNotFoundException("Kaynak dosya bulunamadı.", sourcePath);

        var source = ConversionCatalog.Find(sourcePath)
            ?? throw new NotSupportedException("Kaynak dosya biçimi desteklenmiyor.");
        var target = ConversionCatalog.Find("file." + targetExtension.TrimStart('.'))
            ?? throw new NotSupportedException("Hedef dosya biçimi desteklenmiyor.");
        if (!ConversionCatalog.CanConvert(source, target))
            throw new NotSupportedException($"{source.Extension} → {target.Extension} dönüşümü desteklenmiyor.");

        Directory.CreateDirectory(outputDirectory);
        var outputPath = Path.Combine(outputDirectory, Path.GetFileNameWithoutExtension(sourcePath) + "." + target.Extension);
        for (var number = 2; File.Exists(outputPath); number++)
            outputPath = Path.Combine(outputDirectory, $"{Path.GetFileNameWithoutExtension(sourcePath)} ({number}).{target.Extension}");

        await engine.ConvertAsync(sourcePath, outputPath, cancellationToken);
        if (!File.Exists(outputPath) || new FileInfo(outputPath).Length == 0)
            throw new IOException("Dönüştürme tamamlandı ancak geçerli bir çıktı oluşmadı.");
        return outputPath;
    }
}
