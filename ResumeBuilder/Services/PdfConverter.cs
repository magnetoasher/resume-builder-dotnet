using System;
using System.Diagnostics;
using System.IO;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading;
using System.Threading.Tasks;

namespace ResumeBuilder.Services;

public static class PdfConverter
{
    private const int WdExportFormatPdf = 17;
    private const int ConvertTimeoutMilliseconds = 90000;

    public static Task ConvertDocxToPdfAsync(string docxPath, string pdfPath)
    {
        var tcs = new TaskCompletionSource<bool>();
        var thread = new Thread(() =>
        {
            try
            {
                ConvertDocxToPdf(docxPath, pdfPath);
                tcs.SetResult(true);
            }
            catch (Exception ex)
            {
                tcs.SetException(ex);
            }
        });
        thread.SetApartmentState(ApartmentState.STA);
        thread.Start();
        return tcs.Task;
    }

    private static void ConvertDocxToPdf(string docxPath, string pdfPath)
    {
        Exception? wordException = null;
        try
        {
            ConvertWithWord(docxPath, pdfPath);
            if (File.Exists(pdfPath))
            {
                return;
            }
        }
        catch (Exception ex)
        {
            wordException = ex;
        }

        if (TryConvertWithLibreOffice(docxPath, pdfPath, out var libreOfficeException))
        {
            return;
        }

        var message = new StringBuilder("PDF export failed. Install Microsoft Word or LibreOffice.");
        if (wordException != null)
        {
            message.Append($" Word: {wordException.Message}");
        }

        if (libreOfficeException != null)
        {
            message.Append($" LibreOffice: {libreOfficeException.Message}");
        }

        throw new InvalidOperationException(message.ToString());
    }

    private static void ConvertWithWord(string docxPath, string pdfPath)
    {
        object? wordApp = null;
        object? doc = null;
        try
        {
            var wordType = Type.GetTypeFromProgID("Word.Application");
            if (wordType == null)
            {
                throw new InvalidOperationException("Microsoft Word is not installed.");
            }

            wordApp = Activator.CreateInstance(wordType)
                ?? throw new InvalidOperationException("Unable to start Microsoft Word.");

            dynamic app = wordApp;
            app.Visible = false;
            doc = app.Documents.Open(docxPath, ReadOnly: true, Visible: false);
            ((dynamic)doc).ExportAsFixedFormat(pdfPath, WdExportFormatPdf);
        }
        catch (COMException ex)
        {
            throw new InvalidOperationException($"Word export failed: {ex.Message}", ex);
        }
        finally
        {
            if (doc != null)
            {
                try
                {
                    ((dynamic)doc).Close(false);
                }
                catch
                {
                    // Best effort cleanup.
                }

                Marshal.FinalReleaseComObject(doc);
            }

            if (wordApp != null)
            {
                try
                {
                    ((dynamic)wordApp).Quit(false);
                }
                catch
                {
                    // Best effort cleanup.
                }

                Marshal.FinalReleaseComObject(wordApp);
            }
        }
    }

    private static bool TryConvertWithLibreOffice(string docxPath, string pdfPath, out Exception? error)
    {
        error = null;
        var outDir = Path.GetDirectoryName(pdfPath);
        if (string.IsNullOrWhiteSpace(outDir))
        {
            error = new InvalidOperationException("Output directory is missing.");
            return false;
        }

        try
        {
            Directory.CreateDirectory(outDir);
            foreach (var executable in EnumerateSofficeExecutables())
            {
                if (TryRunSoffice(executable, docxPath, outDir, out var runError))
                {
                    var expectedPdf = Path.Combine(outDir, $"{Path.GetFileNameWithoutExtension(docxPath)}.pdf");
                    if (!File.Exists(expectedPdf))
                    {
                        continue;
                    }

                    if (!string.Equals(expectedPdf, pdfPath, StringComparison.OrdinalIgnoreCase))
                    {
                        if (File.Exists(pdfPath))
                        {
                            File.Delete(pdfPath);
                        }

                        File.Move(expectedPdf, pdfPath);
                    }

                    return true;
                }

                error = runError;
            }

            error ??= new InvalidOperationException("LibreOffice (soffice) not found.");
            return false;
        }
        catch (Exception ex)
        {
            error = ex;
            return false;
        }
    }

    private static string[] EnumerateSofficeExecutables()
    {
        var programFiles = Environment.GetFolderPath(Environment.SpecialFolder.ProgramFiles);
        var programFilesX86 = Environment.GetFolderPath(Environment.SpecialFolder.ProgramFilesX86);
        return
        [
            "soffice.exe",
            Path.Combine(programFiles, "LibreOffice", "program", "soffice.exe"),
            Path.Combine(programFilesX86, "LibreOffice", "program", "soffice.exe")
        ];
    }

    private static bool TryRunSoffice(string executable, string docxPath, string outDir, out Exception? error)
    {
        error = null;
        try
        {
            var info = new ProcessStartInfo
            {
                FileName = executable,
                Arguments = $"--headless --nologo --norestore --convert-to pdf --outdir \"{outDir}\" \"{docxPath}\"",
                UseShellExecute = false,
                RedirectStandardError = true,
                RedirectStandardOutput = true,
                CreateNoWindow = true
            };

            using var process = Process.Start(info);
            if (process == null)
            {
                error = new InvalidOperationException($"Failed to start converter: {executable}");
                return false;
            }

            if (!process.WaitForExit(ConvertTimeoutMilliseconds))
            {
                try
                {
                    process.Kill(entireProcessTree: true);
                }
                catch
                {
                    // Ignore kill failures.
                }

                error = new TimeoutException("LibreOffice conversion timed out.");
                return false;
            }

            if (process.ExitCode == 0)
            {
                return true;
            }

            var stderr = process.StandardError.ReadToEnd();
            var stdout = process.StandardOutput.ReadToEnd();
            var details = string.IsNullOrWhiteSpace(stderr) ? stdout : stderr;
            error = new InvalidOperationException($"LibreOffice exited with code {process.ExitCode}. {details}".Trim());
            return false;
        }
        catch (Exception ex)
        {
            error = ex;
            return false;
        }
    }
}
