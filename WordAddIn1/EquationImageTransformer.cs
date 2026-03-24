using System;
using System.Drawing;
using System.Drawing.Imaging;
using System.IO;
using System.Runtime.InteropServices;
using System.Threading;
using System.Windows.Forms;
using Word = Microsoft.Office.Interop.Word;

namespace WordAddIn1
{
    internal sealed class EquationImageConversionResult
    {
        public int SuccessCount { get; set; }

        public int FailureCount { get; set; }
    }

    internal static class EquationImageTransformer
    {
        private const int MaxClipboardRetryCount = 3;

        public static EquationImageConversionResult ConvertAllEquationsToImages(Word.Document doc)
        {
            var result = new EquationImageConversionResult();

            if (doc == null)
            {
                throw new ArgumentNullException(nameof(doc));
            }

            Word.OMaths equations = null;

            try
            {
                equations = doc.OMaths;
                if (equations == null || equations.Count == 0)
                {
                    return result;
                }

                for (int i = equations.Count; i >= 1; i--)
                {
                    ConvertSingleEquationToImage(doc, equations, i, result);
                }
            }
            finally
            {
                ReleaseComObject(equations);
            }

            return result;
        }

        private static void ConvertSingleEquationToImage(Word.Document doc, Word.OMaths equations, int index, EquationImageConversionResult result)
        {
            Word.OMath equation = null;
            Word.Range equationRange = null;
            Word.Range imageSourceRange = null;
            Word.Range insertRange = null;
            Word.InlineShape insertedImage = null;

            string tempPath = null;

            try
            {
                equation = equations[index];
                equationRange = equation.Range.Duplicate;
                imageSourceRange = CreateTrimmedRange(equationRange);

                if (imageSourceRange == null || imageSourceRange.Start >= imageSourceRange.End)
                {
                    throw new InvalidOperationException("수식 범위를 찾을 수 없습니다.");
                }

                tempPath = Path.Combine(Path.GetTempPath(), $"eq_{Guid.NewGuid():N}.tiff");

                if (!TrySaveRangeAsTiff(imageSourceRange, tempPath))
                {
                    throw new InvalidOperationException("클립보드에서 수식 이미지를 가져오지 못했습니다.");
                }

                int insertionPoint = equationRange.Start;
                equationRange.Delete();

                insertRange = doc.Range(insertionPoint, insertionPoint);
                insertedImage = insertRange.InlineShapes.AddPicture(tempPath, LinkToFile: false, SaveWithDocument: true);

                result.SuccessCount++;
            }
            catch
            {
                result.FailureCount++;
            }
            finally
            {
                if (!string.IsNullOrEmpty(tempPath) && File.Exists(tempPath))
                {
                    File.Delete(tempPath);
                }

                ReleaseComObject(insertedImage);
                ReleaseComObject(insertRange);
                ReleaseComObject(imageSourceRange);
                ReleaseComObject(equationRange);
                ReleaseComObject(equation);
            }
        }

        private static bool TrySaveRangeAsTiff(Word.Range sourceRange, string outputPath)
        {
            for (int attempt = 1; attempt <= MaxClipboardRetryCount; attempt++)
            {
                try
                {
                    sourceRange.CopyAsPicture();
                    Thread.Sleep(80);

                    if (!Clipboard.ContainsImage())
                    {
                        continue;
                    }

                    using (Image clipboardImage = Clipboard.GetImage())
                    using (Bitmap trimmedBitmap = TrimWhitespace(new Bitmap(clipboardImage)))
                    {
                        trimmedBitmap.Save(outputPath, ImageFormat.Tiff);
                    }

                    return true;
                }
                catch (ExternalException)
                {
                    Thread.Sleep(120);
                }
            }

            return false;
        }

        private static Word.Range CreateTrimmedRange(Word.Range equationRange)
        {
            Word.Document doc = equationRange.Document;
            int start = equationRange.Start;
            int end = equationRange.End;

            while (end > start && IsTrimChar(doc, end - 1))
            {
                end--;
            }

            while (start < end && IsTrimChar(doc, start))
            {
                start++;
            }

            return doc.Range(start, end);
        }

        private static bool IsTrimChar(Word.Document doc, int position)
        {
            Word.Range charRange = null;

            try
            {
                charRange = doc.Range(position, position + 1);
                string text = charRange.Text;

                if (string.IsNullOrEmpty(text))
                {
                    return true;
                }

                return text == "\r" || text == "\a" || text == "\t" || text == " " || text == "\n";
            }
            finally
            {
                ReleaseComObject(charRange);
            }
        }

        private static Bitmap TrimWhitespace(Bitmap source)
        {
            int minX = source.Width;
            int minY = source.Height;
            int maxX = -1;
            int maxY = -1;

            for (int y = 0; y < source.Height; y++)
            {
                for (int x = 0; x < source.Width; x++)
                {
                    Color pixel = source.GetPixel(x, y);
                    bool hasContent = pixel.A > 10 && !(pixel.R > 245 && pixel.G > 245 && pixel.B > 245);

                    if (!hasContent)
                    {
                        continue;
                    }

                    if (x < minX) minX = x;
                    if (y < minY) minY = y;
                    if (x > maxX) maxX = x;
                    if (y > maxY) maxY = y;
                }
            }

            if (maxX < minX || maxY < minY)
            {
                return new Bitmap(source);
            }

            Rectangle cropArea = Rectangle.FromLTRB(minX, minY, maxX + 1, maxY + 1);
            return source.Clone(cropArea, source.PixelFormat);
        }

        private static void ReleaseComObject(object comObject)
        {
            if (comObject != null && Marshal.IsComObject(comObject))
            {
                Marshal.ReleaseComObject(comObject);
            }
        }
    }
}
