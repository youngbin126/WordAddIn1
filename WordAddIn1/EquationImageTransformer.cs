using System;
using System.Drawing;
using System.Drawing.Imaging;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using System.Windows.Forms;
using Word = Microsoft.Office.Interop.Word;

namespace WordAddIn1
{
    internal class EquationImageConversionResult
    {
        public int SuccessCount { get; set; }
        public int FailureCount { get; set; }
    }

    internal static class EquationImageTransformer
    {
        public static EquationImageConversionResult ConvertAllEquationsToImages(Word.Document doc)
        {
            var result = new EquationImageConversionResult();

            if (doc == null)
            {
                return result;
            }

            Word.OMaths equations = doc.OMaths;
            for (int i = equations.Count; i >= 1; i--)
            {
                Word.OMath equation = null;
                Word.Range equationRange = null;

                try
                {
                    equation = equations[i];
                    equationRange = equation.Range;

                    equationRange.CopyAsPicture();
                    equationRange.Collapse(Word.WdCollapseDirection.wdCollapseEnd);
                    equationRange.PasteSpecial(DataType: Word.WdPasteDataType.wdPasteEnhancedMetafile);

                    equation.Range.Delete();
                    result.SuccessCount++;
                }
                catch
                {
                    result.FailureCount++;
                }
                finally
                {
                    if (equationRange != null)
                    {
                        System.Runtime.InteropServices.Marshal.ReleaseComObject(equationRange);
                    }

                    if (equation != null)
                    {
                        System.Runtime.InteropServices.Marshal.ReleaseComObject(equation);
                    }
                }
            }

            if (equations != null)
            {
                System.Runtime.InteropServices.Marshal.ReleaseComObject(equations);
            }

            return result;
    public static class EquationImageTransformer
    {
        public static void TransformAllEquationsToCcittGroup4Tiff(Word.Document doc, int dpi = 300)
        {
            if (doc == null)
            {
                MessageBox.Show("활성 문서를 찾을 수 없습니다.", "수식 변환", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            int targetDpi = dpi == 400 ? 400 : 300;
            Word.OMaths equations = null;

            try
            {
                equations = doc.OMaths;
                if (equations == null || equations.Count == 0)
                {
                    MessageBox.Show("변환할 수식이 없습니다.", "수식 변환", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    return;
                }

                for (int i = equations.Count; i >= 1; i--)
                {
                    TransformEquationAtIndex(doc, equations, i, targetDpi);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"수식 이미지 변환 중 오류가 발생했습니다.\n{ex.Message}", "수식 변환", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            finally
            {
                if (equations != null)
                {
                    Marshal.ReleaseComObject(equations);
                }
            }
        }

        private static void TransformEquationAtIndex(Word.Document doc, Word.OMaths equations, int index, int dpi)
        {
            string tempPngPath = null;
            string tempEmfPath = null;
            string tempTiffPath = null;

            Word.OMath equation = null;
            Word.Range equationRange = null;
            Word.Range replacementRange = null;
            Word.InlineShape insertedImage = null;

            try
            {
                equation = equations[index];
                equationRange = equation.Range.Duplicate;

                equationRange.Copy();
                equationRange.PasteSpecial(
                    Link: false,
                    DataType: Word.WdPasteDataType.wdPasteEnhancedMetafile,
                    Placement: Word.WdOLEPlacement.wdInLine,
                    DisplayAsIcon: false);

                if (equationRange.InlineShapes.Count == 0)
                {
                    throw new InvalidOperationException("수식을 임시 이미지(EMF)로 변환하지 못했습니다.");
                }

                insertedImage = equationRange.InlineShapes[1];
                replacementRange = insertedImage.Range.Duplicate;

                tempPngPath = Path.Combine(Path.GetTempPath(), $"eq_{Guid.NewGuid():N}.png");
                tempEmfPath = Path.Combine(Path.GetTempPath(), $"eq_{Guid.NewGuid():N}.emf");
                tempTiffPath = Path.Combine(Path.GetTempPath(), $"eq_{Guid.NewGuid():N}.tif");

                insertedImage.SaveAsPicture(tempPngPath);
                insertedImage.SaveAsPicture(tempEmfPath);

                SaveAsCcittGroup4Tiff(tempPngPath, tempEmfPath, tempTiffPath, dpi);

                insertedImage.Delete();
                replacementRange.InlineShapes.AddPicture(tempTiffPath, false, true);
            }
            catch (Exception ex)
            {
                MessageBox.Show($"{index}번째 수식 처리 중 오류가 발생했습니다.\n{ex.Message}", "수식 변환", MessageBoxButtons.OK, MessageBoxIcon.Warning);
            }
            finally
            {
                TryDeleteTempFile(tempPngPath);
                TryDeleteTempFile(tempEmfPath);
                TryDeleteTempFile(tempTiffPath);

                if (insertedImage != null)
                {
                    Marshal.ReleaseComObject(insertedImage);
                }
                if (replacementRange != null)
                {
                    Marshal.ReleaseComObject(replacementRange);
                }
                if (equationRange != null)
                {
                    Marshal.ReleaseComObject(equationRange);
                }
                if (equation != null)
                {
                    Marshal.ReleaseComObject(equation);
                }
            }
        }

        public static void SaveAsCcittGroup4Tiff(Image src, string path, int dpi = 300)
        {
            if (src == null) throw new ArgumentNullException(nameof(src));
            if (string.IsNullOrWhiteSpace(path)) throw new ArgumentException("저장 경로가 비어 있습니다.", nameof(path));

            int targetDpi = dpi == 400 ? 400 : 300;
            using (Bitmap oneBit = ConvertTo1Bpp(src))
            {
                oneBit.SetResolution(targetDpi, targetDpi);

                ImageCodecInfo tiffCodec = ImageCodecInfo
                    .GetImageEncoders()
                    .FirstOrDefault(codec => codec.FormatID == ImageFormat.Tiff.Guid);

                if (tiffCodec == null)
                {
                    throw new InvalidOperationException("TIFF 인코더를 찾을 수 없습니다.");
                }

                using (EncoderParameters encoderParams = new EncoderParameters(1))
                {
                    encoderParams.Param[0] = new EncoderParameter(System.Drawing.Imaging.Encoder.Compression, (long)EncoderValue.CompressionCCITT4);
                    oneBit.Save(path, tiffCodec, encoderParams);
                }
            }
        }

        private static void SaveAsCcittGroup4Tiff(string pngPath, string emfPath, string targetTiffPath, int dpi)
        {
            if (!string.IsNullOrWhiteSpace(pngPath) && File.Exists(pngPath))
            {
                using (Image src = Image.FromFile(pngPath))
                {
                    SaveAsCcittGroup4Tiff(src, targetTiffPath, dpi);
                    return;
                }
            }

            if (!string.IsNullOrWhiteSpace(emfPath) && File.Exists(emfPath))
            {
                using (Image src = Image.FromFile(emfPath))
                {
                    SaveAsCcittGroup4Tiff(src, targetTiffPath, dpi);
                    return;
                }
            }

            throw new FileNotFoundException("수식 이미지(PNG/EMF) 임시 파일을 찾을 수 없습니다.");
        }

        private static Bitmap ConvertTo1Bpp(Image src)
        {
            using (Bitmap graySource = new Bitmap(src.Width, src.Height, PixelFormat.Format24bppRgb))
            {
                using (Graphics g = Graphics.FromImage(graySource))
                {
                    g.Clear(Color.White);
                    g.DrawImage(src, 0, 0, src.Width, src.Height);
                }

                Bitmap oneBit = new Bitmap(graySource.Width, graySource.Height, PixelFormat.Format1bppIndexed);
                Rectangle rect = new Rectangle(0, 0, graySource.Width, graySource.Height);

                BitmapData srcData = graySource.LockBits(rect, ImageLockMode.ReadOnly, PixelFormat.Format24bppRgb);
                BitmapData dstData = oneBit.LockBits(rect, ImageLockMode.WriteOnly, PixelFormat.Format1bppIndexed);

                try
                {
                    int srcStride = srcData.Stride;
                    int dstStride = dstData.Stride;
                    int width = graySource.Width;
                    int height = graySource.Height;
                    int threshold = 180;

                    byte[] srcBytes = new byte[srcStride * height];
                    byte[] dstBytes = new byte[dstStride * height];

                    Marshal.Copy(srcData.Scan0, srcBytes, 0, srcBytes.Length);

                    for (int y = 0; y < height; y++)
                    {
                        int srcRow = y * srcStride;
                        int dstRow = y * dstStride;

                        for (int x = 0; x < width; x++)
                        {
                            int srcIndex = srcRow + (x * 3);
                            byte b = srcBytes[srcIndex];
                            byte g = srcBytes[srcIndex + 1];
                            byte r = srcBytes[srcIndex + 2];
                            int luminance = (r * 299 + g * 587 + b * 114) / 1000;

                            if (luminance < threshold)
                            {
                                dstBytes[dstRow + (x >> 3)] |= (byte)(0x80 >> (x & 0x7));
                            }
                        }
                    }

                    Marshal.Copy(dstBytes, 0, dstData.Scan0, dstBytes.Length);
                }
                finally
                {
                    graySource.UnlockBits(srcData);
                    oneBit.UnlockBits(dstData);
                }

                return oneBit;
            }
        }

        private static void TryDeleteTempFile(string path)
        {
            if (string.IsNullOrWhiteSpace(path) || !File.Exists(path))
            {
                return;
            }

            try
            {
                File.Delete(path);
            }
            catch
            {
                // 임시 파일 정리 실패는 무시
            }
        }
    }
}
