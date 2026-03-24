using System;
using System.Runtime.InteropServices;
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
                    ConvertSingleEquationToImage(equations, i, result);
                }
            }
            finally
            {
                ReleaseComObject(equations);
            }

            return result;
        }

        private static void ConvertSingleEquationToImage(Word.OMaths equations, int index, EquationImageConversionResult result)
        {
            Word.OMath equation = null;
            Word.Range equationRange = null;
            Word.Range pasteRange = null;

            try
            {
                equation = equations[index];
                equationRange = equation.Range.Duplicate;
                pasteRange = equationRange.Duplicate;

                // 수식을 그림으로 복사한 뒤, 수식 끝 지점에 EMF로 붙여넣습니다.
                equationRange.CopyAsPicture();
                pasteRange.Collapse(Word.WdCollapseDirection.wdCollapseEnd);
                pasteRange.PasteSpecial(DataType: Word.WdPasteDataType.wdPasteEnhancedMetafile);

                equationRange.Delete();
                result.SuccessCount++;
            }
            catch
            {
                result.FailureCount++;
            }
            finally
            {
                ReleaseComObject(pasteRange);
                ReleaseComObject(equationRange);
                ReleaseComObject(equation);
            }
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
