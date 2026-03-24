using System;
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
        }
    }
}
