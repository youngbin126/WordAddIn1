using System;
using System.Runtime.InteropServices;
using Word = Microsoft.Office.Interop.Word;

namespace WordAddIn1
{
    public static class EquationImageTransformer
    {
        public static void ConvertAllEquationsToImages(Word.Document doc)
        {
            if (doc == null)
            {
                throw new ArgumentNullException(nameof(doc));
            }

            Word.OMaths oMaths = null;

            try
            {
                oMaths = doc.OMaths;

                for (int i = oMaths.Count; i >= 1; i--)
                {
                    Word.OMath oMath = null;
                    Word.Range eqRange = null;
                    Word.Range temp = null;
                    Word.InlineShape pasted = null;

                    try
                    {
                        oMath = oMaths[i];
                        eqRange = oMath.Range.Duplicate;
                        bool isDisplayEquation = oMath.Type == Word.WdOMathType.wdOMathDisplay;

                        eqRange.CopyAsPicture();

                        temp = eqRange.Duplicate;
                        PreparePasteRange(eqRange, temp, isDisplayEquation);

                        int inlineShapeCountBeforePaste = doc.InlineShapes.Count;
                        temp.Paste();
                        pasted = GetPastedInlineShape(temp, doc, inlineShapeCountBeforePaste);

                        eqRange.Delete();

                        if (pasted != null)
                        {
                            Word.Range pastedRange = null;

                            try
                            {
                                pastedRange = pasted.Range;
                                pastedRange.Collapse(Word.WdCollapseDirection.wdCollapseStart);
                            }
                            finally
                            {
                                ReleaseComObject(pastedRange);
                            }
                        }
                    }
                    finally
                    {
                        ReleaseComObject(pasted);
                        ReleaseComObject(temp);
                        ReleaseComObject(eqRange);
                        ReleaseComObject(oMath);
                    }
                }
            }
            finally
            {
                ReleaseComObject(oMaths);
            }
        }

        private static void PreparePasteRange(Word.Range eqRange, Word.Range temp, bool isDisplayEquation)
        {
            if (isDisplayEquation)
            {
                int insertionPoint = Math.Max(eqRange.Start, eqRange.End - 1);
                temp.SetRange(insertionPoint, insertionPoint);
            }
            else
            {
                temp.Collapse(Word.WdCollapseDirection.wdCollapseEnd);
            }
        }

        private static Word.InlineShape GetPastedInlineShape(Word.Range temp, Word.Document doc, int inlineShapeCountBeforePaste)
        {
            if (temp.InlineShapes.Count > 0)
            {
                return temp.InlineShapes[1];
            }

            int inlineShapeCountAfterPaste = doc.InlineShapes.Count;
            if (inlineShapeCountAfterPaste > inlineShapeCountBeforePaste)
            {
                return doc.InlineShapes[inlineShapeCountAfterPaste];
            }

            return null;
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
