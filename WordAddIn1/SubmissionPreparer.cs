using System;
using System.Text.RegularExpressions;
using System.Windows.Forms;
using Word = Microsoft.Office.Interop.Word;

namespace WordAddIn1
{
    public class SubmissionPreparer
    {
        public static void PrepareSubmission()
        {
            // DrawingReferenceRemover의 메서드를 호출
            DrawingReferenceRemover.RemoveDrawingReferences();
        }
    }
}
