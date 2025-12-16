using System;
using System.Text.RegularExpressions;
using System.Windows.Forms;
using Word = Microsoft.Office.Interop.Word;

namespace WordAddIn1
{
    public static class DrawingReferenceRemover
    {
        public static void RemoveDrawingReferences()
        {
            try
            {
                Word.Application app = Globals.ThisAddIn.Application;
                Word.Selection selection = app.Selection;
                Word.Range selRange = selection.Range;

                string selectedText = selRange.Text;
                if (string.IsNullOrWhiteSpace(selectedText))
                {
                    MessageBox.Show("선택된 텍스트가 없습니다.", "알림", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    return;
                }

                // ✅ 메모 삭제 확인
                DialogResult result = MessageBox.Show(
                    "선택된 영역에서 메모가 삽입되어 있을 경우 오류가 발생할 수 있습니다. 메모를 삭제한 후 진행하시겠습니까?",
                    "메모 삭제 확인",
                    MessageBoxButtons.YesNo,
                    MessageBoxIcon.Question);

                if (result == DialogResult.Yes)
                {
                    CommentRemover.RemoveCommentsInSelection(selection);
                }


                Regex regex = new Regex(@"\(([0-9]+[A-Za-z]?)(,\s*[0-9]+[A-Za-z]?)*\)", RegexOptions.IgnoreCase);
                MatchCollection matches = regex.Matches(selectedText);

                if (matches.Count == 0)
                {
                    MessageBox.Show("도면부호 패턴이 발견되지 않았습니다.", "알림", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    return;
                }

                for (int i = matches.Count - 1; i >= 0; i--)
                {
                    Match match = matches[i];
                    int start = selRange.Start + match.Index;
                    int end = start + match.Length;

                    Word.Range toDelete = selRange.Document.Range(start, end);
                    toDelete.Delete();
                }

                //                 MessageBox.Show("도면부호가 삭제되었습니다.", "완료", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            catch (Exception ex)
            {
                MessageBox.Show("오류 발생: " + ex.Message, "오류", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }
    }
}