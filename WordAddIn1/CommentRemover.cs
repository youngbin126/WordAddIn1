using System;
using System.Text.RegularExpressions;
using System.Windows.Forms;
using Word = Microsoft.Office.Interop.Word;

namespace WordAddIn1
{
    public static class CommentRemover
    {
        /// 지정된 범위(selection)의 모든 메모를 삭제
        /// <param name="selection">삭제할 범위</param>
        public static void RemoveCommentsInSelection(Word.Selection selection)
        {
            if (selection != null && selection.Comments.Count > 0)
            {
                for (int i = selection.Comments.Count; i >= 1; i--)
                {
                    selection.Comments[i].Delete();
                }

            }
            else
            {
                MessageBox.Show("선택된 부분에 메모가 없습니다.");
            }
        }
    }

}

