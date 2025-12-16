using Microsoft.Office.Interop.Word;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Word = Microsoft.Office.Interop.Word;

namespace WordAddIn1
{
    public static class TextChanger 
    {
        public static void ReplaceText(Word.Range range, string oldText, string newText)
        {
            Word.Find findObject = range.Find;
            findObject.ClearFormatting();
            findObject.Text = oldText;
            findObject.Replacement.ClearFormatting();
            findObject.Replacement.Text = newText;

            object replaceAll = Word.WdReplace.wdReplaceAll;
            findObject.Execute(Replace: ref replaceAll);
        }

        public static void DeleteLinesContainingText(Word.Document doc, string target, bool deleteNextLine)
        {
            // 컬렉션에서 항목을 삭제할 때는 반드시 역방향으로 반복해야 합니다.
            for (int i = doc.Paragraphs.Count; i >= 1; i--)
            {
                // 현재 단락의 텍스트를 가져옵니다.
                string text = doc.Paragraphs[i].Range.Text.Trim();

                // 목표 텍스트가 포함되어 있는지 확인합니다.
                if (text.Contains(target))
                {
                    // '다음 줄 삭제' 옵션이 켜져 있고, 현재 단락이 마지막 단락이 아닐 경우
                    if (deleteNextLine && i < doc.Paragraphs.Count)
                    {
                        // 중요: 다음 단락(i + 1)을 *먼저* 삭제합니다.
                        // 이 작업은 i번째 단락의 인덱스에 영향을 주지 않습니다.
                        doc.Paragraphs[i + 1].Range.Delete();
                    }

                    // 이제 현재 단락(i)을 안전하게 삭제합니다.
                    doc.Paragraphs[i].Range.Delete();
                }
            }
        }

        public static void GoToKeyword(Word.Document doc, string keyword)
        {
            Word.Range range = doc.Content;
            Word.Find find = range.Find;

            find.ClearFormatting();
            find.Text = keyword;

            if (find.Execute())
            {
                Word.Range foundRange = range.Duplicate;
                foundRange.Select();  // 커서 이동 및 스크롤 이동
            }
            else
            {
                MessageBox.Show($"\"{keyword}\" 텍스트를 찾을 수 없습니다.");
            }
        }

    }
}
