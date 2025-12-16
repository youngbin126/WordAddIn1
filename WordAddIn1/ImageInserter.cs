using System.IO;
using System.Windows.Forms;
using Word = Microsoft.Office.Interop.Word;

namespace WordAddIn1
{
    public class ImageInserter
    {
        public void InsertImages(Word.Document doc, Word.Selection sel)
        {
/*           bool hasSubTitleStyle = StyleExists(doc, "소제목");
            bool hasNormalStyle = StyleExists(doc, "표준");*/

            using (OpenFileDialog dlg = new OpenFileDialog())
            {
                dlg.Title = "이미지 파일 선택 (복수 선택 가능)";
                dlg.Filter = "모든 이미지 파일|*.tif;*.tiff;*.jpg;*.jpeg;*.png;*.bmp;*.gif|모든 파일|*.*";
                dlg.Multiselect = true;

                if (dlg.ShowDialog() == DialogResult.OK)
                {
                    foreach (string filePath in dlg.FileNames)
                    {
                        string fileName = Path.GetFileNameWithoutExtension(filePath).ToLower();


                        // 캡션 삽입
                        Word.Range captionRange = sel.Range;
                        captionRange.Text = $"【도 {fileName}】\r\n";

                        // 캡션 단락이 다음 단락(이미지)과 같은 페이지에 있도록 설정
                        captionRange.ParagraphFormat.KeepWithNext = -1;

                        /*                        if (hasSubTitleStyle)
                                                {
                                                    captionRange.set_Style("소제목");
                                                }*/
                        captionRange.Collapse(Word.WdCollapseDirection.wdCollapseEnd);
                        captionRange.Select();

                        // 이미지 삽입
                        Word.InlineShape img = sel.InlineShapes.AddPicture(filePath, false, true);

                        // 이미지 단락이 다음 단락(캡션)과 같은 페이지에 있지 않아도 되도록 설정
                        captionRange.ParagraphFormat.KeepWithNext = 0;

                        // 이미지 단락의 줄 간격을 1줄로 설정 (고정 크기의 단락이 설정되면 이미지 짤림)
                        captionRange.ParagraphFormat.LineSpacingRule = Word.WdLineSpacing.wdLineSpaceSingle;

                        sel.Collapse(Word.WdCollapseDirection.wdCollapseEnd);

/*                        // 이미지 스타일 적용
                        if (hasNormalStyle)
                        {
                            sel.MoveLeft(Word.WdUnits.wdCharacter, 1);
                            sel.set_Style("표준");
                            sel.MoveRight(Word.WdUnits.wdCharacter, 1);
                        }*/

                        // 한 줄 띄우기
                        sel.TypeParagraph();
                    }
                }
            }
        }

        private bool StyleExists(Word.Document doc, string styleName)
        {
            foreach (Word.Style style in doc.Styles)
            {
                if (style.NameLocal == styleName)
                    return true;
            }
            return false;
        }
    }
}
