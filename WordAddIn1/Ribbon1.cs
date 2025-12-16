using Microsoft.Office.Interop.Word;
using Microsoft.Office.Tools.Ribbon;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using WordAddIn1.WordAddIn1;
using Word = Microsoft.Office.Interop.Word;

namespace WordAddIn1
{
    public partial class Ribbon1
    {
        private void Ribbon1_Load(object sender, RibbonUIEventArgs e)
        {

        }
        private void TableComparer_Click(object sender, RibbonControlEventArgs e)
        {
            TableComparer comparer = new TableComparer(Globals.ThisAddIn.Application);
            comparer.CompareEachRow();
        }

        private void ImageInserter_Click(object sender, RibbonControlEventArgs e)
        {
            Word.Document doc = Globals.ThisAddIn.Application.ActiveDocument;
            Word.Selection sel = Globals.ThisAddIn.Application.Selection;

            ImageInserter inserter = new ImageInserter();
            inserter.InsertImages(doc, sel);
        }

private void AuthorChanger_Click(object sender, RibbonControlEventArgs e)
{
    try
    {
        var app = Globals.ThisAddIn.Application;
        string currentAuthor = app.UserName;

        using (Form inputForm = new Form())
        {
            inputForm.Text = "작성자 수정";
            inputForm.AutoScaleMode = AutoScaleMode.Dpi; // DPI에 따라 자동 스케일링
            inputForm.StartPosition = FormStartPosition.CenterParent;
            inputForm.FormBorderStyle = FormBorderStyle.FixedDialog;
            inputForm.Width = 400;
            inputForm.Height = 220;
            inputForm.MaximizeBox = false;
            inputForm.MinimizeBox = false;
            inputForm.ShowInTaskbar = false;

            TextBox textBox = new TextBox
            {
                Text = currentAuthor,
                Dock = DockStyle.Top,
                Margin = new Padding(10),
            };
            textBox.SelectionStart = textBox.Text.Length; // 커서를 끝으로
            textBox.SelectionLength = 0;

            Button okButton = new Button
            {
                Text = "확인",
                DialogResult = DialogResult.OK,
                Height = 35, // 원하는 높이 지정
                Width = 80,  // 폭도 직접 지정 가능
                Top = 60,    // 폼 안에서 위치 지정
                Left = (inputForm.ClientSize.Width - 80) / 2, // 가운데 정렬
                Anchor = AnchorStyles.Top  // 위치 고정
            };

                    inputForm.Controls.Add(textBox);
            inputForm.Controls.Add(okButton);
            inputForm.AcceptButton = okButton;

            if (inputForm.ShowDialog() == DialogResult.OK)
            {
                string newAuthor = textBox.Text.Trim();
                if (!string.IsNullOrEmpty(newAuthor))
                {
                    app.UserName = newAuthor;
                }
            }
        }
    }
    catch (Exception ex)
    {
        MessageBox.Show("작성자 변경 중 오류 발생: " + ex.Message);
    }
}


        private void TextSpliter_Click(object sender, RibbonControlEventArgs e)
        {
            try
            {
                // 현재 활성 문서 가져오기
                Word.Document doc = Globals.ThisAddIn.Application.ActiveDocument;
                Word.Selection selection = doc.Application.Selection;

                // 선택된 텍스트가 있는지 확인
                if (selection.Type == Word.WdSelectionType.wdSelectionNormal)
                {
                    // 선택된 범위의 문단 가져오기
                    Word.Range selectionRange = selection.Range;
                    Word.Paragraphs paragraphs = selectionRange.Paragraphs;

                    // 문단이 있는 경우에만 처리
                    if (paragraphs.Count > 0)
                    {
                        // 문단의 개수로 표 생성
                        CreateTable(doc, selectionRange, paragraphs);
                    }
                    else
                    {
                        System.Windows.Forms.MessageBox.Show("선택된 문단이 없습니다.", "알림",
                            System.Windows.Forms.MessageBoxButtons.OK,
                            System.Windows.Forms.MessageBoxIcon.Information);
                    }
                }
                else
                {
                    System.Windows.Forms.MessageBox.Show("텍스트를 선택해 주세요.", "알림",
                        System.Windows.Forms.MessageBoxButtons.OK,
                        System.Windows.Forms.MessageBoxIcon.Information);
                }
            }
            catch (Exception ex)
            {
                System.Windows.Forms.MessageBox.Show($"오류가 발생했습니다: {ex.Message}", "오류",
                    System.Windows.Forms.MessageBoxButtons.OK,
                    System.Windows.Forms.MessageBoxIcon.Error);
            }
        }

        private void CreateTable(Word.Document doc, Word.Range selectionRange, Word.Paragraphs paragraphs)
        {
            Word.Table table = null;
            Word.Range tableRange = null;
            try
            {
                // 유효한 문단만 필터링
                List<Word.Range> validParagraphRanges = new List<Word.Range>();
                foreach (Word.Paragraph para in paragraphs)
                {
                    string paraText = para.Range.Text?.Trim();
                    if (!string.IsNullOrEmpty(paraText) && paraText != "\r")
                    {
                        validParagraphRanges.Add(para.Range);
                    }
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(para);
                }

                // 유효한 문단이 없는 경우
                if (validParagraphRanges.Count == 0)
                {
                    System.Windows.Forms.MessageBox.Show("유효한 텍스트가 포함된 문단이 없습니다.", "알림",
                        System.Windows.Forms.MessageBoxButtons.OK,
                        System.Windows.Forms.MessageBoxIcon.Information);
                    return;
                }

                // 선택된 범위의 끝으로 이동하여 표 삽입
                tableRange = selectionRange.Duplicate;
                tableRange.Collapse(Word.WdCollapseDirection.wdCollapseEnd);

                // 선택된 위치 이후에 표 삽입 (유효한 문단 수로 표 생성)
                table = doc.Tables.Add(tableRange, validParagraphRanges.Count, 2,
                    Word.WdDefaultTableBehavior.wdWord9TableBehavior,
                    Word.WdAutoFitBehavior.wdAutoFitContent);



                // 유효한 문단을 표 셀에 삽입
                for (int i = 0; i < validParagraphRanges.Count; i++)
                {
                    Word.Range paragraphRange = validParagraphRanges[i];
                    Word.Cell cell1 = table.Cell(i + 1, 1);
                    cell1.Range.FormattedText = paragraphRange.FormattedText; // 문단 콘텐츠 복사
                }

                // 열 너비를 균등하게 분배
                table.Columns.DistributeWidth();
            }
            catch (Exception ex)
            {
                System.Windows.Forms.MessageBox.Show($"표 생성 중 오류가 발생했습니다: {ex.Message}\nStackTrace: {ex.StackTrace}", "오류",
                    System.Windows.Forms.MessageBoxButtons.OK,
                    System.Windows.Forms.MessageBoxIcon.Error);
            }
            finally

            {
                // COM 객체 해제
                if (paragraphs != null)
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(paragraphs);
                if (tableRange != null)
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(tableRange);
                if (table != null)
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(table);
            }
        }

        private void PrepareSubmissionFile_Click(object sender, RibbonControlEventArgs e)
        {
            var doc = Globals.ThisAddIn.Application.ActiveDocument;
        }

        private void CommentRemover_Click(object sender, RibbonControlEventArgs e)
        {
            var selection = Globals.ThisAddIn.Application.Selection;
            CommentRemover.RemoveCommentsInSelection(selection);
        }


        private void DrawingReferenceRemover_Click(object sender, RibbonControlEventArgs e)
        {
            DrawingReferenceRemover.RemoveDrawingReferences();
        }

        private void DrawingValidatorClick(object sender, RibbonControlEventArgs e)
        {
            var validator = new DrawingValidator();
            validator.ValidateDrawings(Globals.ThisAddIn.Application.ActiveDocument);
        }

        private void PCTTransformer_Click(object sender, RibbonControlEventArgs e)
        {
          
            var doc = Globals.ThisAddIn.Application.ActiveDocument;
            PCTTransformer.PCTTransform(doc);
        }

        private void ClaimParaphraser_Click(object sender, RibbonControlEventArgs e)
        {
            var generator = new ClaimTableGenerator();
            generator.GenerateClaimTableAndAskTrackChange();
        }

        private void btnGenerateClaimTable_Click(object sender, RibbonControlEventArgs e)
        
            {
                ClaimTableGenerator generator = new ClaimTableGenerator();
                generator.GenerateClaimTableFromSelection();
            }

        private void DocumentComparer_Click(object sender, RibbonControlEventArgs e)
        {
                DocumentComparer.CompareActiveDocument();
        }

        private void Formater_Click(object sender, RibbonControlEventArgs e)
        {
           

            }

        private void btnApplyPatentFormat_Click(object sender, RibbonControlEventArgs e)
        {
            Word.Application app = Globals.ThisAddIn.Application;
            Word.Document doc = app.ActiveDocument;

            Formater.ApplyPatentFormat(app, doc);

        }

        private void button14_Click(object sender, RibbonControlEventArgs e)
        {
            Word.Application app = Globals.ThisAddIn.Application;
            Word.Selection sel = app.Selection;
            if (sel == null) return;

            Word.Range rng = sel.Range;
            Word.Revisions revisions = rng.Revisions;

            // COM 컬렉션은 1-based, Accept() 하면 컬렉션이 줄어들므로 뒤에서 앞으로 순회
            for (int i = revisions.Count; i >= 1; i--)
            {
                Word.Revision rev = revisions[i];

                switch (rev.Type)
                {
                    case Word.WdRevisionType.wdRevisionProperty:
                    case Word.WdRevisionType.wdRevisionStyle:
                    case Word.WdRevisionType.wdRevisionStyleDefinition:
                    case Word.WdRevisionType.wdRevisionParagraphNumber:
                    case Word.WdRevisionType.wdRevisionCellMerge:
                    case Word.WdRevisionType.wdRevisionCellSplit:
                    case Word.WdRevisionType.wdRevisionTableProperty:
                        rev.Accept();
                        break;
                    default:
                        // 삽입/삭제 등 내용 변경은 그대로 둠
                        break;
                }
            }
        }

    }
}


