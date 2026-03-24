using System;
using System.Collections.Generic;
using System.Text;
using System.Windows.Forms;
using Word = Microsoft.Office.Interop.Word;


namespace WordAddIn1
{
    internal sealed class EquationTransformOptions
    {
        public bool IncludeHeaderFooterStories { get; set; }
        public bool IncludeCommentStory { get; set; }
    }

    internal sealed class EquationAnchorInfo
    {
        public Word.WdStoryType StoryType { get; set; }
        public int Start { get; set; }
        public int End { get; set; }
        public int ParagraphNumber { get; set; }
        public bool IsInsideTableCell { get; set; }
    }

    public static class PCTTransformer
    {
        public static void PCTTransform(Word.Document doc)
        {
            // 추적 변경 기능 켜기
            doc.TrackRevisions = true;

            Word.Range docRange = doc.Content;

            // 텍스트 치환
            TextChanger.ReplaceText(docRange, "【발명의 배경이 되는 기술】", "【배경기술】");
            TextChanger.ReplaceText(docRange, "【과제의 해결 수단】", "【기술적 해결방법】");
            TextChanger.ReplaceText(docRange, "【발명을 실시하기 위한 구체적인 내용】", "【발명의 실시를 위한 형태】");
            TextChanger.ReplaceText(docRange, "【특허청구범위】", "【청구의 범위】");
            TextChanger.ReplaceText(docRange, "【청구범위】", "【청구의 범위】");

            // 특정 문단 삭제
            TextChanger.DeleteLinesContainingText(doc, "【요약】", deleteNextLine: false);
            TextChanger.DeleteLinesContainingText(doc, "【대표도】", deleteNextLine: true);

            // 추적 변경 기능 끔
            //            doc.TrackRevisions = false;

            // 【청구의 범위】로 이동
            TextChanger.GoToKeyword(doc, "【청구의 범위】");

            // 수식(OMaths) 치환 옵션 선택 후 실행
            EquationTransformOptions options = AskEquationTransformOptions();
            TransformEquationsToInlineShapes(doc, options);

        }

        private static EquationTransformOptions AskEquationTransformOptions()
        {
            var includeHeaderFooter = MessageBox.Show(
                "수식 치환 시 머리말/바닥말 스토리도 포함할까요?",
                "수식 치환 옵션",
                MessageBoxButtons.YesNo,
                MessageBoxIcon.Question) == DialogResult.Yes;

            var includeComments = MessageBox.Show(
                "수식 치환 시 주석(댓글) 스토리도 포함할까요?",
                "수식 치환 옵션",
                MessageBoxButtons.YesNo,
                MessageBoxIcon.Question) == DialogResult.Yes;

            return new EquationTransformOptions
            {
                IncludeHeaderFooterStories = includeHeaderFooter,
                IncludeCommentStory = includeComments
            };
        }

        private static void TransformEquationsToInlineShapes(Word.Document doc, EquationTransformOptions options)
        {
            var logLines = new List<string>();
            var failures = new List<string>();
            int convertedCount = 0;
            int skippedStoryCount = 0;

            foreach (Word.Range storyRange in doc.StoryRanges)
            {
                Word.Range currentStory = storyRange;
                while (currentStory != null)
                {
                    var nextStory = currentStory.NextStoryRange;
                    try
                    {
                        if (!ShouldProcessStory(currentStory.StoryType, options))
                        {
                            skippedStoryCount++;
                            continue;
                        }

                        convertedCount += ConvertStoryOMaths(currentStory, logLines, failures);
                    }
                    finally
                    {
                        currentStory = nextStory;
                    }
                }
            }

            ShowEquationTransformSummary(convertedCount, skippedStoryCount, logLines, failures);
        }

        private static bool ShouldProcessStory(Word.WdStoryType storyType, EquationTransformOptions options)
        {
            if (storyType == Word.WdStoryType.wdMainTextStory || storyType == Word.WdStoryType.wdTextFrameStory)
            {
                return true;
            }

            if (storyType == Word.WdStoryType.wdCommentsStory)
            {
                return options.IncludeCommentStory;
            }

            if (IsHeaderOrFooterStory(storyType))
            {
                return options.IncludeHeaderFooterStories;
            }

            return false;
        }

        private static bool IsHeaderOrFooterStory(Word.WdStoryType storyType)
        {
            switch (storyType)
            {
                case Word.WdStoryType.wdPrimaryHeaderStory:
                case Word.WdStoryType.wdEvenPagesHeaderStory:
                case Word.WdStoryType.wdFirstPageHeaderStory:
                case Word.WdStoryType.wdPrimaryFooterStory:
                case Word.WdStoryType.wdEvenPagesFooterStory:
                case Word.WdStoryType.wdFirstPageFooterStory:
                    return true;
                default:
                    return false;
            }
        }

        private static int ConvertStoryOMaths(Word.Range storyRange, List<string> logLines, List<string> failures)
        {
            int converted = 0;
            // 뒤에서 앞으로 순회해야 치환 시 인덱스가 안정적임
            for (int i = storyRange.OMaths.Count; i >= 1; i--)
            {
                Word.OMath equation = null;
                try
                {
                    equation = storyRange.OMaths[i];
                    Word.Range eqRange = equation.Range;

                    var anchorInfo = BuildAnchorInfo(storyRange, eqRange);
                    logLines.Add(
                        $"[BEFORE] StoryType={anchorInfo.StoryType}, Start={anchorInfo.Start}, End={anchorInfo.End}, Paragraph={anchorInfo.ParagraphNumber}, InCell={anchorInfo.IsInsideTableCell}");

                    Word.InlineShape insertedShape;
                    if (anchorInfo.IsInsideTableCell)
                    {
                        // 셀 내부 수식은 셀 범위 내에서만 치환
                        insertedShape = ReplaceEquationInsideCell(eqRange);
                    }
                    else
                    {
                        insertedShape = ReplaceEquationRange(eqRange);
                    }

                    if (insertedShape == null || insertedShape.Range == null)
                    {
                        failures.Add(
                            $"치환 실패: StoryType={anchorInfo.StoryType}, Start={anchorInfo.Start}, End={anchorInfo.End} (삽입 InlineShape 없음)");
                        continue;
                    }

                    if (insertedShape.Range.StoryType != anchorInfo.StoryType)
                    {
                        failures.Add(
                            $"치환 실패: StoryType 불일치 (before={anchorInfo.StoryType}, after={insertedShape.Range.StoryType}, Start={anchorInfo.Start}, End={anchorInfo.End})");
                        continue;
                    }

                    converted++;
                }
                catch (Exception ex)
                {
                    failures.Add($"치환 실패: StoryType={storyRange.StoryType}, Index={i}, Error={ex.Message}");
                }
            }

            return converted;
        }

        private static EquationAnchorInfo BuildAnchorInfo(Word.Range storyRange, Word.Range equationRange)
        {
            return new EquationAnchorInfo
            {
                StoryType = equationRange.StoryType,
                Start = equationRange.Start,
                End = equationRange.End,
                ParagraphNumber = GetParagraphNumberInStory(storyRange, equationRange),
                IsInsideTableCell = IsInsideTableCell(equationRange)
            };
        }

        private static int GetParagraphNumberInStory(Word.Range storyRange, Word.Range targetRange)
        {
            Word.Range countingRange = storyRange.Duplicate;
            countingRange.End = targetRange.Start;
            return countingRange.Paragraphs.Count + 1;
        }

        private static bool IsInsideTableCell(Word.Range range)
        {
            try
            {
                return range.Cells != null && range.Cells.Count > 0;
            }
            catch
            {
                return false;
            }
        }

        private static Word.InlineShape ReplaceEquationInsideCell(Word.Range equationRange)
        {
            Word.Range cellRange = equationRange.Cells[1].Range.Duplicate;
            int start = equationRange.Start;
            int end = equationRange.End;

            if (start < cellRange.Start)
            {
                start = cellRange.Start;
            }

            if (end > cellRange.End)
            {
                end = cellRange.End;
            }

            Word.Range boundedRange = equationRange.Document.Range(start, end);
            return ReplaceEquationRange(boundedRange);
        }

        private static Word.InlineShape ReplaceEquationRange(Word.Range equationRange)
        {
            equationRange.CopyAsPicture();
            equationRange.Text = string.Empty;
            equationRange.PasteSpecial(DataType: Word.WdPasteDataType.wdPasteEnhancedMetafile);
            return equationRange.InlineShapes.Count > 0
                ? equationRange.InlineShapes[1]
                : null;
        }

        private static void ShowEquationTransformSummary(
            int convertedCount,
            int skippedStoryCount,
            List<string> logs,
            List<string> failures)
        {
            var summary = new StringBuilder();
            summary.AppendLine("수식(OMaths) 치환 요약");
            summary.AppendLine($"- 성공: {convertedCount}");
            summary.AppendLine($"- 제외된 스토리 수: {skippedStoryCount}");
            summary.AppendLine($"- 실패: {failures.Count}");
            summary.AppendLine();
            summary.AppendLine("[앵커 로그]");
            foreach (var line in logs)
            {
                summary.AppendLine(line);
            }

            if (failures.Count > 0)
            {
                summary.AppendLine();
                summary.AppendLine("[실패 목록]");
                foreach (var failure in failures)
                {
                    summary.AppendLine($"- {failure}");
                }
            }

            MessageBox.Show(summary.ToString(), "수식 치환 결과", MessageBoxButtons.OK, MessageBoxIcon.Information);
        }
    }
}
