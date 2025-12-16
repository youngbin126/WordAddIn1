using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Word = Microsoft.Office.Interop.Word;

namespace WordAddIn1   // 프로젝트 네임스페이스와 맞추세요
{
    using System.Windows.Forms;
    using Word = Microsoft.Office.Interop.Word;

    namespace WordAddIn1   // 프로젝트 네임스페이스에 맞게 수정
    {
        internal static class Formater
        {
            // 단어 잘림(하이픈) 허용 여부 기억용 (처음 기본값: 허용)
            private static bool _allowWordBreaking = true;

            internal static void ApplyPatentFormat(Word.Application app, Word.Document doc)
            {
                if (app == null || doc == null)
                    return;

                // 1. 문서 여백 설정 (cm → point)
                Word.PageSetup ps = doc.PageSetup;
                ps.TopMargin = app.CentimetersToPoints(4.0f);   // 위 4cm
                ps.BottomMargin = app.CentimetersToPoints(2.0f);   // 아래 2cm
                ps.LeftMargin = app.CentimetersToPoints(2.5f);   // 왼쪽 2.5cm
                ps.RightMargin = app.CentimetersToPoints(2.0f);   // 오른쪽 2cm

                // 2. 전체 문서 글꼴/크기/스타일 설정
                Word.Range allRange = doc.Content;
                allRange.Font.NameFarEast = "바탕";    // 한글
                allRange.Font.Name = "Batang";  // 영문 (환경에 따라 "바탕"으로 조정)
                allRange.Font.Size = 12;
                allRange.Font.Bold = 0;
                allRange.Font.Italic = 0;

                // 3. 단어 잘림(하이픈) 허용 여부 질문
                MessageBoxDefaultButton defaultBtn =
                    _allowWordBreaking ? MessageBoxDefaultButton.Button1 : MessageBoxDefaultButton.Button2;

                DialogResult result = MessageBox.Show(
                    "단어 잘림(하이픈)을 허용하시겠습니까?\r\n" +
                    "Yes: 단어 잘림 허용\r\n" +
                    "No: 단어 잘림 허용 안 함",
                    "단어 잘림 설정",
                    MessageBoxButtons.YesNo,
                    MessageBoxIcon.Question,
                    defaultBtn);

                _allowWordBreaking = (result == DialogResult.Yes);

                // 4. 문서 전체에 공통 단락 서식 일괄 적용
                Word.ParagraphFormat pf = allRange.ParagraphFormat;

                // 첫줄/마지막줄 분리 방지 해제
                pf.WidowControl = 0;  // false

                // 줄간격 2.14배
                pf.LineSpacingRule = Word.WdLineSpacing.wdLineSpaceMultiple;
                pf.LineSpacing = app.LinesToPoints(2.14f);

                // 첫줄 들여쓰기 1.41cm
                pf.FirstLineIndent = app.CentimetersToPoints(1.41f);

                // 단어 잘림 허용/비허용 (0: false, -1: true)
                pf.Hyphenation = _allowWordBreaking ? -1 : 0;

                // 전자 문장 부호/간격 옵션 해제 (0: false)
                pf.HalfWidthPunctuationOnTopOfLine = 0;   // 줄 첫머리 전자 문장 부호 반자로: 해제
                pf.AddSpaceBetweenFarEastAndAlpha = 0;   // 한글-영어 간격 자동 조절: 해제
                pf.AddSpaceBetweenFarEastAndDigit = 0;   // 한글-숫자 간격 자동 조절: 해제

                // 같은 스타일 단락 사이 공백 없애기
                pf.SpaceBefore = 0f;
                pf.SpaceAfter = 0f;

                // [페이지 설정] 문자 수에 맞춰 문자 간격 조정, 같은 스타일 단락 사이 공백 삽입 안 함
                // → Interop에 속성이 직접 노출되지 않는 경우가 있어 dynamic으로 처리
                try
                {
                    dynamic dpf = pf;
                    // [페이지 설정]에서 지정된 문자 수에 맞춰 문자 간격 조정(W) 해제
                    dpf.SnapToGrid = 0;   // 또는 false

                    // 같은 스타일의 단락 사이에 공백 삽입 안 함(C) 체크
                    dpf.NoSpaceBetweenParagraphsOfSameStyle = -1;   // 또는 true
                }
                catch
                {
                    // 해당 속성이 없는 구버전/다른 PIA에서는 무시
                }

                // 5. 【로 시작하는 단락에 대한 예외 처리
                foreach (Word.Paragraph para in doc.Paragraphs)
                {
                    Word.Range pr = para.Range;
                    Word.ParagraphFormat fmt = para.Format;

                    string text = pr.Text;
                    if (string.IsNullOrWhiteSpace(text))
                        continue;

                    string heading = text.Trim('\r', '\n', ' ', '\t');

                    if (!heading.StartsWith("【"))
                        continue;

                    // 【로 시작하는 모든 단락: 들여쓰기 없음 + 개요 수준 2
                    fmt.FirstLineIndent = 0f;
                    fmt.OutlineLevel = Word.WdOutlineLevel.wdOutlineLevel2;

                    // 아래 4개는 개요 수준 1
                    if (heading == "【발명의 설명】" ||
                        heading == "【청구범위】" ||
                        heading == "【요약서】" ||
                        heading == "【도면】")
                    {
                        fmt.OutlineLevel = Word.WdOutlineLevel.wdOutlineLevel1;
                    }
                }
            }
        }
    }

}
