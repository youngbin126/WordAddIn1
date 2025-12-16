using Word = Microsoft.Office.Interop.Word;
using System.Windows.Forms;
using System.Text.RegularExpressions;
using System.Collections.Generic;
using System.Linq;

namespace WordAddIn1
{
    public partial class DrawingValidator
    {
        // 리본 메뉴 버튼 클릭 시 호출
        public void ValidateDrawings(Word.Document doc)
        {
            // 1~5. 도면 삽입 확인 및 테이블 생성
            CheckDrawingInsertion(doc);
        }

        // 1~5. 도면 삽입 확인 및 테이블 생성
        private void CheckDrawingInsertion(Word.Document doc)
        {
            // 1. 도면 설명 영역 식별
            Word.Range drawingSection = IdentifyDrawingSection(doc);
            if (drawingSection == null)
            {
                MessageBox.Show("도면 설명 섹션을 찾을 수 없습니다.");
                return;
            }

            // 2~3. 도면 문장 및 번호 식별
            List<string> drawingNumbers = FindDrawingNumbers(drawingSection);
            if (drawingNumbers.Count == 0)
            {
                MessageBox.Show("도면 번호 문장을 찾을 수 없습니다.");
                return;
            }

            // 4~5. 도면 태그 및 이미지 확인, 테이블 생성
            var drawingData = CheckDrawingImages(doc, drawingNumbers);
            if (drawingData.Any(d => d.HasError))
            {
                MessageBox.Show("도면 또는 캡션에 문제가 있습니다. 자세한 내용은 테이블을 확인하세요.");
            }
            else
            {
                MessageBox.Show("모든 도면이 정상적으로 삽입되었습니다.");
            }

            // 별도 창에 테이블 생성
            CreateDrawingTableInNewDocument(doc, drawingData);
        }

        // 1. 도면 설명 영역 식별
        private Word.Range IdentifyDrawingSection(Word.Document doc)
        {
            Word.Range drawingSection = null;
            bool foundStart = false;

            foreach (Word.Paragraph para in doc.Paragraphs)
            {
                string text = para.Range.Text.Trim();
                if (!foundStart && text.Contains("【도면의 간단한 설명】"))
                {
                    foundStart = true;
                    drawingSection = para.Range.Document.Range(para.Range.End);
                    continue;
                }
                if (foundStart && text.StartsWith("【"))
                {
                    drawingSection.End = para.Range.Start - 1;
                    break;
                }
            }

            return drawingSection;
        }

        // 2~3. 도면 문장 및 번호 식별
        private List<string> FindDrawingNumbers(Word.Range section)
        {
            var numbers = new List<string>();
            string pattern = @"도\s?\d+[a-zA-Z]?\s*(?:는|은)"; // 숫자와 "은/는" 사이 공백 처리

            foreach (Word.Paragraph para in section.Paragraphs)
            {
                // 텍스트 정규화: 제어 문자 및 다중 공백 제거
                string text = para.Range.Text;
                text = Regex.Replace(text, @"[\r\n\t]+", " "); // 제어 문자 제거
                text = text.Trim(); // 양쪽 공백 제거
                text = Regex.Replace(text, @"\s+", " "); // 다중 공백을 단일 공백으로

                if (Regex.IsMatch(text, pattern))
                {
                    var matches = Regex.Matches(text, @"도\s?(\d+[a-zA-Z]?)\s*(?:는|은)");
                    foreach (Match match in matches)
                    {
                        if (match.Success)
                        {
                            numbers.Add(match.Groups[1].Value); // 예: "1", "3a", "5B"
                        }
                    }
                }
                else
                {
                    // 디버깅: 매칭 실패 시 텍스트 출력 (배포 시 제거 가능)
                    if (!string.IsNullOrWhiteSpace(text))
                    {
                        // MessageBox.Show($"매칭 실패: '{text}'", "디버깅");
                    }
                }
            }

            return numbers;
        }

        // 4~5. 도면 태그 및 이미지 확인
        private List<DrawingData> CheckDrawingImages(Word.Document doc, List<string> drawingNumbers)
        {
            var drawingData = new List<DrawingData>();

            foreach (string number in drawingNumbers)
            {
                string searchText = $"【도 {number}】";
                Word.Range range = doc.Content;
                bool found = false;
                Word.InlineShape image = null;
                bool hasError = false;
                string errorMessage = null;

                // 모든 캡션 인스턴스 찾기
                while (range.Find.Execute(FindText: searchText))
                {
                    found = true;
                    Word.Paragraph nextPara = range.Paragraphs[1].Next();
                    if (nextPara == null || nextPara.Range.InlineShapes.Count == 0)
                    {
                        hasError = true;
                        errorMessage = $"도면 누락: {searchText}";
                        MessageBox.Show(errorMessage);
                    }
                    else
                    {
                        foreach (Word.InlineShape shape in nextPara.Range.InlineShapes)
                        {
                            if (shape.Type == Word.WdInlineShapeType.wdInlineShapePicture)
                            {
                                image = shape; // 첫 번째 이미지 사용
                                break;
                            }
                        }
                        if (image == null)
                        {
                            hasError = true;
                            errorMessage = $"도면 누락: {searchText}";
                            MessageBox.Show(errorMessage);
                        }
                    }
                    // 다음 검색을 위해 범위 이동
                    range = doc.Range(range.End, doc.Content.End);
                }

                if (!found)
                {
                    hasError = true;
                    errorMessage = $"캡션 누락: {searchText}";
                    MessageBox.Show(errorMessage);
                }

                drawingData.Add(new DrawingData
                {
                    Number = number,
                    Image = image,
                    HasError = hasError,
                    ErrorMessage = errorMessage
                });
            }

            return drawingData;
        }

        // 별도 창에 테이블 생성 (저장 없이 표시)
        private void CreateDrawingTableInNewDocument(Word.Document sourceDoc, List<DrawingData> drawingData)
        {
            Word.Application wordApp = sourceDoc.Application;
            Word.Document newDoc = null;

            try
            {
                // 새 임시 문서 생성
                newDoc = wordApp.Documents.Add();
                newDoc.Saved = true; // 저장 프롬프트 방지

                // 문서 시작에 테이블 삽입
                Word.Range tableRange = newDoc.Content;
                tableRange.Collapse(Word.WdCollapseDirection.wdCollapseStart);

                // 2열 테이블 생성 (도면 번호, 이미지)
                Word.Table table = newDoc.Tables.Add(tableRange, drawingData.Count + 1, 2);

                // 헤더 설정
                table.Cell(1, 1).Range.Text = "도면 번호\r";
                table.Cell(1, 2).Range.Text = "도면 이미지\r";
                table.Cell(1, 1).Range.Bold = 1;
                table.Cell(1, 2).Range.Bold = 1;

                // 데이터 삽입
                for (int i = 0; i < drawingData.Count; i++)
                {
                    var data = drawingData[i];
                    int row = i + 2;

                    // 왼쪽 열: 도면 번호
                    table.Cell(row, 1).Range.Text = $"도 {data.Number}\r";

                    // 오른쪽 열: 이미지 또는 에러 메시지
                    if (data.Image != null && !data.HasError)
                    {
                        try
                        {
                            // 이미지 복사
                            data.Image.Select();
                            sourceDoc.Application.Selection.Copy();
                            table.Cell(row, 2).Range.Paste();
                            // 이미지 크기 조정 (선택적)
                            foreach (Word.InlineShape shape in table.Cell(row, 2).Range.InlineShapes)
                            {
                                shape.Width = 100; // 픽셀 단위, 필요 시 조정
                                shape.Height = 100;
                            }
                        }
                        catch
                        {
                            table.Cell(row, 2).Range.Text = "이미지 삽입 실패\r";
                        }
                    }
                    else
                    {
                        table.Cell(row, 2).Range.Text = data.ErrorMessage ?? "이미지 없음\r";
                    }
                }

                // 테이블 스타일링
                table.Rows.Alignment = Word.WdRowAlignment.wdAlignRowCenter;
                table.Columns[1].Width = 100; // 도면 번호 열 너비
                table.Columns[2].Width = 200; // 이미지 열 너비

                // 새 문서 창 활성화
                newDoc.Activate();
                MessageBox.Show("새 창에 도면 테이블이 표시되었습니다. 확인 후 창을 닫아주세요.", "도면 테이블");
            }
            catch (System.Exception ex)
            {
                MessageBox.Show($"새 문서 생성 중 오류 발생: {ex.Message}");
            }
            finally
            {
                // COM 객체 해제
                if (newDoc != null)
                {
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(newDoc);
                }
            }
        }

        // 도면 데이터 클래스
        private class DrawingData
        {
            public string Number { get; set; }
            public Word.InlineShape Image { get; set; }
            public bool HasError { get; set; }
            public string ErrorMessage { get; set; }
        }
    }
}