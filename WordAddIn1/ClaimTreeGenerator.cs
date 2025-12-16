using System;
using System.Collections.Generic;
using System.Text.RegularExpressions;
using System.Windows.Forms;
using static System.Net.Mime.MediaTypeNames;
using Word = Microsoft.Office.Interop.Word;

namespace WordAddIn1
{
    public class ClaimTableGenerator
    {
        // 이전 옵션 기억용 static 필드
        private static Word.WdOrientation lastOrientation = Word.WdOrientation.wdOrientLandscape;
        private static bool lastFixColumnWidth = true;
        private static bool lastShowContents = true;
        private static bool lastShowDrawings = true; // 새로 추가

        // 옵션 선택용 Form
        public class TableOptionsForm : Form
        {
            public Word.WdOrientation SelectedOrientation { get; private set; }
            public bool FixColumnWidth { get; private set; }
            public bool ShowContents { get; private set; }
            public bool ShowDrawings { get; private set; } // 새로 추가

            private RadioButton rbPortrait;
            private RadioButton rbLandscape;
            private CheckBox cbFixWidth;
            private CheckBox cbShowContents;
            private CheckBox cbShowDrawings;
            private Button btnOK;

            public TableOptionsForm()
            {
                Text = "테이블 옵션 선택";
                Width = 320;
                Height = 260;
                FormBorderStyle = FormBorderStyle.FixedDialog;
                StartPosition = FormStartPosition.CenterParent;

                rbPortrait = new RadioButton() { Text = "세로 문서", Left = 20, Top = 20, Checked = (lastOrientation == Word.WdOrientation.wdOrientPortrait) };
                rbLandscape = new RadioButton() { Text = "가로 문서", Left = 160, Top = 20, Checked = (lastOrientation == Word.WdOrientation.wdOrientLandscape) };

                cbFixWidth = new CheckBox() { Text = "너비 고정", Left = 20, Top = 80, Checked = lastFixColumnWidth };
                cbShowContents = new CheckBox() { Text = "내용 표시", Left = 20, Top = 110, Checked = lastShowContents };
                cbShowDrawings = new CheckBox() { Text = "도면칸", Left = 20, Top = 140, Checked = lastShowDrawings };

                btnOK = new Button() { Text = "확인", Left = 120, Top = 180, Width = 80 };
                btnOK.Click += (s, e) =>
                {
                    SelectedOrientation = rbLandscape.Checked ? Word.WdOrientation.wdOrientLandscape : Word.WdOrientation.wdOrientPortrait;
                    FixColumnWidth = cbFixWidth.Checked;
                    ShowContents = cbShowContents.Checked;
                    ShowDrawings = cbShowDrawings.Checked;

                    // 선택값 기억
                    lastOrientation = SelectedOrientation;
                    lastFixColumnWidth = FixColumnWidth;
                    lastShowContents = ShowContents;
                    lastShowDrawings = ShowDrawings;

                    DialogResult = DialogResult.OK;
                    Close();
                };

                Controls.Add(rbPortrait);
                Controls.Add(rbLandscape);
                Controls.Add(cbFixWidth);
                Controls.Add(cbShowContents);
                Controls.Add(cbShowDrawings);
                Controls.Add(btnOK);
            }
        }


        public void GenerateClaimTableFromSelection()
        {
            var app = Globals.ThisAddIn.Application;
            var sel = app?.Selection;

            if (sel == null || sel.Range == null || string.IsNullOrWhiteSpace(sel.Range.Text ?? ""))
            {
                MessageBox.Show("청구항을 포함하는 영역을 선택해주세요.", "정보");
                return;
            }
            // 0. 옵션 선택
            TableOptionsForm optionsForm = new TableOptionsForm();
            if (optionsForm.ShowDialog() != DialogResult.OK)
                return; // 사용자가 취소한 경우

            Word.WdOrientation orientation = optionsForm.SelectedOrientation;
            bool fixColumnWidth = optionsForm.FixColumnWidth;
            bool showContents = optionsForm.ShowContents;
            bool showDrawings = optionsForm.ShowDrawings;


            string selectedText = sel.Range.Text;

            // 1. 【청구항 #】 단위로 나누기
            Regex claimSplitRegex = new Regex(@"【청구항\s*(\d+)】");
            MatchCollection matches = claimSplitRegex.Matches(selectedText);

            if (matches.Count == 0)
            {
                MessageBox.Show("선택 영역 내에서 【청구항 #】을 찾을 수 없습니다.", "정보");
                return;
            }

            // 2. Claim 객체 생성
            Dictionary<int, Claim> claimDict = new Dictionary<int, Claim>();
            List<int> claimNumbers = new List<int>();
            for (int i = 0; i < matches.Count; i++)
            {
                int startIndex = matches[i].Index;
                int endIndex = (i + 1 < matches.Count) ? matches[i + 1].Index : selectedText.Length;
                string claimText = selectedText.Substring(startIndex, endIndex - startIndex).Trim();

                int claimNumber = int.Parse(matches[i].Groups[1].Value);
                claimNumbers.Add(claimNumber);

                string category = ExtractCategory(claimText);
                List<int> directDeps = ExtractDirectDependencies(claimText);

                Claim claim = new Claim
                {
                    Number = claimNumber,
                    Category = category,
                    DirectDependencies = directDeps,
                    Text = claimText
                };

                claimDict[claimNumber] = claim;
            }

            // 3. Dependency 경로 계산 및 최대 종속항 수 확인
            int maxDependencyCount = 0;
            foreach (var kvp in claimDict)
            {
                kvp.Value.DependencyPath = BuildDependencyPathList(kvp.Value, claimDict);
                if (kvp.Value.DependencyPath.Count > maxDependencyCount)
                    maxDependencyCount = kvp.Value.DependencyPath.Count;
            }

            // 열 개수: # + Subject + Dependencies + Contents + Drawings
            int totalColumns = 2 + maxDependencyCount + 1;
            if (showDrawings) totalColumns += 1; // 도면칸 옵션 선택 시에는 +1열 추가


            // 4. 새 문서 생성
            Word.Document newDoc = app.Documents.Add();
            newDoc.PageSetup.Orientation = orientation;
            Word.Range tableRange = newDoc.Range(0, 0);
            Word.Table table = newDoc.Tables.Add(tableRange, claimDict.Count + 1, totalColumns);
            table.Borders.Enable = 1;

            int contentsCol = 3 + maxDependencyCount;
            int drawingsCol = contentsCol + 1;

            if (fixColumnWidth)
            {
                for (int k = 1; k < contentsCol; k++)
                    table.Columns[k].SetWidth(app.CentimetersToPoints(0.5f), Word.WdRulerStyle.wdAdjustNone);

                table.Columns[2].SetWidth(app.CentimetersToPoints(1.18f), Word.WdRulerStyle.wdAdjustNone);
                if (showDrawings)
                    table.Columns[drawingsCol].SetWidth(app.CentimetersToPoints(5f), Word.WdRulerStyle.wdAdjustNone);

                // 가로/세로 문서에 따라 테이블 총 너비 설정
                float tableWidthCm = (orientation == Word.WdOrientation.wdOrientLandscape) ? 26.3f : 18.5f;

                // Contents 열 폭 계산
                float contentsWidth = tableWidthCm - ((contentsCol - 2) * 0.5f + 1.18f + (showDrawings ? 5f : 0f));
                table.Columns[contentsCol].SetWidth(app.CentimetersToPoints(contentsWidth), Word.WdRulerStyle.wdAdjustNone);
            }



            // 5. 헤더
            table.Cell(1, 1).Range.Text = "#";
            table.Cell(1, 2).Range.Text = "Subj.";
            for (int i = 0; i < maxDependencyCount; i++)
                table.Cell(1, 3 + i).Range.Text = "";
            table.Cell(1, contentsCol).Range.Text = "Contents";
            if (showDrawings)
                table.Cell(1, drawingsCol).Range.Text = "Drawings";

            // "Subj." 우측 셀부터 Contents 셀까지 병합
            table.Cell(1, 3).Merge(table.Cell(1, contentsCol));
            // 헤더 서식

            table.Rows[1].Range.Bold = 1;
            table.Rows[1].Range.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter;

            // 6. 내용 삽입
            int row = 2;
            int minClaim = claimNumbers[0];
            int maxClaim = claimNumbers[claimNumbers.Count - 1];

            for (int i = 0; i < claimNumbers.Count; i++)
            {
                int claimNumber = claimNumbers[i];
                var c = claimDict[claimNumber];

                table.Cell(row, 1).Range.Text = c.Number.ToString();

                // Subject
                Word.Range subjectRange = table.Cell(row, 2).Range;
                subjectRange.Text = c.Category;

                // 마침표 없는 경우 체크
                if (!c.Category.Contains("."))
                {
                    subjectRange.End--;
                    int lastSpace = subjectRange.Text.TrimEnd().LastIndexOf(' ');
                    Word.Range lastWordRange;
                    if (lastSpace >= 0)
                    {
                        lastWordRange = subjectRange.Duplicate;
                        lastWordRange.Start = subjectRange.Start + lastSpace + 1;
                    }
                    else
                    {
                        lastWordRange = subjectRange;
                    }
                    newDoc.Comments.Add(lastWordRange, "마침표가 없음");
                }

                // 단어 수 체크 (10개 초과 시 메모)
                string[] words = c.Category.Split(new char[] { ' ' }, StringSplitOptions.RemoveEmptyEntries);
                if (words.Length > 10)
                {
                    newDoc.Comments.Add(subjectRange,
                        "청구항 종결어의 길이를 줄이거나, 종결어 앞에 쉼표를 입력하세요");
                }

                // 종속항의 Category 불일치 검사
                if (c.DirectDependencies.Count > 0)
                {
                    int parentNum = c.DirectDependencies[0];
                    if (claimDict.TryGetValue(parentNum, out Claim parentClaim))
                    {
                        string childCategory = c.Category.Trim();
                        string parentCategory = parentClaim.Category.Trim();

                        // 마침표 제거
                        var cleanChild = childCategory.Replace(".", "");
                        var cleanParent = parentCategory.Replace(".", "");

                        if (!cleanChild.Equals(cleanParent, StringComparison.OrdinalIgnoreCase))
                        {
                            newDoc.Comments.Add(subjectRange, $"인용 청구항 {parentNum} 카테고리와 불일치");
                        }
                    }

                    else
                    {
                        // 인용 청구항이 없는 경우 코멘트 추가
                        newDoc.Comments.Add(subjectRange, $"인용 청구항 {parentNum} 누락");
                    }
                }

                // Dependency
                for (int j = 0; j < maxDependencyCount; j++)
                {
                    Word.Range depRange = table.Cell(row, 3 + j).Range;
                    depRange.End--;

                    if (j < c.DependencyPath.Count)
                    {
                        depRange.Text = c.DependencyPath[j].ToString();
                        depRange.Bold = 0;

                        // 뒤 청구항 의존 체크
                        if (c.DependencyPath[j] > c.Number)
                        {
                            newDoc.Comments.Add(depRange,
                                $"뒤의 청구항 {c.DependencyPath[j]} 인용");
                        }
                    }
                    else
                    {
                        depRange.Text = "";
                        depRange.Bold = 1;
                    }
                }

                // Contents 열 채우기
                if (showContents)
                {
                    Word.Range contentRange = table.Cell(row, contentsCol).Range;
                    contentRange.Text = c.Text;
                    contentRange.ParagraphFormat.FirstLineIndent = app.CentimetersToPoints(1.4f);
                }



                // 공란 Dependency 열 + 마지막 Contents 열 병합
                int mergeStartCol = 3 + c.DependencyPath.Count;

                // Drawings 열 제외
                if (mergeStartCol < contentsCol)
                {
                    Word.Cell startCell = table.Cell(row, mergeStartCol);
                    Word.Cell endCell = table.Cell(row, contentsCol);
                    startCell.Merge(endCell);
                }


                row++;
            }

            // 7. 누락 청구항 메모: 다음 존재하는 청구항 행의 # 셀에 메모 삽입
            for (int i = minClaim; i <= maxClaim; i++)
            {
                if (!claimDict.ContainsKey(i))
                {
                    // 다음 존재하는 청구항 찾기
                    int nextClaimRowIndex = 2 + claimNumbers.FindIndex(x => x > i);
                    if (nextClaimRowIndex >= 2)
                    {
                        Word.Range cellRange = table.Cell(nextClaimRowIndex, 1).Range;
                        newDoc.Comments.Add(cellRange, $"앞의 청구항 누락");
                    }
                }
            }

            // 너비 고정 비선택시
            if (!fixColumnWidth)
            {
                table.AutoFitBehavior(Word.WdAutoFitBehavior.wdAutoFitContent); // 내용에 맞춤
            }

        }



        // ----- Helper Methods -----
        private string ExtractCategory(string claimText)
        {
            // 1. 마지막 마침표 기준으로 문장 자르기
            int lastDot = claimText.LastIndexOf('.');
            if (lastDot < 0) lastDot = claimText.Length - 1;
            string sub = claimText.Substring(0, lastDot + 1).Trim();

            // 2. 쉼표 후보
            string commaCandidate;
            int lastComma = sub.LastIndexOf(',');
            if (lastComma >= 0 && lastComma < sub.Length - 1)
                commaCandidate = sub.Substring(lastComma + 1).Trim();
            else
                commaCandidate = sub;

            // 3. "하는" 후보 (하는 뒤부터)
            string haneunCandidate;
            int haneunIndex = sub.LastIndexOf("하는");
            if (haneunIndex >= 0 && haneunIndex + 2 < sub.Length)
                haneunCandidate = sub.Substring(haneunIndex + 2).Trim(); // "하는" 길이 2
            else
                haneunCandidate = sub;

            // 4. 두 후보 중 짧은 것을 선택
            string finalCandidate = (commaCandidate.Length <= haneunCandidate.Length)
                ? commaCandidate
                : haneunCandidate;

            return finalCandidate;
        }


        private List<int> ExtractDirectDependencies(string claimText)
        {
            Regex depRegex = new Regex(@"(?:제|청구항)\s*(\d+)\s*(?:항)?에 있어서");
            MatchCollection matches = depRegex.Matches(claimText);

            List<int> deps = new List<int>();
            foreach (Match m in matches)
            {
                if (int.TryParse(m.Groups[1].Value, out int num))
                    deps.Add(num);
            }

            return deps;
        }

        private List<int> BuildDependencyPathList(Claim claim, Dictionary<int, Claim> claimDict)
        {
            List<int> path = new List<int>();
            foreach (int dep in claim.DirectDependencies)
            {
                if (claimDict.TryGetValue(dep, out Claim depClaim))
                {
                    var subPath = BuildDependencyPathList(depClaim, claimDict);
                    path.AddRange(subPath);
                    path.Add(dep);
                }
            }
            // 연속 중복 제거
            List<int> deduped = new List<int>();
            foreach (var p in path)
            {
                if (deduped.Count == 0 || deduped[deduped.Count - 1] != p)
                    deduped.Add(p);
            }
            return deduped;
        }

        public class Claim
        {
            public int Number { get; set; }
            public string Category { get; set; }
            public List<int> DirectDependencies { get; set; } = new List<int>();
            public List<int> DependencyPath { get; set; } = new List<int>();
            public string Text { get; set; }
        }

        public void GenerateClaimTableAndAskTrackChange()
        {
            var app = Globals.ThisAddIn.Application;
            var sel = app?.Selection;

            // 선택 영역 체크 (기존 메서드와 동일한 기본 체크만 한 번 더)
            if (sel == null || sel.Range == null || string.IsNullOrWhiteSpace(sel.Range.Text ?? ""))
            {
                MessageBox.Show("청구항을 포함하는 영역을 선택해주세요.", "정보");
                return;
            }

            // 선택 영역에 【청구항 #】이 없는 경우 미리 막아줌
            Regex claimSplitRegex = new Regex(@"【청구항\s*(\d+)】");
            if (!claimSplitRegex.IsMatch(sel.Range.Text))
            {
                MessageBox.Show("선택 영역 내에서 【청구항 #】을 찾을 수 없습니다.", "정보");
                return;
            }

            // 1. 먼저 기존 기능으로 클레임 트리 생성
            GenerateClaimTableFromSelection();

            // 2. 생성된 클레임 트리를 수정할지 물어봄
            var result = MessageBox.Show(
                "생성된 클레임 트리 내용을 수정하시겠습니까?\r\n\r\n" +
                "예: 추적 변경 기능을 켜고 수정합니다.\r\n" +
                "아니오: 추적 변경 기능을 끄고 수정합니다.",
                "클레임 트리 수정 여부",
                MessageBoxButtons.YesNo,
                MessageBoxIcon.Question);

            // 3. 사용자의 선택에 따라 추적 변경 ON/OFF
            if (app.ActiveDocument != null)
            {
                if (result == DialogResult.Yes)
                {
                    // 수정: 추적 변경 ON
                    app.ActiveDocument.TrackRevisions = true;

                }
                else if (result == DialogResult.No)
                {
                    // 수정 안 함: 추적 변경 OFF
                    app.ActiveDocument.TrackRevisions = false;
                }
            }
            // 4. Contents 셀에서 【청구항 #】 삭제 (TrackRevisions 상태를 반영해서)
            RemoveClaimMarkersFromContents(app.ActiveDocument);

        }
        private void RemoveClaimMarkersFromContents(Word.Document doc)
        {
            if (doc == null) return;

            Word.Range rng = doc.Content;
            Word.Find find = rng.Find;

            find.ClearFormatting();
            // 공백/탭 1개 이상 + 숫자 1개 이상 + 】 + 단락기호
            find.Text = "【청구항[ ^t]{1,}[0-9]{1,}】^13";
            find.Replacement.ClearFormatting();
            // 줄바꿈까지 먹어버리고 한 줄로 붙이고 싶으면 뒤에 ^p를 안 넣음
            find.Replacement.Text = "일 실시예에 따르면, ";
            find.Forward = true;
            find.Wrap = Word.WdFindWrap.wdFindStop;
            find.MatchWildcards = true;
            find.Execute(Replace: Word.WdReplace.wdReplaceAll);

            // 1) "청구항 #에 있어서,^p" (쉼표 optional, 줄바꿈까지 삭제)
            //    예: "청구항 1에 있어서,^p"
            find.Text = "청구항 [0-9]{1,}에 있어서(,){0,1}^13";
            find.Replacement.Text = "";   // 전체 문장 + 줄바꿈 삭제
            find.Execute(Replace: Word.WdReplace.wdReplaceAll);

            // 2) "제# 항에 있어서,^p" (쉼표 optional, 줄바꿈까지 삭제)
            //    예: "제1 항에 있어서,^p"
            find.Text = "제[0-9]{1,} 항에 있어서(,){0,1}^13";
            find.Replacement.Text = "";
            find.Execute(Replace: Word.WdReplace.wdReplaceAll);
        }



    }
}

 