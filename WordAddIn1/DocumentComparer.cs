using System;
using System.IO;
using System.Linq;
using System.Windows.Forms;
using Microsoft.Office.Interop.Word;
using Microsoft.Office.Core;

namespace WordAddIn1
{
    public static class DocumentComparer
    {
        public static void CompareActiveDocument()
        {
            try
            {
                var app = Globals.ThisAddIn.Application;
                Document activeDoc = app.ActiveDocument;
                if (activeDoc == null)
                {
                    MessageBox.Show("활성 문서가 없습니다.", "알림");
                    return;
                }

                string originalPath = activeDoc.FullName;
                string folderPath = null;

                if (Uri.IsWellFormedUriString(originalPath, UriKind.Absolute))
                {
                    var uri = new Uri(originalPath);
                    folderPath = Path.GetDirectoryName(uri.LocalPath);
                }
                else
                {
                    folderPath = Path.GetDirectoryName(originalPath);
                }

                if (folderPath.StartsWith("https:", StringComparison.OrdinalIgnoreCase))
                {
                    MessageBox.Show("이 문서는 클라우드(OneDrive)에 저장되어 있어 로컬 경로를 탐색할 수 없습니다.\n비교할 문서를 직접 선택하세요.", "안내");
                    // 바로 파일 선택창 열기
                }
                if (string.IsNullOrEmpty(folderPath) || !Directory.Exists(folderPath))
                    folderPath = Environment.CurrentDirectory;

                string fileNameWithoutExt = Path.GetFileNameWithoutExtension(originalPath);

                // 대소문자 구분 없이 동일 파일명.rtf 찾기
                string rtfPath = null;
                try
                {
                    rtfPath = Directory.GetFiles(folderPath, "*.rtf", SearchOption.TopDirectoryOnly)
                        .FirstOrDefault(f => string.Equals(
                            Path.GetFileNameWithoutExtension(f),
                            fileNameWithoutExt,
                            StringComparison.OrdinalIgnoreCase));
                }
                catch
                {
                    rtfPath = null;
                }

                Document revisedDoc = null;

                if (rtfPath != null && File.Exists(rtfPath))
                {
                    DialogResult result = MessageBox.Show(
                        $"동일한 파일명(.rtf) 문서가 있습니다.\n\n'{Path.GetFileName(rtfPath)}' 파일과 비교하시겠습니까?",
                        "문서 비교",
                        MessageBoxButtons.YesNo,
                        MessageBoxIcon.Question);

                    if (result == DialogResult.Yes)
                    {
                        revisedDoc = app.Documents.Open(rtfPath, ReadOnly: true);
                    }
                    else
                    {
                        return;
                    }
                }
                else
                {
                    // 알림 후 파일 선택창 열기
                    MessageBox.Show(
                        "동일한 이름의 .rtf 문서를 찾을 수 없습니다.\n비교할 문서를 직접 선택하세요.",
                        "문서 선택",
                        MessageBoxButtons.OK,
                        MessageBoxIcon.Information);

                    Microsoft.Office.Core.FileDialog fileDialog = app.FileDialog[MsoFileDialogType.msoFileDialogFilePicker];
                    fileDialog.Title = "비교할 문서를 선택하세요.";
                    fileDialog.Filters.Clear();
                    fileDialog.Filters.Add("모든 Word 파일", "*.docx;*.doc;*.rtf;*.txt");
                    app.ChangeFileOpenDirectory(folderPath); // 기본 폴더 설정

                    if (fileDialog.Show() == -1)
                    {
                        string selectedFile = fileDialog.SelectedItems.Item(1);
                        revisedDoc = app.Documents.Open(selectedFile, ReadOnly: true);
                    }
                    else
                    {
                        return;
                    }
                }

                if (revisedDoc != null)
                {
                    Document resultDoc = app.CompareDocuments(
                        activeDoc,
                        revisedDoc,
                        WdCompareDestination.wdCompareDestinationNew,
                        WdGranularity.wdGranularityWordLevel,
                        true,  // CompareFormatting
                        true,  // CompareCaseChanges
                        true,  // CompareWhitespace
                        true,  // CompareTables
                        true,  // CompareHeaders
                        true,  // CompareFootnotes
                        true,  // CompareTextboxes
                        true,  // CompareFields
                        true,  // CompareComments
                        true,  // CompareMoves
                        app.UserName, // RevisedAuthor
                        false  // IgnoreAllComparisonWarnings
                    );

                    resultDoc.Activate();
                    revisedDoc.Close(WdSaveOptions.wdDoNotSaveChanges);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"오류가 발생했습니다:\n{ex.Message}", "오류", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }
    }
}
