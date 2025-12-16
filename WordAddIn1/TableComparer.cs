using System;
using System.IO;
using System.Runtime.InteropServices;
using Microsoft.Office.Interop.Word;
using Microsoft.VisualBasic;

namespace WordAddIn1
{
    public class TableComparer
    {
        private Application wordApp;
        private Document activeDoc;

        public TableComparer(Application application)
        {
            wordApp = application ?? throw new ArgumentNullException(nameof(application));
            activeDoc = wordApp.ActiveDocument;
        }

        public void CompareEachRow()
        {
            if (wordApp.Selection.Tables.Count == 0)
            {
                System.Windows.Forms.MessageBox.Show("표 안에 커서를 두세요.");
                return;
            }

            Table table = wordApp.Selection.Tables[1];

            string author = activeDoc.BuiltInDocumentProperties["Author"].Value?.ToString() ?? "Unknown";
            string revisedAuthor = PromptForRevisedAuthor(author);

            for (int i = 1; i <= table.Rows.Count; i++)
            {
                Cell leftCell = table.Cell(i, 1);
                Cell rightCell = table.Cell(i, 2);

                string leftText = TrimCellText(leftCell.Range.Text);
                string rightText = TrimCellText(rightCell.Range.Text);

                string leftPath = @"C:\1.docx";
                string rightPath = @"C:\2.docx";
                string rePath = @"C:\3.docx";

                SaveTempDoc(leftText, leftPath);
                SaveTempDoc(rightText, rightPath);

                Document leftDoc = wordApp.Documents.Open(leftPath, ReadOnly: false, Visible: false);
                Document rightDoc = wordApp.Documents.Open(rightPath, ReadOnly: false, Visible: false);

                Document compareDoc = wordApp.CompareDocuments(
                    OriginalDocument: leftDoc,
                    RevisedDocument: rightDoc,
                    Destination: WdCompareDestination.wdCompareDestinationNew,
                    Granularity: WdGranularity.wdGranularityWordLevel,
                    CompareFormatting: true,
                    CompareCaseChanges: true,
                    CompareWhitespace: true,
                    CompareTables: true,
                    CompareHeaders: true,
                    CompareFootnotes: true,
                    CompareTextboxes: true,
                    CompareFields: true,
                    CompareComments: true,
                    CompareMoves: true,
                    RevisedAuthor: revisedAuthor,
                    IgnoreAllComparisonWarnings: false
                );

                compareDoc.Activate();
                compareDoc.Content.Copy();
                compareDoc.SaveAs2(rePath, WdSaveFormat.wdFormatXMLDocument);
                compareDoc.Close(WdSaveOptions.wdSaveChanges);
                Marshal.ReleaseComObject(compareDoc);

                rightCell.Range.Paste();

                leftDoc.Close(WdSaveOptions.wdDoNotSaveChanges);
                rightDoc.Close(WdSaveOptions.wdDoNotSaveChanges);
                Marshal.ReleaseComObject(leftDoc);
                Marshal.ReleaseComObject(rightDoc);

                try
                {
                    File.Delete(leftPath);
                    File.Delete(rightPath);
                    File.Delete(rePath);
                }
                catch (IOException ex)
                {
                    System.Diagnostics.Debug.WriteLine("파일 삭제 실패: " + ex.Message);
                }
            }
        }

        private string TrimCellText(string text)
        {
            return text.Length >= 2 ? text.Substring(0, text.Length - 2) : text;
        }

        private string PromptForRevisedAuthor(string defaultAuthor)
        {
            return Interaction.InputBox(
                "편집자 이름을 입력하세요:", "편집자 이름 입력", defaultAuthor);
        }

        private void SaveTempDoc(string content, string path)
        {
            Document tempDoc = wordApp.Documents.Add();
            tempDoc.Content.Text = content;
            tempDoc.SaveAs2(path, WdSaveFormat.wdFormatXMLDocument);
            tempDoc.Close(WdSaveOptions.wdSaveChanges);
            Marshal.ReleaseComObject(tempDoc);
        }
    }
}
