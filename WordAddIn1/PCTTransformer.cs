using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Word = Microsoft.Office.Interop.Word;


namespace WordAddIn1
{
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


        }

    }
}
