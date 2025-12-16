using Microsoft.Office.Interop.Word;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using WordAddIn1;
using Word = Microsoft.Office.Interop.Word;

namespace WordAddIn1
{
    public static class ClaimParaphraser
    {
        public static void ClaimPraphraser(Word.Document doc)
        {
            var app = Globals.ThisAddIn.Application;
            Word.Range selectedRange = app.Selection.Range;


            // 추적 변경 기능 켜기
            doc.TrackRevisions = true;


            //            Word.Range docRange = doc.Content;

            // 텍스트 치환
                // 방법 청구항
            TextChanger.ReplaceText(selectedRange, "동작을 더 포함하고,", "동작을 포함할 수 있다. 일 실시예에 따르면,"); 
            TextChanger.ReplaceText(selectedRange, "동작을 포함하고,^p", "동작을 포함할 수 있다. 일 실시예에 따르면, ");
            TextChanger.ReplaceText(selectedRange, "동작을 포함하고, ", "동작을 포함할 수 있다. 일 실시예에 따르면, ");
            TextChanger.ReplaceText(selectedRange, "동작;을 포함하고,^p", "동작을 포함할 수 있다. 일 실시예에 따르면, ");
            TextChanger.ReplaceText(selectedRange, "동작;을 포함하고, ", "동작을 포함할 수 있다. 일 실시예에 따르면, "); 
            TextChanger.ReplaceText(selectedRange, "동작을 포함하는,", "동작을 포함할 수 있다."); 
            TextChanger.ReplaceText(selectedRange, "동작;을 포함하는,", "동작을 포함할 수 있다.");
            TextChanger.ReplaceText(selectedRange, "동작; 및^p", "동작을 포함할 수 있다. 상기 방법은 ");
            TextChanger.ReplaceText(selectedRange, "동작;^p", "동작을 포함할 수 있다. 상기 방법은 ");
            TextChanger.ReplaceText(selectedRange, "동작 - ", "동작을 포함할 수 있고, ");
            TextChanger.ReplaceText(selectedRange, "함 - ", "할 수 있다.");
            TextChanger.ReplaceText(selectedRange, "됨 - ", "될 수 있다.");
            TextChanger.ReplaceText(selectedRange, "; 및^p", "일 실시예에 따르면, 상기 방법은 ");
//          TextChanger.ReplaceText(selectedRange, ";^p", "일 실시예에 따르면, 상기 방법은 ");
            TextChanger.ReplaceText(selectedRange, ":^p", " ");


            // 장치 청구항

            TextChanger.ReplaceText(selectedRange, "하고; 및^p", "하도록 할 수 있다. 상기 명령어들은, 상기 하나 이상의 프로세서에 의해 개별적으로 또는 공동적으로 실행될 때, 상기 전자 장치로 하여금 ");
            TextChanger.ReplaceText(selectedRange, "하고;^p", "하도록 할 수 있다. 상기 명령어들은, 상기 하나 이상의 프로세서에 의해 개별적으로 또는 공동적으로 실행될 때, 상기 전자 장치로 하여금 ");
            TextChanger.ReplaceText(selectedRange, "하고 - ", "하도록 할 수 있고, ");


            // 추적 변경 기능 끔
            //            doc.TrackRevisions = false;


        }


    }
}
