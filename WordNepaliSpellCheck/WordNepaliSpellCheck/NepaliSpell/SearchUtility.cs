using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Word = Microsoft.Office.Interop.Word;

namespace WordNepaliSpellCheck.NepaliSpell
{
    class SearchUtility
    {

        /// <summary>
        /// Makes wrong word search and replacement by suggested word in the given range.
        /// </summary>
        /// <param name="suggestedWord"></param>
        /// <param name="currentWord"></param>
        /// <param name="givenRange"></param>
        public static void ReplaceWrongWordInRange(String suggestedWord, NepaliWord currentWord, Word.Range givenRange)
        {
            Object missing = System.Reflection.Missing.Value;
            object searchText = currentWord.wordText;
            object replaceText = suggestedWord;
            object matchCase = false;
            object matchWholeWord = true;
            object replaceAll = Word.WdReplace.wdReplaceAll;
            Word.Range searchArea = givenRange;
            Word.Find searchObject = searchArea.Find;
            searchObject.Text = currentWord.wordText;
            searchObject.Replacement.Text = suggestedWord;
            searchObject.Execute(ref searchText, ref matchCase, ref matchWholeWord, ref missing, ref missing,
                                ref missing, ref missing, ref missing, ref missing, ref replaceText,
                                ref replaceAll, ref missing, ref missing, ref missing, ref missing);
        }

        /// <summary>
        /// Find all the replaced words and clear its formatting as a wrong spelling.
        /// </summary>
        /// <param name="correctWord"></param>
        /// <param name="givenRange"></param>
        public static void ClearSpellingErrors(String correctWord, Word.Range givenRange)
        {
            Word.Range searchArea = givenRange;
            Object missing = System.Reflection.Missing.Value;
            object newText = correctWord;
            searchArea.Find.ClearFormatting();
            searchArea.Find.Execute(ref newText, ref missing, ref missing, ref missing, ref
                missing, ref missing, ref missing, ref missing, ref missing, ref missing,
                ref missing, ref missing, ref missing, ref missing, ref missing);
            while (searchArea.Find.Found)
            {
                //Replace is not working for a single word without white spaces.
                //if(givenRange.Words.Count == 1) searchArea.Underline = Word.WdUnderline.wdUnderlineNone;                
                if (searchArea.Underline == Word.WdUnderline.wdUnderlineWavy && searchArea.Font.UnderlineColor == Word.WdColor.wdColorOrange)
                    searchArea.Underline = Word.WdUnderline.wdUnderlineNone;
                searchArea.Find.Execute(ref newText, ref missing, ref missing, ref missing, ref
                missing, ref missing, ref missing, ref missing, ref missing, ref missing,
                ref missing, ref missing, ref missing, ref missing, ref missing);
            }
        }

    }
}
