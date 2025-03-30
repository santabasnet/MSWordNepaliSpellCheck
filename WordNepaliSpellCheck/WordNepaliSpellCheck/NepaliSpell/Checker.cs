using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Threading;
using System.Threading.Tasks;
using Word = Microsoft.Office.Interop.Word;

namespace WordNepaliSpellCheck.NepaliSpell
{
    class Checker
    {       

        /// <summary>
        /// Extracts the selection of a page text with reference to the curernt selection
        /// and perform validation of nepali spelling with respect to the dictionary, local cache
        /// and the ignored words.  
        /// </summary>
        /// <param name="currentSelection"></param>
        public static void FindErrors(Word.Selection currentSelection)
        {
            if (COMUtility.IsTableSelected(currentSelection))
            {
                FindErrorsInTables();
                return;
            }

            if (COMUtility.IsHeaderFooterSelected(currentSelection))
            {
                FindErrorInHeaderFooter();
                return;
            }

            if (COMUtility.IsShapeSelected(currentSelection))
            {
                FindErrorInShapes();
                return;
            }

            FindErrorsInDocument();
        }

        /// <summary>
        /// Replace the wrong word in the document scope with the given suggested word..
        /// </summary>
        /// <param name="suggestedWord"></param>
        /// <param name="currentWord"></param>
        public static void ReplaceWrongWordInDocumentTables(String suggestedWord, NepaliWord currentWord)
        {
            /*Word.Range searchArea = Globals.ThisAddIn.nepaliSpellApp.ActiveDocument.Range();

            //1. Replace all the wrong words by the correct one.
            SearchUtility.ReplaceWrongWordInRange(suggestedWord, currentWord, searchArea);

            //2. Clear formatting.
            SearchUtility.ClearSpellingErrors(suggestedWord, searchArea);

            //3. Add a correct word to local cache from the suggestion.
            currentWord = new NepaliWord(currentWord.wordReference);
            SpellingService.AddCorrectWordToCache(currentWord.MakeACopy());

            ///4. Select the first replaced word.
            currentWord.wordReference.Select();*/
            ReplaceWrongWord(suggestedWord, currentWord);
        }

        /// <summary>
        /// Performs all the wrong word suggestions and replace in Headers and Footers.
        /// </summary>
        public static void ReplaceWrongWordInHeaderFooter(String suggestedWord, NepaliWord currentWord)
        {
            /*
            /// 1. Perform replace operation in header and footer range.
            object replaceAll = Word.WdReplace.wdReplaceAll;
            foreach (Word.Section section in currentWord.wordReference.Application.ActiveWindow.Document.Sections)
            {
                Word.Range headerRange = section.Headers[Word.WdHeaderFooterIndex.wdHeaderFooterPrimary].Range;
                SearchUtility.ReplaceWrongWordInRange(suggestedWord, currentWord, headerRange);
                SearchUtility.ClearSpellingErrors(suggestedWord, headerRange);

                Word.Range footerRange = section.Footers[Word.WdHeaderFooterIndex.wdHeaderFooterPrimary].Range;
                SearchUtility.ReplaceWrongWordInRange(suggestedWord, currentWord, footerRange);
                SearchUtility.ClearSpellingErrors(suggestedWord, footerRange);
            }
            
            /// 2. Update a correct word in the dictionary cache.
            currentWord = new NepaliWord(currentWord.wordReference);
            SpellingService.AddCorrectWordToCache(currentWord.MakeACopy());

            /// 3. Perform make a selection as a replacement has been started from here.            
            currentWord.wordReference.Select();
            */
            ReplaceWrongWord(suggestedWord, currentWord);
        }

        /// <summary>
        /// Performs all the wrong word suggestions and replace in Shapes.
        /// </summary>
        public static void ReplaceWrongWordInShapes(String suggestedWord, NepaliWord currentWord)
        {
            /*var shapeRange = currentWord.wordReference.Application.ActiveWindow.Selection.ShapeRange.TextFrame.TextRange;
            SearchUtility.ReplaceWrongWordInRange(suggestedWord, currentWord, shapeRange);
            SearchUtility.ClearSpellingErrors(suggestedWord, shapeRange);*/
            ReplaceWrongWord(suggestedWord, currentWord);
        }

        /// <summary>
        /// Perform replace of wrong
        /// </summary>
        /// <param name="suggestedWord"></param>
        /// <param name="currentWord"></param>
        private static void ReplaceWrongWord(String suggestedWord, NepaliWord currentWord)
        {

            // 1. Replace all the wrong words by the correct one remove spelling error in the document and table ranges.
            Word.Range searchArea = Globals.ThisAddIn.nepaliSpellApp.ActiveDocument.Range();
            SearchUtility.ReplaceWrongWordInRange(suggestedWord, currentWord, searchArea);
            SearchUtility.ClearSpellingErrors(suggestedWord, searchArea);

            // 2. Replace all the worng words in Shapes and clear the error formatting.
            foreach (Word.Shape shape in  currentWord.wordReference.Application.ActiveDocument.Shapes)
            {
                var shapeRange = shape.TextFrame.TextRange;
                SearchUtility.ReplaceWrongWordInRange(suggestedWord, currentWord, shapeRange);
                SearchUtility.ClearSpellingErrors(suggestedWord, shapeRange);
            }

            // 3. Replace all the worng words in Header and Footers and clear the error formatting.
            object replaceAll = Word.WdReplace.wdReplaceAll;
            foreach (Word.Section section in currentWord.wordReference.Application.ActiveWindow.Document.Sections)
            {
                Word.Range headerRange = section.Headers[Word.WdHeaderFooterIndex.wdHeaderFooterPrimary].Range;
                SearchUtility.ReplaceWrongWordInRange(suggestedWord, currentWord, headerRange);
                SearchUtility.ClearSpellingErrors(suggestedWord, headerRange);

                Word.Range footerRange = section.Footers[Word.WdHeaderFooterIndex.wdHeaderFooterPrimary].Range;
                SearchUtility.ReplaceWrongWordInRange(suggestedWord, currentWord, footerRange);
                SearchUtility.ClearSpellingErrors(suggestedWord, footerRange);
            }

            // 4. Update a correct word in the dictionary cache.
            currentWord = new NepaliWord(currentWord.wordReference);
            SpellingService.AddCorrectWordToCache(currentWord.MakeACopy());

            // 5. Perform make a selection as a replacement has been started from here.            
            currentWord.wordReference.Select();
        }

        /// <summary>
        /// Finds all the wrong spell words in Header and Footers.
        /// </summary>
        private static void FindErrorInHeaderFooter()
        {
            var allWords = RetrieveNepaliWordsInHeaderFooter();
                
            allWords.ForEach(currentWord => currentWord.MakeSpellError());
        }

        /// <summary>
        /// Retrieves the list of words available in the header and footer sections. 
        /// </summary>
        /// <returns>listOfNepaliWords</returns>
        private static List<NepaliWord> RetrieveNepaliWordsInHeaderFooter()
        {
            var headerFooterSection = Globals.ThisAddIn.nepaliSpellApp.ActiveDocument.Sections[1];
            var headerRange = headerFooterSection.Headers[Word.WdHeaderFooterIndex.wdHeaderFooterPrimary].Range;
            var footerRange = headerFooterSection.Footers[Word.WdHeaderFooterIndex.wdHeaderFooterPrimary].Range;            
            return new List<Word.Range>(){ headerRange, footerRange }.Select(RetrieveNepaliWords).SelectMany(item => item).ToList();
        }

        /// <summary>
        /// Accumulates all the spelling errors in the table and list out all the
        /// mis-spelled words with colored underline.
        /// </summary>
        private static void FindErrorsInTables()
        {
            GetTableCells()
                .Select(RetrieveNepaliWords)
                .SelectMany(item => item).ToList()
                .ForEach(currentWord => currentWord.MakeSpellError());
        }

        /// <summary>
        /// Accumulates all the table cells as the list of word ranges.
        /// </summary>
        /// <returns>list of word ranges.</returns>
        private static List<Word.Range> GetTableCells()
        {
            List<Word.Range> tableCellsRanges = new List<Word.Range>();
            foreach (Word.Table currentTable in Globals.ThisAddIn.nepaliSpellApp.ActiveDocument.Tables)
                foreach (Word.Cell tableCell in currentTable.Range.Cells)
                    tableCellsRanges.Add(tableCell.Range);
            return tableCellsRanges;
        }

        /// <summary>
        /// Finds spelling error in the words shapes, digrams.
        /// </summary>
        private static void FindErrorInShapes()
        {
            RetrieveNepaliWordsOfShapes().ForEach(nepaliWord => nepaliWord.MakeSpellError());
        }

        /// <summary>
        /// Retrives Nepali words appeared in the shapes.
        /// </summary>
        /// <returns></returns>
        private static List<NepaliWord> RetrieveNepaliWordsOfShapes()
        {
            return GetListOfShapes(Globals.ThisAddIn.nepaliSpellApp.ActiveDocument)
                 .Select(RetrieveNepaliWords)
                 .SelectMany(selection => selection).ToList();
        }

        private static List<Word.Range> GetListOfShapes(Word.Document currentDocument)
        {
            List<Word.Range> shapesRange = new List<Word.Range>();
            foreach (Word.Shape shape in currentDocument.Shapes)
                shapesRange.Add(shape.TextFrame.TextRange);
            return shapesRange;
        }

        /// <summary>
        /// Finds spelling errors in the document from the retried words.
        /// </summary>
        private static void FindErrorsInDocument()
        {            
            //RetrieveNepaliWordsOfDocument().ForEach(currentWord => currentWord.MakeSpellError());
            GetListOfParagraphs(Globals.ThisAddIn.nepaliSpellApp.ActiveDocument).ForEach(paragraph => CheckSpellingErrors(paragraph));
        }

        /// <summary>
        /// Perform error tagging for a Nepali word in the paragraph.
        /// </summary>
        /// <param name="paragraph"></param>
        private static void CheckSpellingErrors(Word.Paragraph paragraph)
        {
            RetrieveNepaliWords(paragraph.Range).ForEach(word => word.MakeSpellError());
            Thread.Sleep(TextUtility.ELAPSED_TIME);
        }

        /// <summary>
        /// Retrieves all the Nepali words for the documents.
        /// </summary>
        /// <returns>The list of Nepali words.</returns>
        private static List<NepaliWord> RetrieveNepaliWordsOfDocument()
        {
            return GetListOfParagraphs(Globals.ThisAddIn.nepaliSpellApp.ActiveDocument)
                    .Select(RetrieveNepaliWords)
                    .SelectMany(selection => selection)
                    .ToList();
        }

        /// <summary>
        /// Returns the list of paragraphs of the current document.
        /// </summary>
        /// <param name="currentDocument"></param>
        /// <returns>A list of paragraph ranges.</returns>
        private static List<Word.Paragraph> GetListOfParagraphs(Word.Document currentDocument)
        {
            List<Word.Paragraph> paragraphs = new List<Word.Paragraph>();
            foreach (Word.Paragraph paragraph in currentDocument.Paragraphs) paragraphs.Add(paragraph);
            return paragraphs;
        }

        /// <summary>
        /// Retrieves a list of Nepali words from the current paragraph before identifying
        /// the spelling errors.
        /// </summary>
        /// <param name="selectedPageRange"></param>
        /// <returns></returns>
        private static List<NepaliWord> RetrieveNepaliWords(Word.Paragraph selectedParagraph)
        {
            return RetrieveNepaliWords(selectedParagraph.Range);
        }

        /// <summary>
        /// Retrieves a list of Nepali words from the current page before identifying the spelling errors.
        /// It also removes all the corrected words.
        /// </summary>
        /// <param name="selectedRange"></param>
        /// <returns>listOfErrorWords</returns>
        private static List<NepaliWord> RetrieveNepaliWords(Word.Range selectedRange)
        {
            List<NepaliWord> allWords = RetriveRangeWords(selectedRange);
            List<Tuple<NepaliWord, Boolean>>  verifiedWords = NepaliWord.VerifySpellingErrors(allWords);            
            /* *
             * Remove spelling error, if corrected previously or by placing space in between.
             * */
            verifiedWords.ForEach(word => {               
                if (word.Item2) word.Item1.RemoveSpellError();                
            });
            /* *
             * Finally return all the wrong words.
             * */
            return verifiedWords
                .Where(entry => !entry.Item2)
                .Select(entry => entry.Item1)
                .ToList();
        }

        /// <summary>
        /// Header footer range works by single word error processing strategy. 
        /// Old word accumulation with one by one error checking of words.
        /// </summary>
        /// <param name="selectedRange"></param>
        /// <returns>listOfErrorWords</returns>
        private static List<NepaliWord> RetrieveNepaliHeaderWords(Word.Range selectedRange)
        {
            //Old work for spelling check with one by one strategy.
            List<NepaliWord> nepaliWords = new List<NepaliWord>();
            foreach (Word.Range aWord in FormatNepaliWordRanges(selectedRange))
            {
                NepaliWord nepaliWord = RetrieveNepaliWord(aWord);
                if (nepaliWord.HasSpellingError()) nepaliWords.Add(nepaliWord);
                else nepaliWord.RemoveSpellError();
            }
            return nepaliWords;
        }
        /// <summary>
        /// Retrieve Nepali words of the given range.
        /// </summary>
        /// <param name="selectedRange"></param>
        /// <returns>listOfRangeWords</returns>
        private static List<NepaliWord> RetriveRangeWords(Word.Range selectedRange)
        {
            List<NepaliWord> allWords = new List<NepaliWord>();
            foreach (Word.Range aWord in FormatNepaliWordRanges(selectedRange))
            {
                NepaliWord nepaliWord = RetrieveNepaliWord(aWord);
                if(nepaliWord.IsNepaliEncodedText()) allWords.Add(nepaliWord);
            }
            return allWords;
        }

        /// <summary>
        /// Utility that converts a range to multiple word ranges.
        /// (Useful for multiple encoding feature.)
        /// Needs to re-work later.
        /// </summary>
        /// <param name="selectedRange"></param>
        /// <returns></returns>
        private static List<Word.Range> FormatNepaliWordRanges(Word.Range selectedRange)
        {
            List<Word.Range> wordRanges = new List<Word.Range>();
            foreach (Word.Range aWord in selectedRange.Words) wordRanges.Add(aWord);            
            return wordRanges;
            // Needs to re-work with later for Nepali TTF fonts with its feature.
        }

        /// <summary>
        /// Extracts a word that represents a Nepali word, with given that of word range.
        /// </summary>
        /// <param name="wordRange"></param>
        /// <returns></returns>
        public static NepaliWord RetrieveNepaliWord(Word.Range wordRange)
        {
            if (String.IsNullOrEmpty(wordRange.Text)) return NepaliWord.Empty();
            else return new NepaliWord(wordRange.Start, wordRange.End, wordRange.Text, wordRange);

            ///This code section is useful for TTF font only.
            ///In case of Unicode text, needs some simpler logic.
            /*Word.Range sentenceRange = wordRange.Sentences.Last;
            if (IsEmptyText(sentenceRange, wordRange)) return NepaliWord.Empty();
            Word.Range newRange = NepaliWordRange(sentenceRange, wordRange);
            return new NepaliWord(newRange.Start, newRange.End, newRange.Text, wordRange);*/
        }

        /// <summary>
        /// Extracts a word that represents a Nepali Word text from the current selection.
        /// </summary>
        /// <param name="currentSelection"></param>
        /// <returns>An instance of Nepali Word.</returns>
        public static NepaliWord RetrieveNepaliWord(Word.Selection currentSelection)
        {
            Word.Range wordRange = currentSelection.Words.Last;
            return RetrieveNepaliWord(wordRange);
        }

        /// <summary>
        /// Verifies whether the current selection of word during right click needs 
        /// to context menu or not.
        /// </summary>
        /// <param name="currentSelection"></param>
        /// <returns>true if the current selection contains Nepali word.</returns>
        public static Tuple<Boolean, NepaliWord> NeedsToAddMenuItem(Word.Selection currentSelection)
        {
            NepaliWord nepaliWord = RetrieveNepaliWord(currentSelection);
            return Tuple.Create<Boolean, NepaliWord>((nepaliWord.IsFormattedNepali() && nepaliWord.IsErrorWord()), nepaliWord);
        }

        /// <summary>
        /// Builds the range for the Nepali Word representation, especially delimited by the
        /// white spaces. 
        /// It is necessary that, TTF words are encoded by special single byte representation
        /// of the Nepali font.
        /// </summary>
        /// <param name="sentenceRange"></param>
        /// <param name="wordRange"></param>
        /// <returns>A range that encorporates the Nepali Word in the current selection.</returns>
        private static Word.Range NepaliWordRange(Word.Range sentenceRange, Word.Range wordRange)
        {
            Word.Range prefixRange = PrefixWordDelimeter(sentenceRange, wordRange);
            Word.Range postfixRange = PostfixWordDelimeter(sentenceRange, wordRange);
            return sentenceRange.Document.Range(prefixRange.Start, postfixRange.End);
        }

        /// <summary>
        /// Make a lookup to the preceeding whitespace.
        /// </summary>
        /// <param name="sentenceRange"></param>
        /// <param name="wordRange"></param>
        /// <returns>The preceeding word range.</returns>
        private static Word.Range PrefixWordDelimeter(Word.Range sentenceRange, Word.Range wordRange)
        {
            Word.Range beginRange = wordRange.Characters.First.Duplicate;
            for (int startIndex = wordRange.Start; startIndex >= sentenceRange.Start; startIndex--)
            {
                if (TextUtility.IsNonPrintable(beginRange.Characters.First.Text)) break;
                object unitCharacter = Word.WdUnits.wdCharacter;
                object count = -1;
                beginRange.MoveEnd(ref unitCharacter, ref count);
            }
            return beginRange;
        }

        /// <summary>
        /// Make a lookup to the succeding whitespace.
        /// </summary>
        /// <param name="sentenceRange"></param>
        /// <param name="wordRange"></param>
        /// <returns>The succeding word range.</returns>
        private static Word.Range PostfixWordDelimeter(Word.Range sentenceRange, Word.Range wordRange)
        {
            Word.Range endRange = wordRange.Duplicate;
            for (int startIndex = wordRange.End; startIndex < sentenceRange.End; startIndex++)
            {
                if (String.IsNullOrWhiteSpace(endRange.Characters.Last.Text)) break;
                object unitCharacter = Word.WdUnits.wdCharacter;
                object count = 1;
                int result = endRange.MoveEnd(ref unitCharacter, ref count);
            }
            return endRange;
        }

        /// <summary>
        ///  Checks if the current selection of word and sentence range contains text or not.
        /// </summary>
        /// <param name="sentenceRange"></param>
        /// <param name="wordRange"></param>
        /// <returns>true if the current selection of sentence and word does not contain any text.</returns>
        private static Boolean IsEmptyText(Word.Range sentenceRange, Word.Range wordRange)
        {
            return String.IsNullOrWhiteSpace(sentenceRange.Text) || String.IsNullOrWhiteSpace(wordRange.Text);
        }

        /// <summary>
        /// Returns the current page range of the word document with reference to the selected character index.
        /// </summary>
        /// <param name="currentDocument"></param>
        /// <param name="characterIndex"></param>
        /// <returns>Current page range.</returns>
        private static Word.Range GetPageRangeOfSelection(Word.Document currentDocument, int characterIndex)
        {
            int numberOfPages = (int)currentDocument.Content.Information[Word.WdInformation.wdNumberOfPagesInDocument];
            int lastCharInDocument = currentDocument.Range(0).End;
            Word.Range rangeStart = null;
            Word.Range rangeEnd = null;
            for (int page = 1; page <= numberOfPages; page++)
            {
                object objectLink = Word.WdGoToItem.wdGoToPage;
                object direction = Word.WdGoToDirection.wdGoToAbsolute;
                object pageCount = page;
                rangeStart = currentDocument.GoTo(ref objectLink, ref direction, ref pageCount);
                object countPlusOne = page + 1;
                rangeEnd = currentDocument.GoTo(ref objectLink, ref direction, ref countPlusOne);
                if (rangeStart.Start <= characterIndex && characterIndex <= rangeEnd.End)
                    return currentDocument.Range(rangeStart.Start, rangeEnd.End);

            }
            if (rangeEnd != null & rangeEnd.Start <= characterIndex && characterIndex <= lastCharInDocument)
                return currentDocument.Range(rangeEnd.Start, lastCharInDocument);
            return null;
        }

        /// <summary>
        /// Returns the list of page range.
        /// </summary>
        /// <param name="currentDocument"></param>
        /// <returns>The page ranges of the current document.</returns>
        private static List<Word.Range> GetListOfPages(Word.Document currentDocument)
        {
            return Enumerable
                .Range(1, (int)currentDocument.Content.Information[Word.WdInformation.wdNumberOfPagesInDocument])
                .Select(index => GetPageRange(currentDocument, index)).ToList();
        }

        /// <summary>
        /// Gets the page ranges of the given page index.
        /// </summary>
        /// <param name="currentDocument"></param>
        /// <param name="pageIndex"></param>
        /// <returns>The current pageRange.</returns>
        private static Word.Range GetPageRange(Word.Document currentDocument, int pageIndex)
        {
            object objectLink = Word.WdGoToItem.wdGoToPage;
            object direction = Word.WdGoToDirection.wdGoToAbsolute;
            object pageCount = pageIndex;
            Word.Range rangeStart = currentDocument.GoTo(ref objectLink, ref direction, ref pageCount);
            object countPlusOne = pageIndex + 1;
            Word.Range rangeEnd = currentDocument.GoTo(ref objectLink, ref direction, ref countPlusOne);
            Word.Range result = currentDocument.Range(rangeStart.Start, rangeEnd.End);
            return currentDocument.Range(rangeStart.Start, rangeEnd.End);
        }

    }
}
