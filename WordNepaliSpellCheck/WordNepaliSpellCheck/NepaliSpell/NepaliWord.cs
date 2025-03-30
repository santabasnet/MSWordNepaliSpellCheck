using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Microsoft.Office.Interop.Word;
using System.Diagnostics;
using System.Web.Script.Serialization;
using System.Windows.Forms;

namespace WordNepaliSpellCheck.NepaliSpell
{
    public class NepaliWord
    {
        /// <summary>
        /// The start of word selection range.
        /// </summary>
        public int Start { get; set; }

        /// <summary>
        /// The end of word selection range.
        /// </summary>
        public int End { get; set; }

        /// <summary>
        /// The unicode text representation of the given Nepali word.
        /// </summary>
        public String wordText { get; set; }

        /// <summary>
        /// Default font name, if the text is encoded with unicode character range.
        /// </summary>
        public String fontName { get; set; }

        /// <summary>
        /// A reference word i.e. useful to extract the font and other information.
        /// </summary>
        public Range wordReference { get; set; }

        /// <summary>
        /// A constructor to initialize the given Nepali word representation.
        /// </summary>
        /// <param name="Start"></param>
        /// <param name="End"></param>
        /// <param name="wordText"></param>
        /// <param name="wordReference"></param>
        public NepaliWord(int Start, int End, String wordText, Range wordReference)
        {
            this.Start = Start;
            this.End = End;
            this.wordText = wordText.Trim();
            this.wordReference = wordReference;
            this.fontName = InitializeFontName();
        }

        /// <summary>
        /// Overloaded constructor with word text and the word reference.
        /// </summary>
        /// <param name="wordText"></param>
        /// <param name="wordReference"></param>
        public NepaliWord(String wordText, Range wordReference)
        {
            this.Start = wordReference.Start;
            this.End = wordReference.End;
            this.wordText = wordText.Trim();
            this.wordReference = wordReference;
            this.fontName = InitializeFontName();
        }

        /// <summary>
        /// Overloaded constructor with the word reference only.
        /// </summary>
        /// <param name="wordReference"></param>
        public NepaliWord(Range wordReference)
        {
            this.Start = wordReference.Start;
            this.End = wordReference.End;
            this.wordText = wordReference.Text.Trim();
            this.wordReference = wordReference;
            this.fontName = InitializeFontName();
        }

        /// <summary>
        /// Makes a copy of the current object.
        /// </summary>
        /// <returns></returns>
        public NepaliWord MakeACopy()
        {
            return new NepaliWord(this.Start, this.End, this.wordText, this.wordReference);
        }

        /// <summary>
        /// Returns the font name associated with the given Nepali word.
        /// </summary>
        /// <returns>A font name in the word application.</returns>
        private String InitializeFontName()
        {
            if (this.wordReference == null) return TextUtility.EMPTY;
            if (IsUnicodeWord()) return TextUtility.UNICODE;
            else return this.wordReference.FormattedText.Font.Name.ToUpper();
        }

        /// <summary>
        /// Formats the given text of word with reference to the wrong spelled word. 
        /// </summary>
        /// <param name="wordText"></param>
        /// <returns>The formatted text of the given word before copying.</returns>
        private String FormattedDelimeters(String wordText)
        {
            return String.Join(TextUtility.EMPTY, new String[] {
                String.Join(TextUtility.EMPTY, this.wordText.TakeWhile(Char.IsWhiteSpace)),
                wordText,
                String.Join(TextUtility.EMPTY, this.wordText.Reverse().TakeWhile(Char.IsWhiteSpace))
            });
        }

        /// <summary>
        /// Generates a unique key in the current page index of a word, that is used 
        /// for the tokenization purpose.
        /// </summary>
        /// <returns>A unique key of a word.</returns>
        public String GetUniqueKey()
        {
            if (IsEmpty()) return TextUtility.EMPTY;
            else return $"{this.Start}_{this.wordText}_{this.End}";
        }

        /// <summary>
        /// Build a new Nepali word for replacement with reference to the suggestions.
        /// </summary>
        /// <param name="wordText"></param>
        /// <returns></returns>
        public NepaliWord CopyWith(String wordText)
        {
            wordReference = wordReference.Document.Range(Start, Start + wordText.Length);
            return new NepaliWord(wordReference.Start, wordReference.End, wordText, wordReference);
        }

        /// <summary>
        /// Verifies the current Nepali word is encoded by Unicode text or not.
        /// </summary>
        /// <returns>true if the current word is encoded by Unicode devanagari text.</returns>
        public Boolean IsUnicodeWord()
        {
            return IsNepaliEncodedText();
        }

        /// <summary>
        /// Verifies if the given Nepali word is Empty or not.
        /// </summary>
        /// <returns>true if the current word is Empty.</returns>
        public Boolean IsEmpty()
        {
            return wordReference == null || String.IsNullOrWhiteSpace(wordReference.Text) || (Start == End) || (wordReference.Start == wordReference.End);
        }

        /// <summary>
        /// Verifies if the given Nepali word is not Empty.
        /// </summary>
        /// <returns>true if the Nepali word contains some text in it.</returns>
        public Boolean IsNotEmpty()
        {
            return !IsEmpty();
        }

        /// <summary>
        /// Verifies if the current Nepali word is spell error or not. 
        /// </summary>
        /// <returns>true if the current work is spell error.</returns>
        public Boolean IsErrorWord()
        {
            return !IsCorrectWord();
        }

        /// <summary>
        /// Checks whether the given word has mis-spelling or not.
        /// </summary>
        /// <returns>true if the word is mis-spelled according to the Nepali language.</returns>
        public Boolean HasSpellingError()
        {
            return !String.IsNullOrEmpty(this.GetUniqueKey()) && this.IsFormattedNepali() && this.IsErrorWord();
        }

        /// <summary>
        /// Verifies if the current word has spell error underlined.
        /// </summary>
        /// <returns>true if has already underlined spell error.</returns>
        public Boolean HasUnderlineSpellError()
        {
            return IsAlreadyUnderlined();
        }

        /// <summary>
        /// Verifies if the current Nepali word is present in the local cache with mis-spelled status
        /// or not.
        /// </summary>
        /// <returns>true if the current is wrongly spelled.</returns>
        public Boolean IsLocallyPresentError()
        {
            return SpellingService.IsLocallyPresentError(this.wordText, this.fontName);
        }

        /// <summary>
        /// Validates if the current is formatted with Nepali fonts or not.
        /// It also checks, if the current text is encoded by Nepali devanagari characters.
        /// </summary>
        /// <returns>true if the current word is of Nepali text.</returns>
        public Boolean IsFormattedNepali()
        {
            //Check if it falls in unicode range or has TTF Nepali fonts.
            return IsNepaliEncodedText() || SpellingService.IsNepaliFont(this.fontName);
        }

        /// <summary>
        /// Checks if the current word is encoded with Nepali Unicode or not. The strategy is, half of characters
        /// should fall inside the Devanagari UTF-8 range.
        /// </summary>
        /// <returns>true if the word text has majority of characters in Devanagari UTF-8</returns>
        public Boolean IsNepaliEncodedText()
        {
            if (IsEmpty()) return false;
            String currentText = this.wordReference.Text.Trim();
            int devanagariStats = currentText
                .ToCharArray().Where(ch => TextUtility.IsUnicode(ch) || Char.IsPunctuation(ch)).Count();
            return (float)(currentText.Length * 0.95) <= (float)devanagariStats;
        }

        /// <summary>
        /// Checks if the current word is already underlined or not.
        /// </summary>
        /// <returns>true if it is already underlined by red wavy color.</returns>
        public Boolean IsAlreadyUnderlined()
        {           
            return wordReference.Underline == WdUnderline.wdUnderlineWavy && wordReference.Font.UnderlineColor == WdColor.wdColorOrange;
        }

        /// <summary>
        /// Makes the current Nepali word mis-spelled with underlining by red wavy line.
        /// </summary>
        public void MakeSpellError()
        {
            if (!IsAlreadyUnderlined())
            {
                Range currentRange = WordReferenceRange();
                currentRange.Underline = WdUnderline.wdUnderlineWavy;
                currentRange.Font.UnderlineColor = WdColor.wdColorOrange;
            }
        }

        /// <summary>
        /// Makes the current Nepali word mis-spelled with underlining by red wavy line.
        /// </summary>
        public void RemoveSpellError()
        {
            if (IsAlreadyUnderlined()) this.wordReference.Underline = WdUnderline.wdUnderlineNone;
        }

        /// <summary>
        /// Verify whether the given word is present in local or not.
        /// </summary>
        /// <returns></returns>
        public int VerifyFromLocalCache()
        {
            return SpellingService.LocallyAvailableStatus(this.wordText, this.fontName);
        }

        /// <summary>
        /// verifies if the current Nepali word is spell correctly or not.
        /// </summary>
        /// <returns>true if the current work is spell correctly.</returns>
        public Boolean IsCorrectWord()
        {
            return SpellingService.IsCorrectWord(this.wordText, this.fontName);
        }

        /// <summary>
        /// Returns the range of current Nepali Word.
        /// </summary>
        /// <returns>The range of Nepali Word.</returns>
        public Range WordReferenceRange()
        {
            /* *
             * Verifies if the word is followed by special character or white space, it is
             * important for over doing underline.
             * */            
            if (String.IsNullOrWhiteSpace(this.wordReference.Characters.Last.Text)) this.wordReference.SetRange(this.Start, this.End-1);
            else this.wordReference.SetRange(this.Start, this.End);            
            return this.wordReference;
        }

        /// <summary>
        /// Returns the reference word for the current Nepali word.
        /// </summary>
        /// <returns>A reference word.</returns>
        public Range ReferenceRange()
        {
            return this.wordReference;
        }

        /// <summary>
        /// Builds and returns an Empty Nepali word, useful for Empty word representation.
        /// </summary>
        /// <returns>An Empty Nepali word.</returns>                             
        public static NepaliWord Empty()
        {
            //Range emptyRange = Globals.ThisAddIn.nepaliSpellApp.ActiveDocument.Range((Object)0, (Object)0);
            return new NepaliWord(0, 0, TextUtility.EMPTY, null);
        }

        /// <summary>
        /// Receives a list of Nepali words and verifies them either from the local cache or 
        /// from the remote server. Finally it returns the list of words with it verified
        /// status.
        /// </summary>
        /// <param name="nepaliWords"></param>
        /// <returns>listOfVerifiedWords</returns>
        public static List<Tuple<NepaliWord, Boolean>> VerifySpellingErrors(List<NepaliWord> nepaliWords)
        {
             /* *
             * 1. Group words based on index, required to preserve word ordering.
             * 2. Group words based on remote and local availabilities, the rest words are to be sent in
             * remote server for verification.
             * */
            Dictionary<int, List<Tuple<NepaliWord, int>>> groupedWords = Enumerable
                .Range(start: 0, count: nepaliWords.Count)
                .Zip(nepaliWords, (n, value) => new Tuple<NepaliWord, int>(value, n))
                .GroupBy(item => item.Item1.VerifyFromLocalCache())
                .ToDictionary(group => group.Key, group => group.ToList());
            
            /* *
             * Final accumulation and spelling result, whether the given word is either wrong or right.
             * */
            List<Tuple<Tuple<NepaliWord, int>, Boolean>> accumulatedWords = groupedWords.Select(entry =>
            {
                List<Tuple<Tuple<NepaliWord, int>, Boolean>> verifiedWords = new List<Tuple<Tuple<NepaliWord, int>, bool>>();
                switch (entry.Key)
                {
                    case TextUtility.NOT_PRESENT_LOCALLY:
                        List<SayakSuggestion> suggestions = VerifyRemotelyAndUpdateCache(entry.Value);                        
                        verifiedWords = Enumerable
                        .Range(start: 0, count: entry.Value.Count)
                        .Zip(entry.Value, (i, wordWithIndex) => new Tuple<Tuple<NepaliWord, int>, Boolean>(wordWithIndex, suggestions[i].IsCorrectWord()))
                        .ToList();
                        break;

                    case TextUtility.PRESENT_LOCALLY_CORRECT:
                        verifiedWords = FormatLocallyVerifiedWords(entry.Value, true);
                        break;

                    case TextUtility.PRESENT_LOCALLY_WRONG:
                        verifiedWords = FormatLocallyVerifiedWords(entry.Value, false);
                        break;
                }
                return verifiedWords ;
            }).SelectMany(entry => entry).ToList();

            /* *
             * Sort by its original index(ordering) and return the word with its correct status.
             * */
            return accumulatedWords
                .OrderBy(entry => entry.Item1.Item2)
                .Select(entry => new Tuple<NepaliWord, Boolean>(entry.Item1.Item1, entry.Item2))
                .ToList();
        }

        /// <summary>
        /// Format the locally verified word for the correctness of spelling.
        /// </summary>
        /// <param name="indexedWords"></param>
        /// <param name="status"></param>
        /// <returns>formattedListOfTuple</returns>
        private static List<Tuple<Tuple<NepaliWord, int>, Boolean>> FormatLocallyVerifiedWords(List<Tuple<NepaliWord, int>> indexedWords, Boolean status = false)
        {
            return Enumerable
                .Range(start: 0, count: indexedWords.Count)
                .Zip(indexedWords, (i, wordWithIndex) => new Tuple<Tuple<NepaliWord, int>, Boolean>(wordWithIndex, status))
                .ToList();
        }

        /// <summary>
        /// Perform remote verification and update the local cache.
        /// </summary>
        /// <param name="indexedWords"></param>
        /// <returns>listOfSayakSuggestions</returns>
        private static List<SayakSuggestion> VerifyRemotelyAndUpdateCache(List<Tuple<NepaliWord, int>> indexedWords)
        {
            List<SayakSuggestion> suggestions = RemoteIO
                .SayakSuggestionOf(indexedWords.Select(word => word.Item1.wordText)
                .ToList());
            /* *
             * Update suggestions in local cache.
             * */
            suggestions.ForEach(suggestedWord => {
                if (suggestedWord.HasSuggestions())
                {
                    var wordSuggestion = suggestedWord.ToWordSuggestion();//Decode response words.
                    SpellingService.AddToLocalSuggestions(wordSuggestion.wordText, wordSuggestion.fontName, suggestedWord.ToWordSuggestion());
                }
            });

            return suggestions;
        }

        /// <summary>
        /// To String representation of the nepali word.
        /// </summary>
        /// <returns></returns>
        public override string ToString()
        {
            return new JavaScriptSerializer().Serialize(this);
        }

    }
}
