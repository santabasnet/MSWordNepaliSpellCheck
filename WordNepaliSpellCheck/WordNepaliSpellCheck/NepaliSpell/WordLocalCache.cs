using System;
using System.Collections;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace WordNepaliSpellCheck.NepaliSpell
{
    class WordLocalCache
    {
        /// <summary>
        /// The suggestions cache for the current spelling check.
        /// </summary>
        public Dictionary<String, WordSuggestion> wordSuggestions;

        /// <summary>
        /// Some words are ignored in the current spelling context.s
        /// </summary>
        public HashSet<String> ignoredWords;

        /// <summary>
        /// A default constructor that initializes the local cache for local suggestions.
        /// </summary>
        public WordLocalCache()
        {
            this.wordSuggestions = new Dictionary<String, WordSuggestion>();
            this.ignoredWords = new HashSet<string>();
        }

        /// <summary>
        /// Returns the number of entries made in local cache.
        /// </summary>
        /// <returns>noOfLocalEntries</returns>
        public int GetEntryCount()
        {
            return this.wordSuggestions.Count();
        }

        /// <summary>
        /// Adds the current suggestions to local cache for the better performance.
        /// </summary>
        /// <param name="fontName"></param>
        /// <param name="wordText"></param>
        /// <param name="wordSuggestion"></param>
        public void AddToLocalSuggestions(String wordText, String fontName, WordSuggestion wordSuggestion)
        {
            var wordKey = TextUtility.BuildWordKey(fontName, wordText);
            if(!this.wordSuggestions.ContainsKey(wordKey))
                this.wordSuggestions.Add(wordKey, wordSuggestion);
        }

        /// <summary>
        /// Updates the ignored words cache by adding the given parameters.
        /// </summary>
        /// <param name="fontName"></param>
        /// <param name="wordText"></param>
        public void AddIgnoredWord(String fontName, String wordText)
        {
            this.ignoredWords.Add(TextUtility.BuildWordKey(fontName, wordText));
        }

        /// <summary>
        /// Removes the entry from local suggestions, useful to ignored words.
        /// </summary>
        /// <param name="fontName"></param>
        /// <param name="wordText"></param>
        public void RemoveFromLocalSuggestions(String fontName, String wordText)
        {
            this.wordSuggestions.Remove(TextUtility.BuildWordKey(fontName, wordText));
        }

        /// <summary>
        /// Verifies if the current word is already spell checked or not.
        /// </summary>
        /// <param name="wordText"></param>
        /// <param name="fontName"></param>
        /// <returns> </returns>
        public Boolean IsAlreadyChecked(String wordText, String fontName)
        {
            return IsIgnoredWord(wordText, fontName) || IsLocallyPresent(wordText, fontName);
        }

        /// <summary>
        /// The method is used to check whether the given word is ignored in current context
        /// or not.
        /// </summary>
        /// <param name="fontName"></param>
        /// <param name="wordText"></param>
        /// <returns>True for the given word is ignored to check the current spelling.</returns>
        public bool IsIgnoredWord(String wordText, String fontName)
        {
            return this.ignoredWords.Contains(TextUtility.BuildWordKey(fontName, wordText));
        }

        /// <summary>
        /// Verifies the given word in text with given font name is already checked and put
        /// in the local cache or not.
        /// </summary>
        /// <param name="fontName"></param>
        /// <param name="wordText"></param>
        /// <returns>true if the word is already spell checked.</returns>
        public Boolean IsLocallyPresent(String wordText, String fontName)
        {
            return this.wordSuggestions.ContainsKey(TextUtility.BuildWordKey(fontName, wordText));
        }

        /// <summary>
        /// Verifies the word is correctly spelled or not from the local cache.
        /// </summary>
        /// <param name="wordText"></param>
        /// <param name="fontName"></param>
        /// <returns>true indicates that the given word is mis-spelled one.</returns>
        public int VerifyLocally(String wordText, String fontName)
        {
            WordSuggestion localSuggestion = GetLocalSuggestion(wordText, fontName);
            if (localSuggestion.IsEmpty()) return TextUtility.NOT_PRESENT_LOCALLY;
            else return localSuggestion.suggestion.wrongWord ? TextUtility.PRESENT_LOCALLY_WRONG : TextUtility.PRESENT_LOCALLY_CORRECT;
        }

        /// <summary>
        /// Returns the the list of suggestions from local cache if it is already stored in the local cache.
        /// </summary>
        /// <param name="fontName"></param>
        /// <param name="wordText"></param>
        /// <returns>wordSuggestions in the list form.</returns>
        public WordSuggestion GetLocalSuggestion(String wordText, String fontName)
        {
            if (!IsLocallyPresent(wordText, fontName)) return WordSuggestion.Empty();
            else return this.wordSuggestions[TextUtility.BuildWordKey(fontName, wordText)];
        }

    }
}
