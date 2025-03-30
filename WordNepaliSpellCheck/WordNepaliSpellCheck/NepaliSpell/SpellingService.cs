using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace WordNepaliSpellCheck.NepaliSpell
{
    class SpellingService
    {
        /// <summary>
        /// Demo Spelling Data.
        /// </summary>
        private static readonly DemoData demoData = new DemoData();

        /// <summary>
        /// Local cache for spelling check. It defines the local storage of already verified words.
        /// </summary>
        private static WordLocalCache localSuggestionCache = new WordLocalCache();

        /// <summary>
        /// List of Nepali fonts that are used in the system.
        /// </summary>
        private static readonly List<String> nepaliFontNames = new List<String> {
            //TextUtility.UNICODE
            "KANTIPUR", "PREETI", "PCS NEPALI", "HIMALB", "AAKRITI", "AALEKH", "GANESS", "NAVJEEVAN", "UNICODE"
        };

        /// <summary>
        /// Verifies the given word is spelled correctly or not. It also adds the current suggestion
        /// to the local cache.
        /// </summary>
        /// <param name="wordText"></param>
        /// <param name="fontName"></param>
        /// <returns>true if the given word is wrong in spelling.</returns>
        private static Boolean VerifyRemotely(String wordText, String fontName)
        {
            if (RemoteIO.IsLocalServerSource())
                return VerifyFromLocalServer(wordText, fontName);
            else
                return VerifyFromRemoteServer(wordText, fontName);
        }

        /// <summary>
        /// Old code: that has spelling suggestion object for JSON serialization.
        /// </summary>
        /// <param name="wordText"></param>
        /// <param name="fontName"></param>
        /// <returns>true if the word in the given font is available in local store.</returns>
        private static Boolean VerifyFromLocalServer(String wordText, String fontName)
        {
            SpellingSuggestion spellingSuggestion = RemoteIO.GetRemoteSuggestion(wordText, fontName);
            if (spellingSuggestion.IsEmpty()) return true;
            else
            {
                WordSuggestion wordSuggestion = spellingSuggestion.wordSuggestionsList.First();
                localSuggestionCache.AddToLocalSuggestions(wordText, fontName, wordSuggestion);
                return !wordSuggestion.suggestion.wrongWord;
            }
        }

        /// <summary>
        /// New code: that has Sayak suggestion object for JSON serialization, receives the suggestion
        /// from Sayak web site.
        /// </summary>
        /// <param name="wordText"></param>
        /// <param name="fontName"></param>
        /// <returns>true if the word in the given font is available in sayak store.</returns>
        private static Boolean VerifyFromRemoteServer(String wordText, String fontName)
        {
            Tuple<SayakResponse, SayakSuggestion> sayakSuggestionResponse = RemoteIO.SayakSuggestionOf(wordText, fontName);
            if (sayakSuggestionResponse.Item2.HasSuggestions())
            {
                localSuggestionCache.AddToLocalSuggestions(wordText, fontName, sayakSuggestionResponse.Item2.ToWordSuggestion());
                return sayakSuggestionResponse.Item2.IsCorrectWord();
            }
            else
            {
                Tuple<String, String> remoteURLMessage = sayakSuggestionResponse.Item1.BuildRemoteURLMessage();
                DialogResult dialogResult = MessageBox.Show(remoteURLMessage.Item2, ResponseLiterals.SAYAK_SERVICE_NAME, MessageBoxButtons.YesNo, MessageBoxIcon.Asterisk);
                if (dialogResult == DialogResult.Yes) System.Diagnostics.Process.Start(remoteURLMessage.Item1);
                return true;
            }

            /*if (sayakSuggestionResponse.Item2.IsEmpty()) return true;
            else
            {
                localSuggestionCache.AddToLocalSuggestions(wordText, fontName, sayakSuggestionResponse.Item2.ToWordSuggestion());
                return sayakSuggestionResponse.Item2.IsCorrectWord();
            }*/
        }

        /// <summary>
        /// Inserts the given Nepali word in the local cache as a correct one.
        /// [It assumes that the suggestion made by the remote server have all the correct
        /// word entries.]
        /// </summary>
        /// <param name="nepaliWord"></param>
        public static void AddCorrectWordToCache(NepaliWord nepaliWord)
        {
            localSuggestionCache.AddToLocalSuggestions(nepaliWord.wordText, nepaliWord.fontName, WordSuggestion.CorrectWordOf(nepaliWord));
        }

        /// <summary>
        /// Verifies if the current word is present in the local cache with mis-spelled status
        /// or not.
        /// </summary>
        /// <param name="wordText"></param>
        /// <param name="fontName"></param>
        /// <returns>true if the current is wrongly spelled.</returns>
        public static Boolean IsLocallyPresentError(String wordText, String fontName)
        {
            int localStatus = localSuggestionCache.VerifyLocally(wordText, fontName);
            return localStatus == TextUtility.PRESENT_LOCALLY_WRONG;
        }

        /// <summary>
        /// Returns the suggestion for the given word. It always suggest from the local cache.
        /// </summary>
        /// <param name="wordText"></param>
        /// <param name="fontName"></param>
        /// <returns>A list of suggested words.</returns>
        public static List<String> GetSuggestions(String wordText, String fontName = "UNICODE")
        {
            return localSuggestionCache.GetLocalSuggestion(wordText, fontName.ToUpper()).suggestion.wordSuggestions;
        }

        /// <summary>
        /// Returns the suggestion lif or the given Nepali word. 
        /// </summary>
        /// <param name="nepaliWord"></param>
        /// <returns>A list of suggested words.</returns>
        public static List<String> GetSuggestions(NepaliWord nepaliWord)
        {
            return GetSuggestions(nepaliWord.wordText, nepaliWord.fontName);
        }

        /// <summary>
        /// Validates if the given word text with the font name is a correct word or not.
        /// Here, the correct word means, either it is ignored word, or locally present
        /// as already checked word, or verified a wrong word from the remote service.
        /// </summary>
        /// <param name="wordText"></param>
        /// <param name="fontName"></param>
        /// <returns></returns>
        public static Boolean IsCorrectWord(String wordText, String fontName = "UNICODE")
        {
            if (localSuggestionCache.IsIgnoredWord(wordText, fontName)) return true;
            int localStatus = localSuggestionCache.VerifyLocally(wordText, fontName);
            switch (localStatus)
            {
                case TextUtility.NOT_PRESENT_LOCALLY:
                    return VerifyRemotely(wordText, fontName);
                case TextUtility.PRESENT_LOCALLY_CORRECT:
                    return true;
                case TextUtility.PRESENT_LOCALLY_WRONG:
                    return false;
                default:
                    return false;
            }
        }

        /// <summary>
        /// Verifies if the given word information is present in the local cache or not by their status.
        /// </summary>
        /// <param name="wordText"></param>
        /// <param name="fontName"></param>
        /// <returns>localStatus</returns>
        public static int LocallyAvailableStatus(String wordText, String fontName = "UNICODE")
        {
            if (localSuggestionCache.IsIgnoredWord(wordText, fontName)) return TextUtility.PRESENT_LOCALLY_CORRECT;
            return localSuggestionCache.VerifyLocally(wordText, fontName);
        }

        /// <summary>
        /// Checks whether the given font is Nepali or not.
        /// </summary>
        /// <param name="fontName"></param>
        /// <returns>true if the given font name is Nepali.</returns>
        public static Boolean IsNepaliFont(String fontName)
        {
            return SpellingService.nepaliFontNames.Contains(fontName.ToUpper());
        }

        /// <summary>
        /// Returns all the word entries count.
        /// </summary>
        /// <returns>totalCount</returns>
        public static int TotalEntries()
        {
            return localSuggestionCache.GetEntryCount();
        }

        public static void AddIgnoreWord(String wordText, String fontName = "UNICODE")
        {
            localSuggestionCache.AddIgnoredWord(fontName, wordText);
        }

        public static void RemoveIgnoredWordFromSuggestions(String wordText, String fontName = "UNICODE")
        {
            localSuggestionCache.RemoveFromLocalSuggestions(fontName, wordText);
        }

        public static void AddToLocalSuggestions(String wordText, String fontName, WordSuggestion wordSuggestion)
        {
            localSuggestionCache.AddToLocalSuggestions(wordText, fontName, wordSuggestion);
        }

        public static String ListIgnoredWords()
        {
            return String.Join(", ", localSuggestionCache.ignoredWords);
        }

    }
}
