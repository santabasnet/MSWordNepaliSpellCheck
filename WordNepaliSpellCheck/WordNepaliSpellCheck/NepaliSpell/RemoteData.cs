using System;
using System.Collections.Generic;
using System.Linq;
using System.Web.Script.Serialization;

namespace WordNepaliSpellCheck.NepaliSpell
{
    /// <summary>
    /// Defines class to represent the remote server data to be sent.
    /// </summary>
    class RemoteData
    {           
        /// <summary>
        /// Action name: whether to spell check, add new word or remove word.
        /// </summary>
        public String actionName { get; set; }

        /// <summary>
        /// Associated font name.
        /// </summary>
        public String fontName { get; set; }

        /// <summary>
        /// List of words to be sent to spell check.
        /// </summary>
        public List<String> wordText { get; set; }

        /// <summary>
        /// Default constructor to the remote data instance.
        /// </summary>
        /// <param name="actionName"></param>
        /// <param name="fontName"></param>
        /// <param name="wordText"></param>
        public RemoteData(String actionName, String fontName, List<String> wordText)
        {
            this.actionName = actionName;
            this.fontName = fontName;
            this.wordText = wordText;
        }

        /// <summary>
        /// Returns the json serialized data of the current instance.
        /// </summary>
        /// <returns>jsonData</returns>
        override
        public String ToString()
        {
           return new JavaScriptSerializer().Serialize(this);
        }

        public static RemoteData of(String actionName, String fontName, List<String> wordText)
        {
            return new RemoteData(actionName, fontName, wordText);
        }

    }

    class FontWord
    {
        public String wordText { get; set; }
        public String fontName { get; set; } 

        public FontWord(String wordText, String fontName)
        {
            this.wordText = wordText;
            this.fontName = fontName;
        }

        override
        public String ToString()
        {
            return new JavaScriptSerializer().Serialize(this);
        }

        public FontWord WithEncoded()
        {
            return FontWord.of(TextUtility.UrlEncode(this.wordText), this.fontName);
        }

        public FontWord WithDecoded()
        {
            return FontWord.of(TextUtility.UrlDecode(this.wordText), this.fontName);
        }

        public Boolean IsEmpty()
        {
            return String.IsNullOrWhiteSpace(this.wordText);
        }

        public static FontWord of(String wordText, String fontName)
        {
            return new FontWord(wordText, fontName);
        }

        public static FontWord ofUnicode(String wordText)
        {
            return new FontWord(wordText, TextUtility.UNICODE);
        }

        public static FontWord empty()
        {
            return new FontWord(TextUtility.EMPTY, TextUtility.UNICODE);
        }
    }

    class SpellParams
    {
        public String lang { get; set; }
        public List<FontWord> words { get; set; }

        public SpellParams(String lang, List<FontWord> words)
        {
            this.lang = lang;
            this.words = words;
        }

        override
        public String ToString()
        {
            return new JavaScriptSerializer().Serialize(this);
        }

        public static SpellParams of(String lang, List<FontWord> words)
        {
            return new SpellParams(lang, words);
        }

        public static SpellParams ofDefaultLanguage(List<FontWord> words)
        {
            return SpellParams.of(TextUtility.DEFAULT_LANGUAGE, words);
        }
    }

    class SpellingData
    {
        public String requestMethod { get; set; }
        public SpellParams spellParams { get; set; }

        public String token { get; set; }

        //public String wordPluginId { get; set; }

        public SpellingData(String requestMethod, SpellParams spellParams)
        {
            this.requestMethod = requestMethod;
            this.spellParams = spellParams;
            this.token = TextUtility.EMPTY;
            //this.wordPluginId = COMUtility.APP_ID;
        }

        override
        public String ToString()
        {
            return new JavaScriptSerializer().Serialize(this);
        }

        public static SpellingData of(String requestMethod, SpellParams spellParams)
        {
            return new SpellingData(requestMethod, spellParams);
        }

        public static SpellingData ofSpellCheck(SpellParams spellParams)
        {
            return SpellingData.of(TextUtility.ACTION_SPELL_CHECK, spellParams);
        }
    }

    class RemoteServiceData
    {
        /// <summary>
        /// Dictionary representation for the client data.
        /// </summary>
        public Dictionary<String, String> clientInput { get; set; }

        public RemoteServiceData(Dictionary<String, String> clientInput)
        {
            this.clientInput = clientInput;
        }

        override
        public String ToString()
        {
            return new JavaScriptSerializer().Serialize(clientInput);
        }

        public static RemoteServiceData of(Dictionary<String, String> clientInput)
        {
            return new RemoteServiceData(clientInput);
        }

        public static SpellingData ofSpellingSuggestion(String fontName, List<String> wordsText)
        {
            List<FontWord> words = wordsText
                .Select(word => FontWord.of(word, fontName).WithEncoded())
                .ToList();
            return SpellingData.ofSpellCheck(SpellParams.ofDefaultLanguage(words));
        }

        public static SpellingData ofUnicodeSpellingSuggestion(List<String> wordsText)
        {
            return RemoteServiceData.ofSpellingSuggestion(TextUtility.UNICODE, wordsText);
        }

        public static RemoteServiceData ofServiceUnavailability(String appId)
        {
            var clientInput = new Dictionary<string, string>() {
                {TextUtility.ACTION_NAME, TextUtility.ACTION_CHECK_SERVICE_AVAILABILITY},
                {TextUtility.APP_ID, appId}
            };
            return RemoteServiceData.of(clientInput);
        }
    }
}
