using Newtonsoft.Json;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace WordNepaliSpellCheck.NepaliSpell
{
    /// <summary>
    /// Represents the class for spelling server response. The server returns the suggestions
    /// for the list of words. Generally, the response contains a single element.
    /// </summary>
    class SpellingSuggestion
    {
        /// <summary>
        /// Holds the list of suggestions from remote server.
        /// </summary>
        public List<WordSuggestion> wordSuggestionsList { get; set; }

        /// <summary>
        /// Default constructor to represent Empty suggestion list.
        /// </summary>
        public SpellingSuggestion()
        {
            this.wordSuggestionsList = new List<WordSuggestion> { WordSuggestion.Empty() };
        }

        /// <summary>
        /// Verifies if the current spelling suggestion is an Empty instance.
        /// </summary>
        /// <returns>true if the current suggestion has Empty instance.</returns>
        public Boolean IsEmpty()
        {
            return this.wordSuggestionsList.First().IsEmpty();
        }

        /// <summary>
        /// Returns an Empty suggestion for a single word, used for rescue purpose.
        /// </summary>
        /// <returns>An Empty suggstion for a word.</returns>
        public static SpellingSuggestion Empty()
        {
            return new SpellingSuggestion();
        }
    }

    /// <summary>
    /// Represents the class for a single word suggestion.
    /// </summary>
    class WordSuggestion
    {
        /// <summary>
        /// The font name of the word.
        /// </summary>
        public String fontName { get; set; }

        /// <summary>
        /// The text content for a word.
        /// </summary>
        public String wordText { get; set; }

        /// <summary>
        /// The suggestion made by the server. It is used to check whether the word is correct or not too.
        /// </summary>
        public Suggestion suggestion { get; set; }

        /// <summary>
        /// Default constructor for Empty word suggestion. 
        /// </summary>
        public WordSuggestion()
        {
            this.fontName = TextUtility.EMPTY;
            this.wordText = TextUtility.EMPTY;
            this.suggestion = Suggestion.Empty();
        }

        public WordSuggestion(String fontName, String wordText, Suggestion suggestion)
        {
            this.fontName = fontName;
            this.wordText = wordText;
            this.suggestion = suggestion;
        }

        /// <summary>
        /// Verify whether the current suggestion is an Empty instance or not.
        /// </summary>
        /// <returns>true if the current instance is Empty.</returns>
        public Boolean IsEmpty()
        {
            return String.IsNullOrEmpty(this.fontName) || String.IsNullOrEmpty(this.wordText);
        }

        public override String ToString()
        {
            return JsonConvert.SerializeObject(this, Formatting.Indented);
        }

        /// <summary>
        /// Builds an Empty Word suggestion to represent an Empty instance.
        /// </summary>
        /// <returns></returns>
        public static WordSuggestion Empty()
        {
            return new WordSuggestion();
        }

        public static WordSuggestion CorrectWordOf(NepaliWord nepaliWord)
        {
            return new WordSuggestion(nepaliWord.fontName, nepaliWord.wordText, Suggestion.Correct());
        }
        
        public static WordSuggestion of(String fontName, String wordText, Suggestion suggestion)
        {
            return new WordSuggestion(fontName, wordText, suggestion);
        }
    }

    /// <summary>
    /// Represents class for the list of word suggestions.
    /// </summary>
    class Suggestion
    {
        /// <summary>
        /// Flag for the given word, whether it is wrong word or not.
        /// </summary>
        public Boolean wrongWord { get; set; }

        /// <summary>
        /// Holds the list of suggested words.
        /// </summary>
        public List<String> wordSuggestions { get; set; }

        /// <summary>
        /// Default constructor to represent Empty object instance.
        /// </summary>
        public Suggestion()
        {
            this.wrongWord = true;
            this.wordSuggestions = new List<String> { };
        }

        public Suggestion(Boolean wrongWord, List<String> wordSuggestions)
        {
            this.wrongWord = wrongWord;
            this.wordSuggestions = wordSuggestions;
        }

        /// <summary>
        /// Builds an Empty Suggestion object.
        /// </summary>
        /// <returns></returns>
        public static Suggestion Empty()
        {
            return new Suggestion();
        }

        public static Suggestion Correct()
        {
            return new Suggestion(false, new List<String>());
        } 

        public static Suggestion Of(Boolean wrongWord, List<String> wordSuggestions)
        {
            return new Suggestion(wrongWord, wordSuggestions);
        }
    }

    class SayakSuggestion
    {
        public FontWord wordInfo { get; set; }
        public Boolean wrongWord { get; set; }

        public List<String> suggestionsList { get; set; }

        private List<String> decodedSuggestions()
        {
            return suggestionsList.Select(text => TextUtility.UrlDecode(text)).ToList();
        }

        public SayakSuggestion(FontWord wordInfo, Boolean wrongWord, List<String> suggestionsList)
        {
            this.wordInfo = wordInfo;
            this.wrongWord = wrongWord;
            this.suggestionsList = suggestionsList;
        }

        public Boolean IsWrongWord()
        {
            return this.wrongWord == true;
        }

        public Boolean IsCorrectWord()
        {
            return this.wrongWord == false;
        }

        public SayakSuggestion WithDecodedWords()
        {
            return SayakSuggestion.of(wordInfo.WithDecoded(), this.wrongWord, this.decodedSuggestions());
        }

        public WordSuggestion ToWordSuggestion()
        {
            var decodedSuggestion = this.WithDecodedWords();
            return WordSuggestion.of(decodedSuggestion.wordInfo.fontName, decodedSuggestion.wordInfo.wordText, decodedSuggestion.ToSuggestion());
        }

        private Suggestion ToSuggestion()
        {
            return Suggestion.Of(this.wrongWord, this.suggestionsList);
        }

        public Boolean IsEmpty()
        {
            return this.wordInfo.IsEmpty();
        } 
        
        public Boolean HasSuggestions()
        {
            return !IsEmpty();
        }              

        public List<String> SuggestedWords()
        {
            return this.suggestionsList;
        }

        public override String ToString()
        {
            return JsonConvert.SerializeObject(this, Formatting.Indented);
        }

        public static SayakSuggestion of(FontWord wordInfo, Boolean wrongWord, List<String> suggestionsList)
        {
            return new SayakSuggestion(wordInfo, wrongWord, suggestionsList);
        }

        public static SayakSuggestion Empty()
        {
            return SayakSuggestion.of(FontWord.empty(), true, new List<String>() );
        }

    }

    class WordDumpResponse
    {
        public Dictionary<String, String> data { get; set; }

        public WordDumpResponse(Dictionary<String, String> data)
        {
            this.data = data;
        }

        public  String GetMessage()
        {
            return data["message"];
        }

        public String GetStatus()
        {
            return data["status"];
        }

        public Boolean IsSuccess()
        {
            return data["status"] == "success";
        }

        public override String ToString()
        {
            return JsonConvert.SerializeObject(this.data, Formatting.Indented);
        }

        public static WordDumpResponse Empty()
        {
            return new WordDumpResponse(new Dictionary<String, String>());
        }

        public static WordDumpResponse buildWith(String jsonResponse)
        {
            var data = JsonConvert.DeserializeObject<Dictionary<String, String>>(jsonResponse);
            return new WordDumpResponse(data);
        }
    }
}
