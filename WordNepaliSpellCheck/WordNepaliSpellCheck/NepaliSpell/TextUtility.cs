using Microsoft.IdentityModel.Tokens;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Web;

namespace WordNepaliSpellCheck.NepaliSpell
{
    class TextUtility
    {
        /// <summary>
        /// Empty string representation.
        /// </summary>
        public static readonly String EMPTY = "";

        /// <summary>
        /// Empty list representation.
        /// </summary>
        public static readonly List<String> EMPTY_LIST = new List<String>();

        /// <summary>
        /// Unicode String literal.
        /// </summary>
        public static readonly String UNICODE = "UNICODE";
        public static readonly string ACTION_NAME = "actionName";
        public static readonly string FONT_NAME = "fontName";
        public static readonly string WORD_TEXT = "wordText";
        public static readonly string USER_ID = "userId";
        public static readonly string APP_ID = "appId";
        public static readonly string DATA = "data";

        /// <summary>
        /// Spell actions literals to show in button captions.
        /// </summary>
        public static readonly string SPELL_ACTION_ON = "Spelling Action: ON";
        public static readonly string SPELL_ACTION_OFF = "Spelling Action: OFF";

        /// <summary>
        /// A white space text.
        /// </summary>
        public static readonly String WHITE_SPACE = " ";

        /// <summary>
        /// Token separator delimiter.
        /// </summary>
        public static readonly String TOKEN_SEPARATOR = "-";

        /// <summary>
        /// Action name to check spell for a word of remote request.
        /// </summary>
        public static readonly String ACTION_SPELL_CHECK = "spellcheck";

        /// <summary>
        /// Action name to add new word to the remote spelling dictionary.
        /// </summary>
        public static readonly String ACTION_ADD_WORD = "addword";

        /// <summary>
        /// Action name to delete an existing word from the remote spelling dictionary.
        /// </summary>
        public static readonly String ACTOIN_DELETE_WORD = "deleteword";

        /// <summary>
        /// Action name to check if the spelling service availability. 
        /// </summary>
        public static readonly String ACTION_CHECK_SERVICE_AVAILABILITY = "checkserviceavailability";

        /// <summary>
        /// Request method of server field.
        /// </summary>
        public static readonly string REQUEST_METHOD = "requestMethod";

        /// <summary>
        /// Spell paramters for server field.
        /// </summary>
        public static readonly string SPELL_PARAMS = "spellParams";

        /// <summary>
        /// Nepali language symbol of default language.
        /// </summary>
        public static readonly string DEFAULT_LANGUAGE = "np";

        /// <summary>
        /// Language field for server.
        /// </summary>
        public static readonly string LANGUAGE = "lang";

        /// <summary>
        /// Words literals.
        /// </summary>
        public static readonly string WORDS = "words";

        public static readonly string DEFAULT_TOKEN = "eyJhbGciOiJIUzI1NiJ9.eyJqdGkiOiI2M2U2MzU5Yi01ZDZmLTQ4NTEtOGZiNS0wZWRmODRiNWMyNWMiLCJpYXQiOjE1MTIyNzkwNjIsImlzcyI6IlNwZWxsQ2hlY2tlciJ9.6-zsgynC9JUV1HfQWfudWom3GmhqgkAetoYhiCIpMVk";

        /// <summary>
        /// List of Nepali Punctuations.
        /// </summary>
        private static readonly List<Char> NEPALI_PUNCTUATIONS = new List<Char>() {
            ',', ';', ':', '?', '!', '"', '—', '-'
        };
        
        /// <summary>
        /// Varibale for spelling status check: Denotes the word is not present in the local cache.
        /// </summary>
        public const int NOT_PRESENT_LOCALLY = 0;

        /// <summary>
        /// Varibale for spelling status check: Denotes the word is present in the local cache
        /// but has wrong spell status.
        /// </summary>
        public const int PRESENT_LOCALLY_WRONG = 1;

        /// <summary>
        /// Varibale for spelling status check: Denotes the word is present in the local cache
        /// but has correct spell status.
        /// </summary>
        public const int PRESENT_LOCALLY_CORRECT = 2; 

        /// <summary>
        /// Negative count reference.
        /// </summary>
        public static readonly object NEGATIVE_COUNT = -1;

        /// <summary>
        /// Positive count reference.
        /// </summary>
        public static readonly object POSITIVE_COUNT = -1;

        /// <summary>
        /// Waiting time to restart the spelling correction job.
        /// </summary>
        public static readonly int WAITING_TIME = 200;

        /// <summary>
        /// 400 Milli-Seconds waiting tick.
        /// </summary>
        public static readonly int WAITING_TICK = 100;

        /// <summary>
        /// Intermediate elasped time between two consecutive paragraphs.
        /// </summary>
        public static readonly int ELAPSED_TIME = 50;

        /// <summary>
        /// This method builds a key representation for the given word text attached 
        /// </summary>
        /// <param name="fontName"></param>
        /// <param name="wordText"></param>
        /// <returns>A key for the word represented for the given arguments.</returns>
        public static String BuildWordKey(String fontName, String wordText)
        {
            return $"{fontName}_{wordText}";
        }

        /// <summary>
        /// Validates whether the given string of single character is non-printable 
        /// or not.
        /// </summary>
        /// <param name="charText"></param>
        /// <returns>true after successful validation.</returns>
        public static Boolean IsNonPrintable(String charText)
        {
            return !String.IsNullOrEmpty(charText) && (Char.IsControl(charText[0]) || Char.IsWhiteSpace(charText[0]));
        }

        /// <summary>
        /// Validates the give word has text inside or not.
        /// </summary>
        /// <param name="wordText"></param>
        /// <returns></returns>
        public static Boolean IsValidWord(String wordText)
        {
            return !String.IsNullOrEmpty(wordText);
        }

        /// <summary>
        /// Validates the given character code lies in the Nepali Unicode range or not.
        /// </summary>
        /// <param name="charCode"></param>
        /// <returns>true if it is a valid Nepali Unicode.</returns>
        public static Boolean IsUnicode(Char charCode)
        {
            return charCode >= 0x900 && charCode <= 0x97F;
        }

        /// <summary>
        /// Convert the caption text to tag item.
        /// </summary>
        /// <param name="captionText"></param>
        /// <returns>A tag representation of the caption word phrase.</returns>
        public static String ConvertTagItem(String captionText)
        {
            var result = captionText.Split(' ').Where(word => !String.IsNullOrWhiteSpace(word)).Select(word => word.ToLower());
            return String.Join("|", new List<String>() { String.Join("_", result), System.Guid.NewGuid().ToString() });
        }

        /// <summary>
        /// Encode the given plain text in Base64.
        /// </summary>
        /// <param name="text"></param>
        /// <returns>encodedText</returns>
        public static String Base64Encode(String text)
        {
            return Base64UrlEncoder.Encode(text);
        }

        /// <summary>
        /// Decode given text from Base64 text.
        /// </summary>
        /// <param name="encodedText"></param>
        /// <returns>decodedText</returns>
        public static String Base64Decode(String encodedText)
        {
            return Base64UrlEncoder.Decode(encodedText);
        }

        public static String UrlEncode(String text)
        {
            return HttpUtility.UrlEncode(text);
        }

        public static String UrlDecode(String text)
        {
            return HttpUtility.UrlDecode(text);
        }

        /// <summary>
        /// Verifies if the given character is Nepali punctuation or not.
        /// </summary>
        /// <param name="ch"></param>
        /// <returns></returns>
        public static Boolean IsNepaliPunctuation(char ch)
        {
            return NEPALI_PUNCTUATIONS.Contains(ch);
        }

    }
}
