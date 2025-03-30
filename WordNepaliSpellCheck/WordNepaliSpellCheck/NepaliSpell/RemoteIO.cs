using System;
using System.Web;
using System.Net.Http;
using System.Threading.Tasks;
using System.Collections.Generic;
using System.Web.Script.Serialization;
using System.Text;
using Newtonsoft.Json;
using System.Linq;
using System.Diagnostics;
using Newtonsoft.Json.Linq;

namespace WordNepaliSpellCheck.NepaliSpell
{
    class RemoteIO
    {
        /// <summary>
        /// Remote server url for spelling check.
        /// </summary>
        private static readonly String SERVER_URL = SpellSettings.Default.remoteURL;

        /// <summary>
        /// Remote server to dump the user suggested words.
        /// </summary>
        private static readonly String USER_SUGGESTION_URL = SpellSettings.Default.userSuggestionURL;

        /// <summary>
        /// Server source, whether it is local or remote.
        /// </summary>
        public static readonly String SERVER_SOURCE = new Uri(SERVER_URL).Host.ToUpper();

        /// <summary>
        /// Default timeout for remote connection.
        /// </summary>
        private static readonly int DEFAULT_TIMEOUT = 3000;

        /// <summary>
        /// Spelling client to make remote request.
        /// </summary>
        private static readonly HttpClient spellClient = initClient();

        /// <summary>
        /// Sends asynchronous get request to spelling server.
        /// </summary>
        /// <param name="urlData"></param>
        /// <returns>A task of spelling response.</returns>
        private static async Task<String> SendRemoteRequest(String urlData)
        {
            HttpResponseMessage spellResponse = await spellClient.GetAsync(urlData);
            return await spellResponse.Content.ReadAsStringAsync();
        }

        /// <summary>
        /// Sends asynchronuus post reqeust to the spelling server.
        /// </summary>
        /// <param name="jsonBody"></param>
        /// <returns></returns>
        private static async Task<String> SendMultipartPostRequest(String jsonBody, String serverURL)
        {
            //var multipartData = new MultipartFormDataContent();          
            //multipartData.Add(new StringContent(jsonBody), String.Format("\"{0}\"", TextUtility.DATA));
            //HttpResponseMessage spellResponse = await spellClient.PostAsync(serverURL, multipartData);

            StringContent content = new StringContent(jsonBody, Encoding.UTF8, "application/json");
            HttpResponseMessage spellResponse = await spellClient.PostAsync(serverURL, content);


            return await spellResponse.Content.ReadAsStringAsync();
        }

        /// <summary>
        /// Formats the server data for a word before remote request.
        /// </summary>
        /// <param name="actionName"></param>
        /// <param name="fontName"></param>
        /// <param name="wordText"></param>
        /// <returns>A spelling suggested response.</returns>
        private static String FormatServerData(String actionName, String fontName, String wordText)
        {
            return FormatServerData(actionName, fontName, new List<String> { wordText });
        }

        /// <summary>
        /// Formats the server data for the list of words before remote request.
        /// </summary>
        /// <param name="actionName"></param>
        /// <param name="fontName"></param>
        /// <param name="wordText"></param>
        /// <returns>A list of spelling suggested responses.</returns>
        private static String FormatServerData(String actionName, String fontName, List<String> wordText)
        {
            String remoteData = RemoteData.of(actionName, fontName, wordText).ToString();
            return SERVER_URL + HttpUtility.UrlEncode(remoteData);
        }

        /// <summary>
        /// Formats the service check url.
        /// </summary>
        /// <returns></returns>
        private static String FormatServiceElibility()
        {
            // Generate AppId here.
            String appId = COMUtility.GenerateAppId();
            //return SERVER_URL + HttpUtility.UrlEncode(RemoteServiceData.ofServiceUnavailability(appId).ToString());
            return RemoteServiceData.ofServiceUnavailability(appId).ToString();
        }

        /// <summary>
        /// Make a get request to spell suggestions.
        /// It returns whether the given request has mis-spelled information or not.
        /// If the word is mis-spelled, it returns with suggestions.
        /// It also decodes the URL encoded data from the server.
        /// </summary>
        /// <param name="urlData"></param>
        /// <returns>The list of suggested words with mis-spelled information.</returns>
        private static String SuggestSpellWithGetRequest(String urlData)
        {
            try {
                Task<String> spellingRequestTask = SendRemoteRequest(urlData);
                return HttpUtility.UrlDecode(spellingRequestTask.Result);
            }
            catch {
                return String.Empty;
            }
        }

        /// <summary>
        /// Make a POST request to get spelling suggestions.
        /// </summary>
        /// <param name="jsonBody"></param>
        /// <returns></returns>
        public static String SuggestSpellWithPostRequest(String jsonBody)
        {
            try {
                Task<String> spellingRequestTask = SendMultipartPostRequest(jsonBody, SERVER_URL);
                return spellingRequestTask.Result;
            }
            catch {
                return String.Empty;
            }
        }

        /// <summary>
        /// Make a POST request to dump the word from user part.
        /// </summary>
        /// <param name="jsonBody"></param>
        /// <returns>jsonResponse</returns>
        public static String AddUserSuggestedWord(String jsonBody)
        {
            try {
                return SendMultipartPostRequest(jsonBody, USER_SUGGESTION_URL).Result;
            }
            catch {
                return String.Empty;
            }
        }

        /// <summary>
        /// Formats the given request to remote server and extracts the response. It deserializes
        /// the spelling response to local object SpellingSuggestion. It returns an Empty
        /// instance if it fails to do remote call and deserialization.
        /// </summary>
        /// <param name="wordText"></param>
        /// <param name="fontName"></param>
        /// <returns>An instance of SpellingSuggestion for a Nepali word.</returns>
        public static SpellingSuggestion GetRemoteSuggestion(String wordText, String fontName)
        {
            String dataToServer = FormatServerData(TextUtility.ACTION_SPELL_CHECK, fontName, wordText);
            String spellingResult = RemoteIO.SuggestSpellWithGetRequest(dataToServer);
            if (String.IsNullOrEmpty(spellingResult)) return SpellingSuggestion.Empty();
            else return new JavaScriptSerializer().Deserialize<SpellingSuggestion>(spellingResult);


            ///Needs to change the code searializer.

        }

        /// <summary>
        /// Make a request to update a new word as suggested from users.
        /// </summary>
        /// <param name="wordText"></param>
        /// <returns>wordDumpResponse</returns>
        public static WordDumpResponse SuggestedWordFromClient(String wordText, String fontName = "UNICODE")
        {
            var serverData = RemoteIO.buildAddNewWordData(wordText, fontName);
            String serverResponse = RemoteIO.AddUserSuggestedWord(serverData);
            if (String.IsNullOrEmpty(serverResponse)) return WordDumpResponse.Empty();
            else return WordDumpResponse.buildWith(serverResponse);
        }

        /// <summary>
        /// Get suggestions of the list of words specified with same font name.
        /// </summary>
        /// <param name="words"></param>
        /// <param name="fontName"></param>
        /// <returns>tupleOfSayakResponse with Suggestions.</returns>
        public static Tuple<SayakResponse, List<SayakSuggestion>> SayakSuggestionOf(List<String> words, String fontName)
        {
            String remoteData = RemoteServiceData.ofSpellingSuggestion(fontName, words).ToString();
            String spellingResult = RemoteIO.SuggestSpellWithPostRequest(remoteData);
            SayakResponse sayakResponse = SayakResponse.Empty();
            List<SayakSuggestion> sayakSuggestions = new List<SayakSuggestion>() { SayakSuggestion.Empty() };

            JToken jsonToken = JToken.Parse(spellingResult);
            switch (jsonToken.Type)
            {
                case JTokenType.Array:
                    sayakSuggestions = RemoteIO.buildSayakSuggestions(jsonToken);
                    break;
                case JTokenType.Object:
                    sayakResponse = SayakResponse.buildWith(spellingResult);
                    break;
            }             
            return new Tuple<SayakResponse, List<SayakSuggestion>>(sayakResponse, sayakSuggestions);
        }

        /// <summary>
        /// Get suggestions of the list of words specified with UNICODE font name.
        /// </summary>
        /// <param name="words"></param>
        /// <returns>List of Sayak suggestions.</returns>
        public static List<SayakSuggestion> SayakSuggestionOf(List<String> words)
        {
            String fontName = "UNICODE";
            String remoteData = RemoteServiceData.ofSpellingSuggestion(fontName, words).ToString();
            String spellingResult = RemoteIO.SuggestSpellWithPostRequest(remoteData);
            return RemoteIO.buildSayakSuggestions(JToken.Parse(spellingResult));
        }

        public static Tuple<SayakResponse, SayakSuggestion> SayakSuggestionOf(String wordText, String fontName)
        {
            Tuple<SayakResponse, List<SayakSuggestion>> sayakResponse = SayakSuggestionOf(new List<String> { wordText }, fontName);
            return new Tuple<SayakResponse, SayakSuggestion>(sayakResponse.Item1, sayakResponse.Item2.First());
        }

        public static Tuple<Boolean, SayakResponse> GetClientEligibility()
        {
            try
            {
                var suggestion = SayakSuggestionOf(DemoData.SINGLE_WORD, TextUtility.UNICODE);
                return new Tuple<Boolean, SayakResponse>(suggestion.Item2.HasSuggestions(), suggestion.Item1);
            }
            catch
            {
                return new Tuple<Boolean, SayakResponse>(false,SayakResponse.serverNotAvailable());
            }
        }

        /// <summary>
        /// Checks if the remote spelling server/service is available or not.
        /// </summary>
        /// <returns>true</returns> if the servier is available.
        public static Boolean IsLocalServerAvailable()
        {
            try
            {
                using (var httpClient = initClient())
                {
                    var response = httpClient.GetAsync(FormatServerData(TextUtility.ACTION_SPELL_CHECK, TextUtility.UNICODE, DemoData.SINGLE_WORD)).Result;
                    return response.IsSuccessStatusCode;
                }
            }
            catch
            {
                return false;
            }
        }

        /// <summary>
        /// Verifies the availability of remote server by sending a single word to check spelling.
        /// </summary>
        /// <returns>Returns true if the result is not empty.</returns>
        public static Boolean IsRemoteServerAvailable()
        {
            try
            {
                HttpClient client = new HttpClient();                
                HttpResponseMessage response = client.GetAsync(SpellSettings.Default.pingURL).Result;
                return response.IsSuccessStatusCode;
            }
            catch (Exception e)
            {
                return false;
            }           
        }

        public static Tuple<Boolean, SayakResponse> IsEligibleAccount()
        {
            // Checks the user account information by sending the appId information.
            var suggestion = SayakSuggestionOf(DemoData.SINGLE_WORD, TextUtility.UNICODE);
            return new Tuple<Boolean, SayakResponse>(suggestion.Item2.HasSuggestions(), suggestion.Item1);
        }

        /// <summary>
        /// Validates if the spelling service is available in local server.
        /// </summary>
        /// <returns>true if the spelling service is available from localhost.</returns>
        public static Boolean IsLocalServerSource()
        {
            return SERVER_SOURCE == "LOCALHOST" || SERVER_SOURCE == "127.0.0.1";
        }

        /// <summary>
        /// Initialize the http client object.
        /// </summary>
        /// <returns></returns>
        private static HttpClient initClient()
        {
            var httpClient = new HttpClient();
            httpClient.Timeout = TimeSpan.FromMilliseconds(DEFAULT_TIMEOUT);
            return httpClient;
        }

        private static List<SayakSuggestion> buildSayakSuggestions(String spellingResult)
        {
            try
            {
                return JsonConvert.DeserializeObject<List<SayakSuggestion>>(spellingResult);
            }
            catch
            {
                return new List<SayakSuggestion>() { SayakSuggestion.Empty()};
            }
        }

        private static List<SayakSuggestion> buildSayakSuggestions(JToken jsonToken)
        {
            try
            {
                return jsonToken.ToObject<List<SayakSuggestion>>();
            }
            catch
            {
                return new List<SayakSuggestion>() { SayakSuggestion.Empty() };
            }
        }

        public static String BuildServiceRenewURL()
        {           
            return $"{SpellSettings.Default.serviceRenewURL}/{COMUtility.APP_ID}";
        }

        public static String BuildServicePaymentURL()
        {
            String pcName = Environment.MachineName;            
            return $"{SpellSettings.Default.servicePaymentURL}/{pcName}/{COMUtility.APP_ID}";
        }

        public static String BuildContactURL()
        {
            return "https://hijje.com/#/user/contact";
        }

        public static String buildAddNewWordData(String word, String fontName, String key = "fontconversion-9998")
        {
            String template = @"{"
                    + "\"@context\": \"http://semantro.com/\","
                    + "\"@type\": \"SayakMutation\","
                    + "\"actionName\": \"dumpPluginSuggestedWord\","
                    + "\"data\": {"
                                + "\"@context\": \"http://semantro.com/\","
                                + "\"@type\": \"WordPlugin\","
                                + "\"wordPluginId\": \"CLIENT_KEY\","
                                + "\"hasUsedWordPlugin\": {"
                                      + "\"@context\": \"http://semantro.com/\","
                                      + "\"@type\": \"WordPluginUseLog\","
                                      + "\"description\": \"SUGGESTEDWORDFROMPLUGIN\","
                                      + "\"sameAs\": \"SUGGESTED_WORD\","
                                      + "\"disambiguatingDescription\": \"FONT_NAME\""
                                      + "}"
                            + "}"
                    + "}";
            return template
                .Replace("CLIENT_KEY", key)
                .Replace("SUGGESTED_WORD", word)
                .Replace("FONT_NAME", fontName);
        }

    }
}
