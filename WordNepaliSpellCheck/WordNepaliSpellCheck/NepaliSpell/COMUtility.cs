using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Microsoft.Office.Interop.Word;
using Microsoft.VisualBasic;
using System.Net.NetworkInformation;
using System.Text.RegularExpressions;
using System.Security.Cryptography;

namespace WordNepaliSpellCheck.NepaliSpell
{
    class COMUtility
    {
        /// <summary>
        /// Indicates that the word is be selected from document object model.
        /// </summary>
        public static readonly int DOCUMENT = 0;

        /// <summary>
        /// Indicates that the word is to be selected from table cells.
        /// </summary>
        public static readonly int TABLE = 1;

        /// <summary>
        /// Indicates that the word is to be selected from Shapes. 
        /// </summary>
        public static readonly int SHAPE = 2;

        /// <summary>
        /// Indicates that the word is to be selected from the Header or Footer.
        /// </summary>
        public static readonly int HEADER_FOOTER = 3;
    
        /// <summary>
        /// Represents the selection type, and is used while making a right click.
        /// </summary>
        public static readonly String SELECTION_TYPE = "Selection";

        /// <summary>
        /// App Id for the Spelling Plugin.
        /// </summary>
        public static readonly String APP_ID = COMUtility.GenerateAppId();

        /// <summary>
        /// It determines the all the sections that are excluded to check the nepali spelling errors. 
        /// </summary>
        /// <param name="currentSelection"></param>
        /// <returns></returns>
        public static Boolean IsIgnoredObjectSelected(Selection currentSelection)
        {
            //if ((bool)currentSelection.Information[WdInformation.wdInHeaderFooter]) return true;
            if ((bool)currentSelection.Information[WdInformation.wdInFootnoteEndnotePane]) return true;
            if ((bool)currentSelection.Information[WdInformation.wdInCommentPane]) return true;
            if ((bool)currentSelection.Information[WdInformation.wdInClipboard]) return true;
            if ((bool)currentSelection.Information[WdInformation.wdInBibliography]) return true;
            if ((bool)currentSelection.Information[WdInformation.wdInCitation]) return true;
            if ((bool)currentSelection.Information[WdInformation.wdInContentControl]) return true;
            //if (currentSelection.Application.ActiveWindow.Selection.ShapeRange.Count > 0) return true;
            return false;
        }

        public static Boolean IsTextSelected(Selection Sel)
        {
            return Information.TypeName(Sel) == COMUtility.SELECTION_TYPE && !String.IsNullOrWhiteSpace(Sel.Text);
        }

        public static Boolean IsTableSelected(Selection currentSelection)
        {
            return (bool)currentSelection.Information[WdInformation.wdWithInTable] == true;
        }

        public static Boolean IsHeaderFooterSelected(Selection currentSelection)
        {
            return (bool)currentSelection.Information[WdInformation.wdInHeaderFooter] == true;
        }

        public static Boolean IsShapeSelected(Selection currentSelection)
        {
            return currentSelection.Application.ActiveWindow.Selection.ShapeRange.Count > 0;  
        }

        /// <summary>
        /// Validates if the current event of selections needs to be processed further or not.
        /// </summary>
        /// <returns>true if all the validation passed.</returns>
        public static Boolean IsNepaliSpellingProcessable()
        {
            Boolean spellingAction = Globals.Ribbons.SpellingRibbon.spellingCheckButton.Checked;
            Boolean initialized = !(Globals.ThisAddIn.spellingWorker == null);
            return (spellingAction && initialized);
        }

        /// <summary>
        /// It generates the application ID based on the installed machine.
        /// And, it encrypts with the given encryption function.
        /// </summary>
        /// <returns></returns>
        public static String GenerateAppId()
        {
            String appId = GenerateId();
            //String appId = "fontconversion-9998";
            //return encryptId(appId);
            return appId;
        }

        private static String GenerateId()
        {
            List<String> macAddress = MacAddress();
            List<String> osInformation = OSInformation();
            List<String> generatedTokens = new List<String>()
            {
                ServiceName(),
                osInformation.First(),
                macAddress.First(),                
                macAddress.Last(),
                osInformation.Last()
            };
            return String.Join(TextUtility.TOKEN_SEPARATOR, generatedTokens);
        }

        private static bool isValidMAC(NetworkInterface nic)
        {
            string physicalAddress = nic.GetPhysicalAddress().ToString();

            return nic.Speed > -1 
                && !string.IsNullOrEmpty(physicalAddress) 
                && nic.NetworkInterfaceType != NetworkInterfaceType.Loopback
                && physicalAddress.Length >= 12;
        }

        /// <summary>
        /// Generates list of Hex codes of 6 digits from the first machine MAC address.
        /// </summary>
        /// <returns>listOfMACAddressOfHexCodes</returns>
        private static List<String> MacAddress()
        {
            String macAddress =  NetworkInterface
                .GetAllNetworkInterfaces()
                .Where(nic => nic.OperationalStatus == OperationalStatus.Up && nic.NetworkInterfaceType != NetworkInterfaceType.Loopback)
                .Select(nic => nic.GetPhysicalAddress().ToString())
                .FirstOrDefault();

            return (from Match token in Regex.Matches(macAddress, @"\S{6}") select token.Value).ToList();
        }

        /// <summary>
        /// Generates the list of Hex codes of 6 digits from the current version of Installed OS.
        /// </summary>
        /// <returns>listOfOSInformationInHexCodes</returns>
        private static List<String> OSInformation()
        {
            String osInformation = String.Join(TextUtility.EMPTY, Environment.OSVersion.VersionString.ToCharArray().Select(letter => ((int)letter).ToString("X2")));
            return (from Match token in Regex.Matches(osInformation, @"\S{6}") select token.Value).ToList();
        }

        /// <summary>
        /// Extracts service name from the settings.
        /// </summary>
        /// <returns>serviceName</returns>
        private static String ServiceName()
        {
            return SpellSettings.Default.serviceName;
        }

        /// <summary>
        /// Generates the string representation of the Hex digits from the installed office version number.
        /// </summary>
        /// <returns>wordVersionInHexCode</returns>
        private static String WordVersion()
        {
            String wordVersion = Globals.ThisAddIn.Application.Version.ToString();
            return ((int)(Double.Parse(wordVersion))).ToString("X2");
        }

        /// <summary>
        /// Encryption with AES, with hashed sha256 (From Sachin).
        /// </summary>
        /// <param name="text"></param>
        /// <param name="key"></param>
        /// <returns>encryptedKey</returns>
        private static String encryptId(String text, String key = "h!jje2@1")
        {
            RijndaelManaged rijndaelCipher = new RijndaelManaged();
            rijndaelCipher.Mode = CipherMode.CBC;
            rijndaelCipher.Padding = PaddingMode.PKCS7;

            SHA256 sha256 = SHA256.Create();
            byte[] passwordHashBytes = sha256.ComputeHash(Encoding.UTF8.GetBytes(key));
            byte[] passwordIvBytes = Encoding.UTF8.GetBytes(key);
            byte[] keyBytes = new byte[0x20];
            byte[] iVBytes = new byte[0x10];

            int lenIvBytes = passwordIvBytes.Length;
            int lenKeyBytes = passwordHashBytes.Length;
            Array.Copy(passwordHashBytes, keyBytes, lenKeyBytes);
            Array.Copy(passwordIvBytes, iVBytes, lenIvBytes);

            rijndaelCipher.Key = keyBytes;
            rijndaelCipher.IV = iVBytes;
            ICryptoTransform transform = rijndaelCipher.CreateEncryptor();
            byte[] plainText = Encoding.UTF8.GetBytes(text);

            return Convert.ToBase64String(transform.TransformFinalBlock(plainText, 0, plainText.Length)).Replace('+', '-').Replace('/', '_');
        }


    }
}
