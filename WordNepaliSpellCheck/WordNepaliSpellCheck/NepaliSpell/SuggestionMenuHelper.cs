using Microsoft.Office.Core;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace WordNepaliSpellCheck.NepaliSpell
{
    class SuggestionMenuHelper
    {
        /// <summary>
        /// Option to ignore word.
        /// </summary>
        public static readonly String IGNORE_WORD = "बेवास्ता गर्नुहोस्";

        /// <summary>
        /// Option to add to dictionary.
        /// </summary>
        public static readonly String ADD_TO_DICTIONARY = "शब्दकोषमा थप्नुहोस्";

        /// <summary>
        /// Option for no spellling suggestions.
        /// </summary>
        public static readonly String NO_SUGGESTIONS = "सुझाउहरू उपलब्ध हुन सकेनन्";

        /// <summary>
        /// Default menu options, it needs to be done in each suggestions.
        /// </summary>
        public static readonly List<String> WORD_ADDITIONAL_ACTIONS = new List<String> { IGNORE_WORD, ADD_TO_DICTIONARY };

        /// <summary>
        /// No suggestions menu items.
        /// </summary>
        public static readonly List<String> NO_SUGGESTIONS_ITEM = new List<string> { NO_SUGGESTIONS };

        /// <summary>
        /// Constant representation of tag to Nepali Spelling Check command.
        /// </summary>
        public static readonly String COMMAND_BAR_POPUP_TAG = "nepali_spell_check";

        /// <summary>
        /// Represents Nepali spelling check option literal.
        /// </summary>
        public static readonly String NEPALI_SPELL_CHECK_COMMAND = "नेपाली हिज्जे प्रणाली";

        /// <summary>
        /// Text option for command bar.
        /// </summary>
        public static readonly String COMMAND_BAR_TEXT = "Text";

        /// <summary>
        /// Text option for command bar in tables.
        /// </summary>
        public static readonly String COMMAND_BAR_TABLES = "Table Text";

        /// <summary>
        /// List option for command bar in bullets/numbers.
        /// </summary>
        public static readonly String COMMAND_BAR_LISTS = "Lists";
        
    }
}
