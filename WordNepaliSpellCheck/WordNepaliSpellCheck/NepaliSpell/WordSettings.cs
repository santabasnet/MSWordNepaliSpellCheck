using System;

namespace WordNepaliSpellCheck.NepaliSpell
{
    public partial class WordSettings
    {
        private Boolean spellingCheckOption;
        private Boolean grammarCheckOption;

        /// <summary>
        /// Default constructor, that initializes the initial settings of the current 
        /// MS word application.
        /// </summary>
        public WordSettings()
        {
            this.spellingCheckOption = Globals.ThisAddIn.Application.Options.CheckSpellingAsYouType;
            this.grammarCheckOption = Globals.ThisAddIn.Application.Options.CheckGrammarAsYouType;
        }

        /// <summary>
        /// It resets the MS word spelling and grammar check option to false for Nepali language.
        /// </summary>
        public void ResetToNepaliSettings()
        {
            Globals.ThisAddIn.Application.Options.CheckSpellingAsYouType = false;
            Globals.ThisAddIn.Application.Options.CheckGrammarAsYouType = false;
        }

        /// <summary>
        /// Member to restore the previous setting of MS word.
        /// </summary>
        public void Restore()
        {
            Globals.ThisAddIn.Application.Options.CheckSpellingAsYouType = this.spellingCheckOption;
            Globals.ThisAddIn.Application.Options.CheckGrammarAsYouType = this.grammarCheckOption;
        }

    }
}
