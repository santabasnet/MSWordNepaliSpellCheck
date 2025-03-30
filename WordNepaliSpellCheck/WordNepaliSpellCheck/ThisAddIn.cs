using System;
using Word = Microsoft.Office.Interop.Word;
using System.ComponentModel;
using WordNepaliSpellCheck.NepaliSpell;
using Microsoft.Office.Core;
using System.Linq;
using System.Collections.Generic;
using System.Diagnostics;
using System.Windows.Forms;
using System.Runtime.InteropServices;
using System.Threading;

namespace WordNepaliSpellCheck
{
    public partial class ThisAddIn
    {                
        /// <summary>
        /// Control for Nepali Spell Check.
        /// </summary>
        public SpellingControl spellingControl;

        /// <summary>
        /// MS word app, name alias.
        /// </summary>
        public Word.Application nepaliSpellApp = null;

        /// <summary>
        /// MS word app setting to keep the initial environment.
        /// </summary>
        public WordSettings wordSettings = null;

        /// <summary>
        /// Spelling worker and its associated flag.
        /// </summary>
        public BackgroundWorker spellingWorker = null;

        /// <summary>
        /// Flag that represents to the state of spelling worker thread.
        /// </summary>
        public Boolean restartWorker = false;

        /// <summary>
        /// Task-pane representation, it is used to show the statistics of spelling
        /// correction for Nepali Language.
        /// </summary>
        public CustomTaskPane spellingTaskPane;

        /// <summary>
        /// Current selection of text. Basically it is used to represent the current index
        /// to perform the spelling  correction.
        /// </summary>
        public Word.Selection currentSelection = null;

        /// <summary>
        /// Representation of current word.
        /// </summary>
        public NepaliWord currentWord = null;

        /// <summary>
        /// Event startup implementation.
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void ThisAddIn_Startup(object sender, EventArgs e)
        {
            ApplicationInitialize();                        
            Globals.ThisAddIn.Application.DocumentOpen += new Word.ApplicationEvents4_DocumentOpenEventHandler(Application_DocumentOpen);
            ((Word.ApplicationEvents4_Event)Application).NewDocument += Application_NewDocument;
        }

       
        /// <summary>
        /// Shutdown the word application plugin, it restores the the previous MS word settings
        /// before doing this.
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void ThisAddIn_Shutdown(object sender, System.EventArgs e)
        {
            this.wordSettings.Restore();
        }               

        private void Application_DocumentOpen(Word.Document Doc)
        {
            ApplicationInitialize();
        }

        private void Application_NewDocument(Word.Document Doc)
        {
            ApplicationInitialize();
        }

        /// <summary>
        /// Initializes the application environment.
        /// </summary>
        private void ApplicationInitialize()
        {
            if(this.nepaliSpellApp == null) this.nepaliSpellApp = Application;
            if(this.wordSettings == null) this.wordSettings = new WordSettings();
            if(this.currentWord == null) this.currentWord = NepaliWord.Empty();

            Globals.ThisAddIn.Application.DocumentBeforeClose += NepaliSpellApp_DocumentBeforeClose;
            Globals.ThisAddIn.Application.WindowSelectionChange +=
                new Word.ApplicationEvents4_WindowSelectionChangeEventHandler(Application_WindowSelectionChange);
            Globals.ThisAddIn.Application.WindowBeforeRightClick +=
                new Word.ApplicationEvents4_WindowBeforeRightClickEventHandler(Application_WindowBeforeRightClick);

            ///Initialize event hooks.          
        }

        /// <summary>
        /// Event handler for document close action.
        /// The initial MS word setting needs to be Restore before document close.
        /// </summary>
        /// <param name="Doc"></param>
        /// <param name="Cancel"></param>
        private void NepaliSpellApp_DocumentBeforeClose(Word.Document Doc, ref bool Cancel)
        {
            RemoveNepaliSpellContextMenu(this.nepaliSpellApp.CommandBars[SuggestionMenuHelper.COMMAND_BAR_TEXT]);
            this.wordSettings.Restore();
        }

        /// <summary>
        /// Event handler before right click. If the current selection is of wrong
        /// Nepali Word, then it needs to have spelling suggestion in context menu.
        /// </summary>
        /// <param name="Sel"></param>
        /// <param name="Cancel"></param>
        private void Application_WindowBeforeRightClick(Word.Selection Sel, ref bool Cancel)
        {
            //Check initial selection validation.
            if (!COMUtility.IsTextSelected(Sel) || COMUtility.IsIgnoredObjectSelected(Sel))
            {
                RemoveNepaliSpellContextMenu(Sel);
                return;
            }

            Tuple<Boolean, NepaliWord> contextMenuInfo = Checker.NeedsToAddMenuItem(Sel);
            /// Verify initial condition that a context menu is to be built on right click or not.
            if (!COMUtility.IsNepaliSpellingProcessable() || !contextMenuInfo.Item1)
            {
                RemoveNepaliSpellContextMenu(Sel);
                return;
            }           

            /// Initialize the current selection.
            this.currentSelection = Sel;

            /// Check where and which type of object to create the context menu.
            /// 1. Create a context menu for table or context text.
            /// 

            if (Sel.Type == Word.WdSelectionType.wdSelectionIP)
            {
                if ((Boolean)Sel.Information[Word.WdInformation.wdWithInTable]) AddNepaliSpellContextMenuInTable(contextMenuInfo.Item2);
                if (IsBulletNumbersSelection(Sel)) AddNepaliSpellContextMenuInBullets(contextMenuInfo.Item2);
                else AddNepaliSpellContextMenu(contextMenuInfo.Item2);
            }

        }

        /// <summary>
        /// Checks if the current select is in the list format or not.
        /// </summary>
        /// <param name="Sel"></param>
        /// <returns></returns>
        private Boolean IsBulletNumbersSelection(Word.Selection Sel)
        {
            //return Enum.IsDefined(typeof(Word.WdListType), Sel.Range.ListFormat.ListType);
            return Sel.Range.ListFormat.ListType == Word.WdListType.wdListBullet
                || Sel.Range.ListFormat.ListType == Word.WdListType.wdListListNumOnly
                || Sel.Range.ListFormat.ListType == Word.WdListType.wdListMixedNumbering
                || Sel.Range.ListFormat.ListType == Word.WdListType.wdListOutlineNumbering
                || Sel.Range.ListFormat.ListType == Word.WdListType.wdListPictureBullet
                || Sel.Range.ListFormat.ListType == Word.WdListType.wdListSimpleNumbering;
        }

        /// <summary>
        /// Build and display context menu for the current mis-spelled Nepali word.
        /// </summary>
        private void AddNepaliSpellContextMenu(NepaliWord currentWord)
        {
            RemoveNepaliSpellContextMenu(this.nepaliSpellApp.CommandBars[SuggestionMenuHelper.COMMAND_BAR_TEXT]);
            CommandBar spellingCommandBar = this.nepaliSpellApp.CommandBars[SuggestionMenuHelper.COMMAND_BAR_TEXT];
            CommandBarPopup parentCommandBarControl = CreateSubMenu(spellingCommandBar);
            GenerateSuggestionMenu(parentCommandBarControl, currentWord);
        }

        /// <summary>
        /// Build and display context menu for the current mis-spelled Nepali word.
        /// </summary>
        private void AddNepaliSpellContextMenuInTable(NepaliWord currentWord)
        {
            RemoveNepaliSpellContextMenu(this.nepaliSpellApp.CommandBars[SuggestionMenuHelper.COMMAND_BAR_TABLES]);
            CommandBar spellingCommandBar = this.nepaliSpellApp.CommandBars[SuggestionMenuHelper.COMMAND_BAR_TABLES];
            CommandBarPopup parentCommandBarControl = CreateSubMenu(spellingCommandBar);
            GenerateSuggestionMenu(parentCommandBarControl, currentWord);
        }

        /// <summary>
        /// Build and display context menu for the current mis-spelled Nepali word in Number/Bullet lists.
        /// </summary>
        private void AddNepaliSpellContextMenuInBullets(NepaliWord currentWord)
        {
            RemoveNepaliSpellContextMenu(this.nepaliSpellApp.CommandBars[SuggestionMenuHelper.COMMAND_BAR_LISTS]);
            CommandBar spellingCommandBar = this.nepaliSpellApp.CommandBars[SuggestionMenuHelper.COMMAND_BAR_LISTS];
            CommandBarPopup parentCommandBarControl = CreateSubMenu(spellingCommandBar);
            GenerateSuggestionMenu(parentCommandBarControl, currentWord);
        }

        /// <summary>
        /// Generates the suggestion menu for the current Nepali word.
        /// </summary>
        /// <param name="suggestionOptions"></param>
        /// <param name="currentWord"></param>
        private void GenerateSuggestionMenu(CommandBarPopup suggestionOptions, NepaliWord currentWord)
        {
            List<String> suggestedWords = SpellingService.GetSuggestions(currentWord);
            if (!suggestedWords.Any()) SuggestionMenuHelper.NO_SUGGESTIONS_ITEM
                .ForEach(additionalAction => BuildNoSuggestionMenuItem(suggestionOptions, additionalAction));
            else
            {
                this.currentWord = currentWord;
                suggestedWords.ForEach(word => BuildContextMenuItem(suggestionOptions, word, currentWord.fontName));               
            }
            SuggestionMenuHelper
                .WORD_ADDITIONAL_ACTIONS
                .ForEach(additionalAction => BuildAdditionalMenuItem(suggestionOptions, additionalAction));
        }

        /// <summary>
        /// Builds the menu item for the given word and its associated font name.
        /// </summary>
        /// <param name="suggestionOptions"></param>
        /// <param name="wordText"></param>
        /// <param name="fontName"></param>
        private void BuildContextMenuItem(CommandBarPopup suggestionOptions, String wordText, String fontName)
        {
            var commandBarButton = (CommandBarButton)suggestionOptions
                .Controls.Add(MsoControlType.msoControlButton, Type.Missing, Type.Missing, Type.Missing, true);
            commandBarButton.Click += SpellEventHandler;
            commandBarButton.Caption = wordText;
            commandBarButton.FaceId = 340;
            commandBarButton.Tag = TextUtility.BuildWordKey(fontName, wordText);
            commandBarButton.BeginGroup = false;
        }

        /// <summary>
        /// Builds the menu item for additional word actions such as ignore word, add to dictionary.
        /// </summary>
        /// <param name="suggestionOptions"></param>
        /// <param name="itemText"></param>
        private void BuildAdditionalMenuItem(CommandBarPopup suggestionOptions, String itemText)
        {
            var commandBarButton = (CommandBarButton)suggestionOptions
                .Controls.Add(MsoControlType.msoControlButton, Type.Missing, Type.Missing, Type.Missing, true);
            commandBarButton.Click += SpellEventHandler;
            commandBarButton.Caption = itemText;
            commandBarButton.FaceId = 319;
            commandBarButton.Tag = TextUtility.ConvertTagItem(itemText);
            commandBarButton.BeginGroup = true;
        }

        /// <summary>
        /// No suggestions menu item rendering.
        /// </summary>
        /// <param name="suggestionOptions"></param>
        /// <param name="itemText"></param>
        private void BuildNoSuggestionMenuItem(CommandBarPopup suggestionOptions, String itemText)
        {
            var commandBarButton = (CommandBarButton)suggestionOptions
                .Controls.Add(MsoControlType.msoControlButton, Type.Missing, Type.Missing, Type.Missing, true);
            commandBarButton.Caption = itemText;
            commandBarButton.FaceId = 342;
            commandBarButton.Tag = TextUtility.ConvertTagItem(itemText);
            commandBarButton.BeginGroup = true;
        }

        /// <summary>
        /// Creates the sub-menu for Nepali spelling suggestions after right click.
        /// </summary>
        /// <param name="spellingCommandBar"></param>
        /// <returns></returns>
        private CommandBarPopup CreateSubMenu(CommandBar spellingCommandBar)
        {
            bool isFound = false;
            CommandBarPopup parentCommandBarControl = null;

            foreach (var commandBarPopup in spellingCommandBar.Controls.OfType<CommandBarPopup>())
            {
                if (commandBarPopup.Tag.Equals(SuggestionMenuHelper.COMMAND_BAR_POPUP_TAG))
                {
                    isFound = true;
                    parentCommandBarControl = commandBarPopup;
                    break;
                }
            }
            if (!isFound)
            {
                parentCommandBarControl = (CommandBarPopup)spellingCommandBar.Controls
                    .Add(MsoControlType.msoControlPopup, Type.Missing, Type.Missing, Type.Missing, true);
                parentCommandBarControl.Caption = SuggestionMenuHelper.NEPALI_SPELL_CHECK_COMMAND;
                parentCommandBarControl.Tag = SuggestionMenuHelper.COMMAND_BAR_POPUP_TAG;
                parentCommandBarControl.Visible = true;
            }

            return parentCommandBarControl;
        }

        /// <summary>
        /// Removes the context menu for Nepali Spelling Check.
        /// </summary>
        /// <param name="currentSelection"></param>
        private void RemoveNepaliSpellContextMenu(Word.Selection currentSelection)
        {
            if ((bool)currentSelection.Information[Word.WdInformation.wdWithInTable])
            {
                RemoveNepaliSpellContextMenu(this.nepaliSpellApp.CommandBars[SuggestionMenuHelper.COMMAND_BAR_TABLES]);
                return;
            }
            RemoveNepaliSpellContextMenu(this.nepaliSpellApp.CommandBars[SuggestionMenuHelper.COMMAND_BAR_TEXT]);
        }

        /// <summary>
        /// Removes the context menu for Nepali Spelling Check.
        /// </summary>
        /// <param name="spellingCommandBar"></param>
        private void RemoveNepaliSpellContextMenu(CommandBar spellingCommandBar)
        {
            foreach (var commandBarPopup in spellingCommandBar.Controls.OfType<CommandBarPopup>())
            {
                if (commandBarPopup.Tag.Equals(SuggestionMenuHelper.COMMAND_BAR_POPUP_TAG))
                {
                    commandBarPopup.Delete();
                }
            }
        }

        /// <summary>
        /// Make action changes to the text after selecting the correct suggestion for a Nepail word.
        /// </summary>
        /// <param name="suggestionControl"></param>
        /// <param name="CancelDefault"></param>
        private void SpellEventHandler(CommandBarButton suggestionControl, ref bool CancelDefault)
        {
            String correctWord = suggestionControl.Caption;
            if (IsSpecialSelection(correctWord))  HandleSpecialSelection(correctWord);
            else if (COMUtility.IsShapeSelected(Globals.ThisAddIn.currentSelection)) Checker.ReplaceWrongWordInShapes(correctWord, currentWord);
            else if (COMUtility.IsHeaderFooterSelected(Globals.ThisAddIn.currentSelection)) Checker.ReplaceWrongWordInHeaderFooter(correctWord, currentWord);
            else Checker.ReplaceWrongWordInDocumentTables(correctWord, currentWord);
        }
        
        private void HandleSpecialSelection(String subMenu)
        {
            if (subMenu == SuggestionMenuHelper.IGNORE_WORD)
            {
                MakeIgnoreWord();
                return;
            }

            if(subMenu == SuggestionMenuHelper.ADD_TO_DICTIONARY)
            {
                var wordDumpResponse = RemoteIO.SuggestedWordFromClient(this.currentSelection.Words.First.Text.Trim());
                var title = "Add to Dictionary: " + wordDumpResponse.GetStatus();
                MessageBox.Show(wordDumpResponse.GetMessage(), title, MessageBoxButtons.OK, MessageBoxIcon.Information);
                MakeIgnoreWord();
                return;
            }
            // Other options.
        }
        
        private void MakeIgnoreWord()
        {
            String wordText = this.currentSelection.Words.First.Text.Trim();
            String formattedWord = this.currentSelection.Words.First.Text;

            //1. Add to ignore word in the local cache
            SpellingService.AddIgnoreWord(formattedWord.Trim());
            SpellingService.RemoveIgnoredWordFromSuggestions(wordText);

            //2. Clear the formatting in the document content-pane and tables.
            SearchUtility.ClearSpellingErrors(formattedWord, Globals.ThisAddIn.nepaliSpellApp.ActiveDocument.Range());

            //3. Clear the formatting in the available shapes.
            foreach (Word.Shape shape in this.nepaliSpellApp.ActiveDocument.Shapes)
            {
                var shapeRange = shape.TextFrame.TextRange;
                SearchUtility.ClearSpellingErrors(formattedWord, shapeRange);
            }

            //4. Clear the spell error formatting in the header and footer section.
            foreach (Word.Section section in currentWord.wordReference.Application.ActiveWindow.Document.Sections)
            {
                Word.Range headerRange = section.Headers[Word.WdHeaderFooterIndex.wdHeaderFooterPrimary].Range;
                SearchUtility.ClearSpellingErrors(formattedWord, headerRange);
                Word.Range footerRange = section.Footers[Word.WdHeaderFooterIndex.wdHeaderFooterPrimary].Range;
                SearchUtility.ClearSpellingErrors(formattedWord, footerRange);
            }
        } 
        
        private Boolean IsSpecialSelection(String subMenu)
        {
            return SuggestionMenuHelper.WORD_ADDITIONAL_ACTIONS.Contains(subMenu);
        }

        /// <summary>
        /// Event definition for window selection changed.
        /// The main spelling error detection has been done with this event.
        /// </summary>
        /// <param name="currentSelection"></param>
        private void Application_WindowSelectionChange(Word.Selection currentSelection)
        {
            if (!COMUtility.IsNepaliSpellingProcessable()) return;
            this.restartWorker = true;
            this.currentSelection = currentSelection;
            if (!this.spellingWorker.IsBusy) this.spellingWorker.RunWorkerAsync();
        }

        /// <summary>
        /// Check if the header footer is activated now.
        /// </summary>
        /// <returns></returns>
        public static Boolean IsHeaderActivated()
        {
            return COMUtility.IsHeaderFooterSelected(Globals.ThisAddIn.currentSelection);
        }

        /// <summary>
        /// Show spelling panel for statistic purpose.
        /// </summary>
        private void AddSpellingPanel()
        {
            //spellingControl = new SpellingControl();
            //spellingTaskPane = CustomTaskPanes.Add(spellingControl, "Nepali Spelling Tasks");
            //spellingTaskPane.Width = 350;
            //spellingTaskPane.Visible = true;
        }         

        #region VSTO generated code

        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InternalStartup()
        {
            this.Startup += new System.EventHandler(ThisAddIn_Startup);
            this.Shutdown += new System.EventHandler(ThisAddIn_Shutdown);
        }

        #endregion
        ///////////////////////////////////////////////


    }
}
