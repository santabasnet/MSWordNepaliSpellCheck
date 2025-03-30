using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Drawing;
using System.Data;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Word = Microsoft.Office.Interop.Word;
using Microsoft.Office.Core;

namespace WordNepaliSpellCheck
{
    public partial class SpellingControl : UserControl
    {
        public System.Reflection.Assembly assemblyInfo;
        public Word.Application nepaliSpellApp;       
        public Word.Document currentDocument;


        public SpellingControl()
        {
            InitializeComponent();
            nepaliSpellApp = Globals.ThisAddIn.nepaliSpellApp;
            
            assemblyInfo = System.Reflection.Assembly.GetExecutingAssembly();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            //Demo works
            String word = richTextBox1.Text.Trim();
            Boolean result = NepaliSpell.SpellingService.IsCorrectWord(word);
            if (result)
            {
                listBox1.DataSource = NepaliSpell.SpellingService.GetSuggestions(word);
            }

        }

        private void button2_Click(object sender, EventArgs e)
        {
            Word.Document currentDocument = nepaliSpellApp.ActiveDocument;
            Word.Paragraph paragraph = currentDocument.Paragraphs.First;
            Word.Range range = paragraph.Range;            

            range.Underline = Word.WdUnderline.wdUnderlineWavy;
            range.Font.UnderlineColor = Word.WdColor.wdColorOrange;

            String text = currentDocument.Words.First.Text;
                                       
           

            MessageBox.Show(text);

        }

        private void Application_WindowSelectionChange(Word.Selection currentSelection)
        {
            MessageBox.Show(currentSelection.Text);
        }
               
    }
}
