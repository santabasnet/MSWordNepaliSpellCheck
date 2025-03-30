using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace WordNepaliSpellCheck.NepaliSpell
{
    class WordTokenizer
    {
        public static readonly char[] ttfDelimeters = { '\r', '\n', ' ', '\t', '\b', '\v' };

        public static readonly char[] unicodeDelimeter = { };

        public static List<String> getWordTokens(String currentText)
        {
            return currentText
                .Split(ttfDelimeters, StringSplitOptions.RemoveEmptyEntries)
                .Select(word => word.Trim())
                .Where(TextUtility.IsValidWord).ToList();
        }
    }
}
