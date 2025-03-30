using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace WordNepaliSpellCheck.NepaliSpell
{
    class DemoData
    {

        public Dictionary<String, List<String>> kantipur;
        public Dictionary<String, List<String>> unicode;

        public DemoData()
        {
            this.kantipur = loadKantipurData();
            this.unicode = loadUnicodeData();
        }

        private Dictionary<String, List<String>> loadUnicodeData()
        {
            Dictionary<String, List<String>> data = new Dictionary<String, List<String>>();
            data.Add("गरिसकको", new List<String>() { "गरिसकेको", "गरिसेकको", "गरिसके", "गरिसकी", "गरिसक्नको", "गरिसकेकको", "गरिसकिएको", "गरिसकेकोछ", "गरिससकेको", "गरिसकियो", "गरिसक्दो", "गरिसकेकै" }) ;
            data.Add("रामरो", new List<String>() { "रामरज", "रामको", "रामरथ", "रामो", "चामर", "समरो", "रमरम", "कामर", "पामर", "राखर", "रामरोशन", "रामरोहन" });
            data.Add("धमिल", new List<String>() { "धमिलै", "धमिलो", "दमित", "मिल", "दमकल", "दलिल", "नमिल", "तमिल", "धमेल", "धमला", "अमल", "कमल" });
            data.Add("कमसलल", new List<String>() { "कमसल", "कमसलको", "कमला", "कमली", "कमलले", "कमलो", "कमाल", "कमल", "कमले", "कमलै", "कमसँग", "कमसित" });
            data.Add("धाना", new List<String>() { "दाना", "धान", "थाना", "धनका", "धनाइ", "धापा", "धामा", "धागा", "धनदा", "गाना", "घाना", "काना" });
            return data;
        }

        private Dictionary<String, List<String>> loadKantipurData()
        {
            Dictionary<String, List<String>> data = new Dictionary<String, List<String>>();
            data.Add("ul;s]sf]", new List<String>() { "l;s]sf]", "gl;s]sf]", "al;s]sf]", "/l;ssf]", "uO;ss]sf]", "vl;;s]sf]", "uln;s]sf]", "ul/;s]sf]", "ul9;s]sf]", "ul8;s]sf]", "l´s]sf]", "gl;s]sf" });
            data.Add("/fd/f]", new List<String>() { "/fd/h", "/fdsf]", "/fd/y", "/fdf]", "rfd/", ";d/f]", "/d/d", "sfd/", "kfd/", "/fv/", "/fd/f]zg", "/fd/f]xg" });
            data.Add("wldn", new List<String>() { "wldn}", "wldnf]", "bldt", "ldn", "bdsn", "blnn", "gldn", "tldn", "wd]n", "wdnf", "cdn", "sdn" });
            data.Add("sd;nn", new List<String>() { "sd;n", "sd;nsf]", "sdnf", "sdnL", "sdnn]", "sdnf]", "sdfn", "sdn", "sdn]", "sdn}", "sd;“u", "sdl;t" });
            data.Add("wfgf", new List<String>() { "bfgf", "wfg", "yfgf", "wgsf", "wgfO", "wfkf", "wfdf", "wfuf", "wgbf", "ufgf", "3fgf", "sfgf" });
            return data;
        }

        public static readonly string TRIAL_OVER_JSON = "{\"status\" : \"fail\",\"messageId\" : \"REDIRECT_TO_REGISTRATION\",\"message\" : \"Trial Period Over!, We request you to register and make payment.\",\"@type\" : \"SayakResponse\",\"@context\" : \"http://semantro.com\",\"@id\" : \"http://semantro.com/SayakResponse\"}";
        public static readonly string ACCOUNT_EXPIRED_JSON = "{\"status\" : \"fail\",\"messageId\" : \"REDIRECT_TO_SERVICE_RENEW\",\"message\" : \"Sorry, Your subscribed service duration is expired! We request you extend the spelling service.\",\"@type\" : \"SayakResponse\",\"@context\" : \"http://semantro.com\",\"@id\" : \"http://semantro.com/SayakResponse\"}";
        public static readonly string SINGLE_WORD = "कामाडौँ";
    }
}
