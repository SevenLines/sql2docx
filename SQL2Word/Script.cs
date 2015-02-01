using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;

namespace SQLDynamic
{
    public class Script
    {
        String _text;
        String _title;
        Regex reg = new Regex(@"/\*\$(\w+)\*/(.*?)/\*\$\*/");
        Dictionary<string, string> prms = new Dictionary<string, string>();

        public Script(String title, FileInfo file)
        {
            using(var s = new StreamReader(file.OpenRead(), Encoding.Default)) {
                String text = s.ReadToEnd();
                Init(title, text);
            }
        }

        public Script(String title, String text)
        {
            Init(title, text);
        }

        public void Init(String title, String text)
        {
            this._text = text;
            this._title = title;
            foreach(Match m in reg.Matches(text)) {
                prms.Add(m.Groups[1].Value, m.Groups[2].Value);
            }
        }

        public String Title
        {
            get { return _title; }
        }

        public Dictionary<string, string> Params
        {
            get
            {
                return prms;
            }

        }

        public String Text(Dictionary<String, String> prms)
        {
            var evaluator = new MatchEvaluator(new Func<Match, String>( m =>
            {
                if (prms.ContainsKey(m.Groups[1].Value))
                {
                    return m.Value.Insert(m.Value.IndexOf(@"/*$*/"), prms[m.Groups[1].Value]);
                }
                return m.Value;
            }));
            return reg.Replace((String)_text.Clone(), evaluator);
        }
    }
}
