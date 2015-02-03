using System.Collections.Generic;
using System.Linq;
using System.Text.RegularExpressions;
using CommandLine;
using CommandLine.Text;

namespace SQL2Word
{
    class Options
    {

        [Option('i', "input_file", Required = true, HelpText = "docx template file")]
        public string TemplateFile { get; set; }
        [Option('o', "output_file", HelpText = "File to which output goes")]
        public string OutputFile { get; set; }
        [Option('s', "show_result_file_on_end", HelpText = "Open file at the process end")]
        public bool ShowResultFileOnEnd { get; set; }
        [Option('q', "save_with_queries", HelpText = "Сохраняет скрипты вместе с таблицам чтобы потом можно было бы обновить документ")]
        public bool SaveQueries{ get; set; }
        [Option('u', "update_file", HelpText = "Обновить файл")]
        public bool UpdateFile { get; set; }
//
//        [Option('a', "append", HelpText = "append output to existing files in output directory",
//            DefaultValue = false)]
//        public bool Append { get; set; }

        [OptionArray('p', "params",
            HelpText = "pass parameters as list of tuples: key1=value1 key2=value2 ...")]
        public string[] ParametersArg { get; set; }
//
//        [Option('s', "show parameters", HelpText = "show available parameters")]
//        public bool ListParameters { get; set; }

        [Option('c', "connection options", HelpText = "file with connection options:"
            + "\n\tUserID=YourUsername"
            + "\n\tPassword=YourPass"
            + "\n\tDataSource=YourDataSourceName"
            + "\n\tInitialCatalog=YourInitialCatalogName")]
        public string CredentialFile { get; set; }

        public Dictionary<string, string> Parameters
        {
            get
            {
                var reg = new Regex(@"(\w+)=(.+)");
                var output = new Dictionary<string, string>();
                if (ParametersArg != null)
                {
                    foreach (var item in ParametersArg)
                    {
                        var m = reg.Match(item);
                        if (m.Success)
                        {
                            output.Add(m.Groups[1].Value, m.Groups[2].Value);
                        }
                    }
                }
                return output;
            }
        }

        [HelpOption]
        public string GetUsage()
        {
            return HelpText.AutoBuild(this,
                (HelpText current) => HelpText.DefaultParsingErrorsHandler(this, current));
        }
    }
}
