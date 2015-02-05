using System;
using System.Collections.Generic;
using System.Data.SqlClient;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Text.RegularExpressions;
using CommandLine;
using Novacode;

namespace SQL2Word
{
    class Program
    {
        private static void DrawProgressBar(int complete, int maxVal, int barSize, char progressCharacter)
        {
            Console.CursorVisible = false;
            int left = Console.CursorLeft;
            decimal perc = (decimal)complete / (decimal)maxVal;
            int chars = (int)Math.Floor(perc / ((decimal)1 / (decimal)barSize));
            string p1 = String.Empty, p2 = String.Empty;

            for (int i = 0; i < chars; i++) p1 += progressCharacter;
            for (int i = 0; i < barSize - chars; i++) p2 += progressCharacter;

            Console.ForegroundColor = ConsoleColor.Gray;
            Console.Write(p1);
            Console.ForegroundColor = ConsoleColor.DarkGray;
            Console.Write(p2);

            Console.ResetColor();
            Console.Write(" {0}%", (perc * 100).ToString("N2"));
            Console.CursorLeft = left;
        }

        public static SqlConnection connection = new SqlConnection();

        public static KeyValuePair<string, string>? ParseOptionLine(string line)
        {
            var regex = new Regex(@"(\w+):\s*([\w\\(.)/]+)");
            var match = regex.Match(line);

            if (!match.Success) return null;

            return new KeyValuePair<string, string>(match.Groups[1].Value, match.Groups[2].Value);
        }


        static public String GetConnectionString(String credentialFile)
        {
            var credentialDictionary = new Dictionary<string, string>();
            if (File.Exists(credentialFile))
            {
                foreach (var line in File.ReadAllLines(credentialFile))
                {
                    var tuple = ParseOptionLine(line);
                    if (!tuple.HasValue)
                        continue;
                    credentialDictionary.Add(tuple.Value.Key, tuple.Value.Value);
                }
            }
            var connString = new SqlConnectionStringBuilder
            {
                ApplicationName = Assembly.GetEntryAssembly().GetName().Name,
                DataSource = credentialDictionary.Default("DataSource", "(local)"),
                InitialCatalog = credentialDictionary.Default("InitialCatalog", ""),
            };
            
            if (credentialDictionary.ContainsKey("UserID"))
            {
                connString.UserID = credentialDictionary.Default("UserID", "");
                if (credentialDictionary.ContainsKey("Password"))
                    connString.Password = credentialDictionary.Default("Password", "");
            }
            else
            {
                connString.IntegratedSecurity = true;
            }

            return connString.ToString();
        }

        static void Main(string[] args)
        {
            var options = new Options();
            if (!Parser.Default.ParseArguments(args, options))
                return;

            DocX doc;

            try
            {
                doc = DocX.Load(options.TemplateFile);
            }
            catch (FileNotFoundException ex)
            {
                Console.WriteLine(ex.Message);
                Console.ReadKey();
                return;
            }
            catch (IOException ex)
            {
                Console.WriteLine(ex.Message);
                Console.ReadKey();
                return;                
            }

            // замена плэйсхолдеров
            var re = new Regex(@"{.*?}");

            // формируем список параграфов
            IEnumerable<Paragraph> allParagraphas = doc.Paragraphs;
            if (doc.Headers.even != null)
            {
                allParagraphas = allParagraphas.Union(doc.Headers.even.Paragraphs);
            }
            if (doc.Headers.odd != null)
            {
                allParagraphas = allParagraphas.Union(doc.Headers.odd.Paragraphs);
            }
            if (doc.Footers.even != null)
            {
                allParagraphas = allParagraphas.Union(doc.Footers.even.Paragraphs);
            }
            if (doc.Footers.odd != null)
            {
                allParagraphas = allParagraphas.Union(doc.Footers.odd.Paragraphs);
            }
            // заменяем плэйсхолдеры
            foreach (var paragraph in allParagraphas)
            {
                if (re.IsMatch(paragraph.Text))
                {
                    foreach (var p in options.Parameters)
                    {
                        paragraph.ReplaceText("{" + p.Key + "}", p.Value);
                    }
                }
            }

            String credentialFile = String.IsNullOrEmpty(options.CredentialFile)
                ? ".credentials"
                : options.CredentialFile;
            connection.ConnectionString = GetConnectionString(credentialFile);

            Console.WriteLine("Пробую подсоединиться");
            Console.WriteLine(connection.ConnectionString);
            try
            {
                connection.Open();
            }
            catch (SqlException ex)
            {
                Console.WriteLine(ex.Message);
                Console.ReadKey();
                return;
            }

            int i = 0;
            int tablesCount = doc.Tables.Count;
            foreach (var t in doc.Tables)
            {
                DrawProgressBar(i, tablesCount, 30, '█');
                if (options.UpdateFile)
                {
                    Filler.UpdateTable(t, connection, options.Parameters, options.SaveQueries);
                }
                else
                {
                    Filler.FillTable(t, connection, options.Parameters, options.SaveQueries);
                }
                i++;
            }
            DrawProgressBar(1, 1, 30, '█');
            Console.WriteLine();
            Console.WriteLine("Сохраняю резульат в " + options.OutputFile);

            try
            {
                if (options.OutputFile == options.TemplateFile)
                {
                    doc.Save();
                }
                else
                {
                    doc.SaveAs(options.OutputFile);
                }
                if (options.ShowResultFileOnEnd)
                {
                    Process.Start(options.OutputFile);
                }
            }
            catch (IOException ex)
            {
                Console.WriteLine(ex.Message);
                Console.ReadKey();
            }

            connection.Close();
        }
    }
}