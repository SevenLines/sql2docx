using System;
using System.Collections.Generic;
using System.Data.SqlClient;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using Novacode;
using Script = SQLDynamic.Script;

namespace SQL2Word
{
    class Filler
    {
        public enum TOKENS
        {
            CONTENT,
            WITH_COUNTER,
            USE_ZERO_CONUTER,
            AUTO_EXPAND,
            REPLACE_TABLE_ON_EMPTY,
            UPDATE_SCRIPT,
        };

        static readonly Regex re_REPLACE_TABLE_ON_EMPTY = new Regex(@"\[REPLACE_TABLE_ON_EMPTY(:(.*?))?\]");
        static readonly Regex re_UPDATE_SCRIPT = new Regex(@"\[UPDATE_SCRIPT(:(.*?))?\]");

        /// <summary>
        /// Return all tokens should used in current cell,
        /// All tokens should be in first paragraph
        /// </summary>
        /// <param name="cell"></param>
        /// <returns></returns>
        private static List<TOKENS> GetCellTokens(Cell cell, out Dictionary<TOKENS, String> special_parameters)
        {
            var output = new List<TOKENS>();
            special_parameters = new Dictionary<TOKENS, string>();

            Match match;

            special_parameters = new Dictionary<TOKENS, string>();
            var text = cell.Paragraphs.First().Text;
            if (text.Contains("[CONTENT]"))
            {
                output.Add(TOKENS.CONTENT);
//                special_parameters[TOKENS.CONTENT] = "";
            }
            if (text.Contains("[WITH_COUNTER]")) { 
                output.Add(TOKENS.WITH_COUNTER);
                special_parameters[TOKENS.WITH_COUNTER] = "";
            }
            if (text.Contains("[USE_ZERO_CONUTER]")) { 
                output.Add(TOKENS.USE_ZERO_CONUTER);
                special_parameters[TOKENS.USE_ZERO_CONUTER] = "";
            }
            if (text.Contains("[AUTO_EXPAND]")) { 
                output.Add(TOKENS.AUTO_EXPAND);
                special_parameters[TOKENS.AUTO_EXPAND] = "";
            }

            match = re_REPLACE_TABLE_ON_EMPTY.Match(text);
            if (match.Success)
            {
                output.Add(TOKENS.REPLACE_TABLE_ON_EMPTY);
                special_parameters[TOKENS.REPLACE_TABLE_ON_EMPTY] = match.Groups[2].Value;
            }

            match = re_UPDATE_SCRIPT.Match(text);
            if (match.Success)
            {
                output.Add(TOKENS.UPDATE_SCRIPT);
                special_parameters[TOKENS.UPDATE_SCRIPT] = match.Groups[2].Value;
            }

            return output;
        }


        private static List<TOKENS> GetTableTokens(Table table, out Dictionary<TOKENS, string> parameters)
        {
            var row = GetTableScriptRow(table);
            var cell = row.Cells.First();
            return GetCellTokens(cell, out parameters);
        }

        private static String GetTableScript(Table table)
        {
            var row = GetTableScriptRow(table);
            var cell = row.Cells.First();
            return GetCellContent(cell);
        }

        private static Row GetTableScriptRow(Table table)
        {
            return table.Rows.Last();
        }

        private static void SetTableScriptRowText(Table table, String text)
        {
            var cell = GetTableScriptRow(table).Cells.First();
            if (cell != null)
            {
                SetCellContent(cell, text);
            }
        }

        /// <summary>
        /// returns cell text content as all joined paragraphs text within cell
        /// </summary>
        /// <param name="cell"></param>
        /// <returns></returns>
        private static String GetCellContent(Cell cell, int skip=1)
        {
            StringBuilder output = new StringBuilder();
            foreach (var paragraph in cell.Paragraphs.Skip(skip))
            {
                output.AppendLine(paragraph.Text + " ");
            }
            return output.ToString();
        }

        private static void SetCellContent(Cell cell, String text)
        {
            cell.Paragraphs.ForEach(paragraph => paragraph.Remove(false));
            cell.InsertParagraph(text);
        }


        private static bool _fillTable(
            Table table, 
            SqlConnection connection, 
            String script,
            List<TOKENS> tokens,
            bool useFirstRow = false)
        {
            if (String.IsNullOrEmpty(script))
            {
                return false;
            }

            var sql = new SqlCommand(script, connection);
            SqlDataReader reader;
            try
            {
                reader = sql.ExecuteReader();

            }
            catch (SqlException ex)
            {
                SetTableScriptRowText(table, ex.Message);
                return false;
            }

            int counter = 1;
            int iOffset = 0;

            if (tokens.Contains(TOKENS.USE_ZERO_CONUTER))
                counter = 0;
            if (tokens.Contains(TOKENS.WITH_COUNTER))
                iOffset++;

            bool isFirstRow = true;

            do
            {
                while (reader.Read())
                {
                    Row r;
                    if (isFirstRow && useFirstRow)
                    {
                        r = table.Rows.Last();
                        foreach (var paragraph in r.Paragraphs)
                        {
                            if (!String.IsNullOrEmpty(paragraph.Text))
                            {
                                paragraph.RemoveText(0);
                            };
                        }
                    }
                    else
                    {
                        r = table.InsertRow();
                    }

                    if (tokens.Contains(TOKENS.WITH_COUNTER))
                    {
                        var p = r.Cells[0].Paragraphs.First();
                        if (p != null)
                            p.Append(counter.ToString());
                    }

                    if (tokens.Contains(TOKENS.AUTO_EXPAND))
                    {
                        while (r.ColumnCount < reader.FieldCount + iOffset)
                        {
                            table.InsertColumn(table.ColumnCount - 1);
                        }
                    }

                    var maxValue = Math.Min(reader.FieldCount + iOffset, r.ColumnCount);
                    for (int i = iOffset; i < maxValue; i++)
                    {
                        var p = r.Cells[i].Paragraphs.First();
                        if (p != null)
                            p.Append(reader.GetSqlValue(i - iOffset).ToString());
                    }

                    if (isFirstRow)
                    {
                        isFirstRow = false;
                    }

                    counter++;
                }
            } while (reader.NextResult());

            reader.Close();

            return !isFirstRow;
        }

        private static void _addScriptRow(Table table, 
            String script, 
            int contentStart, 
            Dictionary<TOKENS, string> specialParameters)
        {
            script = Regex.Replace(script, "--.*", "").Replace("\n", " "); // remove oneline SQL-comments
            if (specialParameters != null)
            {
                var tokenString = getTokensString(specialParameters);
                if (!specialParameters.ContainsKey(TOKENS.UPDATE_SCRIPT))
                {
                    tokenString = "[UPDATE_SCRIPT:" + contentStart.ToString() + "]" + tokenString;
                }
                script = "/*" + tokenString + "*/" + script;
            }

            var row = table.InsertRow();
            row.MergeCells(0, row.ColumnCount - 1);
            row.Height = 1;

            // скрипт вставляем в последнюю ячейку с невидимым текстом, без полей и высотой 1 чего-то там
            var cell = row.Cells.FirstOrDefault();
            if (cell != null)
            {
                var border = new Border { Tcbs = BorderStyle.Tcbs_none };
                cell.SetBorder(TableCellBorderType.Bottom, border);
                cell.SetBorder(TableCellBorderType.Left, border);
                cell.SetBorder(TableCellBorderType.Right, border);
            }

            row.Paragraphs.ForEach(paragraph => paragraph.Remove(false));
            var p2 = row.Paragraphs.FirstOrDefault();
            p2.Append(script);
            p2.Hide();
        }

        /// <summary>
        /// Fill table with script result, returns true on success, false otherwise
        /// </summary>
        /// <param name="table"></param>
        /// <param name="connection"></param>
        /// <param name="parameters"></param>
        /// <returns></returns>
        public static bool FillTable(
            Table table, SqlConnection connection, 
            Dictionary<string, string> parameters,
            bool saveQueries)
        {
            Dictionary<TOKENS, string> specialParameters;
            var tokens = GetTableTokens(table, out specialParameters);
            if (tokens.Count <= 0)
            {
                return false;
            }

            var dynamicScript = new Script("", GetTableScript(table));
            var script = dynamicScript.Text(parameters);

            var contentRow = GetTableScriptRow(table);
            var contentStart = table.RowCount;
            var hasRows = _fillTable(table, connection, script, tokens);

            // if no values to output
            if (!hasRows)
            {
                if (tokens.Contains(TOKENS.REPLACE_TABLE_ON_EMPTY))
                {
                    table.InsertParagraphBeforeSelf(specialParameters.Default(
                        TOKENS.REPLACE_TABLE_ON_EMPTY, ""));
                    table.Remove();
                }
            }

//            if (hasRows)
//            {
                // удаляем строку со скриптом
                contentRow.Remove();
//            }

            if (saveQueries && !String.IsNullOrEmpty(script))
            {
                // сохраняем запрос
                _addScriptRow(table, script, contentStart, specialParameters);
            }

            return true;
        }

        private static String getTokensString(Dictionary<TOKENS, string> specail_parameters)
        {
            StringBuilder stringBuilder = new StringBuilder();
            foreach (var parameter in specail_parameters)
            {
                stringBuilder.Append("[");
                stringBuilder.Append(parameter.Key);
                if (String.IsNullOrEmpty(parameter.Value))
                {
                    stringBuilder.Append(parameter.Value);
                }
                stringBuilder.Append("]");
            }
            return stringBuilder.ToString();
        }

        public static void UpdateTable(
            Table table, SqlConnection connection,
            Dictionary<string, string> parameters,
            bool saveQueries)
        {
            // скрипт должен быть в последней строке
            var row = table.Rows.Last();
            
            var script = GetCellContent(row.Cells.FirstOrDefault(), 0);
            row.Remove();

            Dictionary<TOKENS, string> specailParameters;
            var tokens = GetCellTokens(row.Cells.FirstOrDefault(), out specailParameters);

            if (!tokens.Contains(TOKENS.UPDATE_SCRIPT))
                return;

            int firstDataRow;
            if (!Int32.TryParse(specailParameters[TOKENS.UPDATE_SCRIPT], out firstDataRow))
                return;

            while (table.RowCount  > 1 && table.RowCount >= firstDataRow)
            {
                table.Rows.Last().Remove();
            }

            var dynamicScript = new Script("", script);
            script = dynamicScript.Text(parameters);

            var contentStart = table.RowCount;

            _fillTable(table, connection, script, tokens, table.RowCount == 1 && firstDataRow == 1);

            if (saveQueries)
            {
                _addScriptRow(table, script, 
                    firstDataRow,
                    null);
            }
        }
    }
}
