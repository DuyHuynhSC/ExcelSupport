using System;
using System.IO;
using System.Text;
using System.Text.RegularExpressions;
using System.Windows.Media;
using System.Xml;
using ICSharpCode.AvalonEdit;
using ICSharpCode.AvalonEdit.Highlighting;
using ICSharpCode.AvalonEdit.Highlighting.Xshd;
using MediaColor = System.Windows.Media.Color;
using MediaColorConverter = System.Windows.Media.ColorConverter;

namespace ExcelSupport.Services
{
    /// <summary>
    /// Cung cấp bộ tô màu cú pháp (Syntax Highlighting) chuyên biệt cho Oracle SQL trong AvalonEdit
    /// Hỗ trợ cả 2 chế độ Dark Theme và Light Theme, cùng công cụ định dạng (Beautify/Format) câu lệnh SQL.
    /// </summary>
    public static class OracleSqlSyntaxHighlighter
    {
        private static IHighlightingDefinition? _lightDefinition;
        private static IHighlightingDefinition? _darkDefinition;
        private static readonly object _lock = new object();

        public static IHighlightingDefinition LightDefinition
        {
            get
            {
                if (_lightDefinition == null)
                {
                    lock (_lock)
                    {
                        _lightDefinition ??= LoadDefinition(isDarkTheme: false);
                    }
                }
                return _lightDefinition;
            }
        }

        public static IHighlightingDefinition DarkDefinition
        {
            get
            {
                if (_darkDefinition == null)
                {
                    lock (_lock)
                    {
                        _darkDefinition ??= LoadDefinition(isDarkTheme: true);
                    }
                }
                return _darkDefinition;
            }
        }

        /// <summary>
        /// Áp dụng giao diện (Màu sắc chữ, nền, số dòng, highlight) cho AvalonEdit TextEditor theo theme hiện tại.
        /// </summary>
        public static void ApplyTheme(TextEditor editor, bool isDarkTheme)
        {
            if (editor == null) return;

            editor.SyntaxHighlighting = isDarkTheme ? DarkDefinition : LightDefinition;

            if (isDarkTheme)
            {
                // Dark Theme: Nền xanh than tối #1E293B, chữ sáng #F8FAFC
                editor.Background = new SolidColorBrush((MediaColor)MediaColorConverter.ConvertFromString("#1E293B"));
                editor.Foreground = new SolidColorBrush((MediaColor)MediaColorConverter.ConvertFromString("#F8FAFC"));
                editor.LineNumbersForeground = new SolidColorBrush((MediaColor)MediaColorConverter.ConvertFromString("#64748B"));
                if (editor.TextArea?.Caret != null)
                {
                    editor.TextArea.Caret.CaretBrush = new SolidColorBrush((MediaColor)MediaColorConverter.ConvertFromString("#38BDF8"));
                }
                if (editor.TextArea != null)
                {
                    editor.TextArea.SelectionBrush = new SolidColorBrush(MediaColor.FromArgb(120, 56, 189, 248));
                    editor.TextArea.SelectionBorder = null;
                    editor.TextArea.SelectionForeground = null;
                }
            }
            else
            {
                // Light Theme: Nền trắng sạch #FFFFFF, chữ đen đậm #0F172A
                editor.Background = new SolidColorBrush((MediaColor)MediaColorConverter.ConvertFromString("#FFFFFF"));
                editor.Foreground = new SolidColorBrush((MediaColor)MediaColorConverter.ConvertFromString("#0F172A"));
                editor.LineNumbersForeground = new SolidColorBrush((MediaColor)MediaColorConverter.ConvertFromString("#94A3B8"));
                if (editor.TextArea?.Caret != null)
                {
                    editor.TextArea.Caret.CaretBrush = new SolidColorBrush((MediaColor)MediaColorConverter.ConvertFromString("#0F172A"));
                }
                if (editor.TextArea != null)
                {
                    editor.TextArea.SelectionBrush = new SolidColorBrush(MediaColor.FromArgb(100, 37, 99, 235));
                    editor.TextArea.SelectionBorder = null;
                    editor.TextArea.SelectionForeground = null;
                }
            }
        }

        private static IHighlightingDefinition LoadDefinition(bool isDarkTheme)
        {
            string commentColor = isDarkTheme ? "#94A3B8" : "#64748B";
            string stringColor = isDarkTheme ? "#FBBF24" : "#B45309";
            string keywordColor = isDarkTheme ? "#38BDF8" : "#005CC5";
            string functionColor = isDarkTheme ? "#C084FC" : "#6F42C1";
            string dataTypeColor = isDarkTheme ? "#22D3EE" : "#0969DA";
            string numberColor = isDarkTheme ? "#34D399" : "#059669";
            string operatorColor = isDarkTheme ? "#CBD5E1" : "#475569";
            string defName = isDarkTheme ? "OracleSQL_Dark" : "OracleSQL_Light";

            string xshd = $@"<?xml version=""1.0""?>
<SyntaxDefinition name=""{defName}"" xmlns=""http://icsharpcode.net/sharpdevelop/syntaxdefinition/2008"">
    <Color name=""Comment"" foreground=""{commentColor}"" fontStyle=""italic"" />
    <Color name=""String"" foreground=""{stringColor}"" />
    <Color name=""Keyword"" foreground=""{keywordColor}"" fontWeight=""bold"" />
    <Color name=""Function"" foreground=""{functionColor}"" fontWeight=""bold"" />
    <Color name=""DataType"" foreground=""{dataTypeColor}"" />
    <Color name=""Number"" foreground=""{numberColor}"" />
    <Color name=""Operator"" foreground=""{operatorColor}"" />

    <RuleSet ignoreCase=""true"">
        <!-- Multi-line comments -->
        <Span color=""Comment"" multiline=""true"">
            <Begin>/\*</Begin>
            <End>\*/</End>
        </Span>

        <!-- Single-line comments -->
        <Span color=""Comment"">
            <Begin>--</Begin>
        </Span>

        <!-- String literals (with escaped quotes '') -->
        <Span color=""String"">
            <Begin>'</Begin>
            <End>'</End>
            <RuleSet>
                <Span begin=""''"" end="""" />
            </RuleSet>
        </Span>

        <!-- Numbers -->
        <Rule color=""Number"">
            \b0[xX][0-9a-fA-F]+|(\b\d+(\.[0-9]+)?|\.[0-9]+)([eE][+-]?[0-9]+)?\b
        </Rule>

        <!-- Keywords -->
        <Keywords color=""Keyword"">
            <Word>SELECT</Word>
            <Word>FROM</Word>
            <Word>WHERE</Word>
            <Word>AND</Word>
            <Word>OR</Word>
            <Word>NOT</Word>
            <Word>IN</Word>
            <Word>LIKE</Word>
            <Word>IS</Word>
            <Word>NULL</Word>
            <Word>BETWEEN</Word>
            <Word>EXISTS</Word>
            <Word>AS</Word>
            <Word>JOIN</Word>
            <Word>INNER</Word>
            <Word>LEFT</Word>
            <Word>RIGHT</Word>
            <Word>FULL</Word>
            <Word>OUTER</Word>
            <Word>CROSS</Word>
            <Word>ON</Word>
            <Word>GROUP</Word>
            <Word>BY</Word>
            <Word>HAVING</Word>
            <Word>ORDER</Word>
            <Word>ASC</Word>
            <Word>DESC</Word>
            <Word>NULLS</Word>
            <Word>FIRST</Word>
            <Word>LAST</Word>
            <Word>INSERT</Word>
            <Word>INTO</Word>
            <Word>VALUES</Word>
            <Word>UPDATE</Word>
            <Word>SET</Word>
            <Word>DELETE</Word>
            <Word>MERGE</Word>
            <Word>USING</Word>
            <Word>MATCHED</Word>
            <Word>CREATE</Word>
            <Word>ALTER</Word>
            <Word>DROP</Word>
            <Word>TRUNCATE</Word>
            <Word>TABLE</Word>
            <Word>VIEW</Word>
            <Word>INDEX</Word>
            <Word>SYNONYM</Word>
            <Word>SEQUENCE</Word>
            <Word>UNION</Word>
            <Word>ALL</Word>
            <Word>INTERSECT</Word>
            <Word>MINUS</Word>
            <Word>DISTINCT</Word>
            <Word>UNIQUE</Word>
            <Word>CASE</Word>
            <Word>WHEN</Word>
            <Word>THEN</Word>
            <Word>ELSE</Word>
            <Word>END</Word>
            <Word>WITH</Word>
            <Word>OVER</Word>
            <Word>PARTITION</Word>
            <Word>ROWNUM</Word>
            <Word>ROWID</Word>
            <Word>FETCH</Word>
            <Word>NEXT</Word>
            <Word>ROWS</Word>
            <Word>ONLY</Word>
            <Word>OFFSET</Word>
            <Word>DUAL</Word>
            <Word>COMMIT</Word>
            <Word>ROLLBACK</Word>
            <Word>GRANT</Word>
            <Word>REVOKE</Word>
            <Word>PRIMARY</Word>
            <Word>FOREIGN</Word>
            <Word>KEY</Word>
            <Word>CONSTRAINT</Word>
            <Word>REFERENCES</Word>
            <Word>CHECK</Word>
            <Word>DEFAULT</Word>
        </Keywords>

        <!-- Oracle Built-in Functions -->
        <Keywords color=""Function"">
            <Word>NVL</Word>
            <Word>NVL2</Word>
            <Word>COALESCE</Word>
            <Word>DECODE</Word>
            <Word>NULLIF</Word>
            <Word>TO_CHAR</Word>
            <Word>TO_DATE</Word>
            <Word>TO_NUMBER</Word>
            <Word>TO_TIMESTAMP</Word>
            <Word>SUBSTR</Word>
            <Word>INSTR</Word>
            <Word>LENGTH</Word>
            <Word>TRIM</Word>
            <Word>LTRIM</Word>
            <Word>RTRIM</Word>
            <Word>UPPER</Word>
            <Word>LOWER</Word>
            <Word>INITCAP</Word>
            <Word>REPLACE</Word>
            <Word>TRANSLATE</Word>
            <Word>LPAD</Word>
            <Word>RPAD</Word>
            <Word>CONCAT</Word>
            <Word>REGEXP_LIKE</Word>
            <Word>REGEXP_SUBSTR</Word>
            <Word>REGEXP_INSTR</Word>
            <Word>REGEXP_REPLACE</Word>
            <Word>ROUND</Word>
            <Word>TRUNC</Word>
            <Word>MOD</Word>
            <Word>CEIL</Word>
            <Word>FLOOR</Word>
            <Word>ABS</Word>
            <Word>POWER</Word>
            <Word>SQRT</Word>
            <Word>COUNT</Word>
            <Word>SUM</Word>
            <Word>AVG</Word>
            <Word>MIN</Word>
            <Word>MAX</Word>
            <Word>MEDIAN</Word>
            <Word>STDDEV</Word>
            <Word>LISTAGG</Word>
            <Word>WM_CONCAT</Word>
            <Word>XMLAGG</Word>
            <Word>ROW_NUMBER</Word>
            <Word>DENSE_RANK</Word>
            <Word>RANK</Word>
            <Word>LEAD</Word>
            <Word>LAG</Word>
            <Word>FIRST_VALUE</Word>
            <Word>LAST_VALUE</Word>
            <Word>ADD_MONTHS</Word>
            <Word>MONTHS_BETWEEN</Word>
            <Word>LAST_DAY</Word>
            <Word>NEXT_DAY</Word>
            <Word>SYSDATE</Word>
            <Word>SYSTIMESTAMP</Word>
            <Word>CURRENT_DATE</Word>
            <Word>CURRENT_TIMESTAMP</Word>
            <Word>SYS_GUID</Word>
            <Word>DUMP</Word>
            <Word>VSIZE</Word>
        </Keywords>

        <!-- Oracle Data Types -->
        <Keywords color=""DataType"">
            <Word>VARCHAR2</Word>
            <Word>NVARCHAR2</Word>
            <Word>CHAR</Word>
            <Word>NCHAR</Word>
            <Word>NUMBER</Word>
            <Word>INTEGER</Word>
            <Word>INT</Word>
            <Word>SMALLINT</Word>
            <Word>FLOAT</Word>
            <Word>DOUBLE</Word>
            <Word>DATE</Word>
            <Word>TIMESTAMP</Word>
            <Word>CLOB</Word>
            <Word>NCLOB</Word>
            <Word>BLOB</Word>
            <Word>BFILE</Word>
            <Word>RAW</Word>
            <Word>LONG</Word>
            <Word>BOOLEAN</Word>
        </Keywords>

        <!-- Operators -->
        <Rule color=""Operator"">
            [=+\-*/&lt;&gt;!~%^&amp;|:;,\.\(\)]+
        </Rule>
    </RuleSet>
</SyntaxDefinition>";

            using var reader = XmlReader.Create(new StringReader(xshd));
            return HighlightingLoader.Load(reader, HighlightingManager.Instance);
        }

        /// <summary>
        /// Định dạng lại câu truy vấn SQL Oracle cho đẹp mắt, dễ đọc với căn lề chuẩn.
        /// </summary>
        public static string FormatSql(string sql)
        {
            if (string.IsNullOrWhiteSpace(sql)) return string.Empty;

            try
            {
                // Chuẩn hóa khoảng trắng dư thừa
                string clean = Regex.Replace(sql.Trim(), @"\s+", " ");

                // Danh sách từ khóa mệnh đề chính bắt đầu dòng mới
                var majorClauses = new[]
                {
                    "SELECT", "FROM", "WHERE", "GROUP BY", "HAVING", "ORDER BY",
                    "LEFT JOIN", "RIGHT JOIN", "INNER JOIN", "FULL JOIN", "CROSS JOIN", "JOIN",
                    "UNION ALL", "UNION", "MINUS", "INTERSECT",
                    "INSERT INTO", "UPDATE", "DELETE FROM"
                };

                // Tách theo từ khóa chính
                foreach (var kw in majorClauses)
                {
                    clean = Regex.Replace(clean, $@"(?i)\b({Regex.Escape(kw)})\b", "\n$1");
                }

                // Căn lề AND / OR bên trong WHERE
                clean = Regex.Replace(clean, @"(?i)\b(AND|OR)\b", "\n  $1");

                // Căn lề ON của JOIN
                clean = Regex.Replace(clean, @"(?i)\b(ON)\b", "\n    $1");

                // Dọn dẹp dòng trống dư thừa
                var sb = new StringBuilder();
                var lines = clean.Split(new[] { '\r', '\n' }, StringSplitOptions.RemoveEmptyEntries);
                foreach (var line in lines)
                {
                    string trimmed = line.TrimEnd();
                    if (!string.IsNullOrWhiteSpace(trimmed))
                    {
                        sb.AppendLine(trimmed);
                    }
                }

                return sb.ToString().TrimEnd();
            }
            catch
            {
                return sql;
            }
        }
    }
}
