using System;
using System.Globalization;
using System.Text;

namespace GhDucky.Utils
{
    /// <summary>
    /// Parses and formats Excel A1-style addresses.
    /// <para>
    /// Accepted forms:
    /// <list type="bullet">
    ///   <item><c>B2</c> — single cell (anchor).</item>
    ///   <item><c>B2:D10</c> — explicit range.</item>
    ///   <item><c>Data!B2:D10</c> / <c>'My Sheet'!B2</c> — sheet-qualified.</item>
    ///   <item><c>$B$2:$D$10</c> — absolute markers are ignored.</item>
    ///   <item><c>MyRange</c> — anything that is not a valid cell reference is
    ///         treated as a defined name / table name for later resolution.</item>
    /// </list>
    /// </para>
    /// Rows and columns are 1-based, matching Excel's own UI (A = 1, row 1 = 1).
    /// </summary>
    internal sealed class ExcelAddress
    {
        /// <summary>Sheet name taken from a <c>Sheet!</c> prefix, or null when absent.</summary>
        public string Sheet { get; private init; }

        /// <summary>
        /// Defined-name / table name when the address was not a cell reference.
        /// When set, the row/column properties are meaningless until the name
        /// has been resolved against a workbook.
        /// </summary>
        public string Name { get; private init; }

        public int StartRow { get; private init; }
        public int StartColumn { get; private init; }

        /// <summary>Null when the address is an open-ended anchor (e.g. <c>B2</c>).</summary>
        public int? EndRow { get; private init; }

        /// <summary>Null when the address is an open-ended anchor (e.g. <c>B2</c>).</summary>
        public int? EndColumn { get; private init; }

        public bool IsName => Name != null;

        /// <summary>True when only an anchor cell was supplied, so the extent is open.</summary>
        public bool IsAnchorOnly => !IsName && (EndRow is null || EndColumn is null);

        /// <summary>The largest column index Excel supports (XFD).</summary>
        public const int MaxColumn = 16384;

        /// <summary>The largest row index Excel supports.</summary>
        public const int MaxRow = 1048576;

        private ExcelAddress() { }

        /// <summary>
        /// Attempts to parse an address. Returns false with a human-readable
        /// <paramref name="error"/> when the text cannot be interpreted.
        /// </summary>
        public static bool TryParse(string text, out ExcelAddress address, out string error)
        {
            address = null;
            error = null;

            if (string.IsNullOrWhiteSpace(text))
            {
                error = "Address is empty.";
                return false;
            }

            var trimmed = text.Trim();

            // ---- split off an optional "Sheet!" prefix ------------------------
            string sheet = null;
            var refPart = trimmed;

            var bang = LastSheetSeparator(trimmed);
            if (bang >= 0)
            {
                sheet = UnquoteSheet(trimmed.Substring(0, bang));
                refPart = trimmed.Substring(bang + 1).Trim();

                if (string.IsNullOrWhiteSpace(sheet))
                {
                    error = "Sheet name in address is empty.";
                    return false;
                }
                if (string.IsNullOrWhiteSpace(refPart))
                {
                    error = "Cell reference after '!' is empty.";
                    return false;
                }
            }

            // ---- range or single cell ---------------------------------------
            var colon = refPart.IndexOf(':');
            if (colon < 0)
            {
                if (!TryParseCell(refPart, out var row, out var col))
                {
                    // Not a cell reference — treat as a defined name / table name.
                    // A sheet-qualified non-reference is not a legal name, so reject it.
                    if (sheet != null)
                    {
                        error = $"'{refPart}' is not a valid cell reference.";
                        return false;
                    }

                    if (!IsPlausibleName(refPart))
                    {
                        error = $"'{refPart}' is neither a valid cell reference nor a valid name.";
                        return false;
                    }

                    address = new ExcelAddress { Name = refPart };
                    return true;
                }

                address = new ExcelAddress
                {
                    Sheet = sheet,
                    StartRow = row,
                    StartColumn = col,
                    EndRow = null,
                    EndColumn = null,
                };
                return true;
            }

            var left = refPart.Substring(0, colon).Trim();
            var right = refPart.Substring(colon + 1).Trim();

            if (!TryParseCell(left, out var r1, out var c1))
            {
                error = $"'{left}' is not a valid cell reference.";
                return false;
            }
            if (!TryParseCell(right, out var r2, out var c2))
            {
                error = $"'{right}' is not a valid cell reference.";
                return false;
            }

            // Normalise reversed ranges (D10:B2 is legal in Excel).
            address = new ExcelAddress
            {
                Sheet = sheet,
                StartRow = Math.Min(r1, r2),
                StartColumn = Math.Min(c1, c2),
                EndRow = Math.Max(r1, r2),
                EndColumn = Math.Max(c1, c2),
            };
            return true;
        }

        /// <summary>Returns a copy with the sheet name replaced.</summary>
        public ExcelAddress WithSheet(string sheet) => new()
        {
            Sheet = string.IsNullOrWhiteSpace(sheet) ? Sheet : sheet,
            Name = Name,
            StartRow = StartRow,
            StartColumn = StartColumn,
            EndRow = EndRow,
            EndColumn = EndColumn,
        };

        /// <summary>Returns a copy with an explicit extent applied.</summary>
        public ExcelAddress WithExtent(int endRow, int endColumn) => new()
        {
            Sheet = Sheet,
            Name = null,
            StartRow = StartRow,
            StartColumn = StartColumn,
            EndRow = endRow,
            EndColumn = endColumn,
        };

        /// <summary>
        /// Builds a concrete address from resolved 1-based bounds, e.g.
        /// <c>Data!B2:D11</c>. The sheet is quoted when it needs to be.
        /// </summary>
        public static string Format(string sheet, int startRow, int startColumn, int endRow, int endColumn)
        {
            var sb = new StringBuilder();
            if (!string.IsNullOrWhiteSpace(sheet))
                sb.Append(QuoteSheet(sheet)).Append('!');

            sb.Append(Cell(startRow, startColumn));
            if (startRow != endRow || startColumn != endColumn)
                sb.Append(':').Append(Cell(endRow, endColumn));

            return sb.ToString();
        }

        /// <summary>Formats a 1-based row/column pair as <c>B2</c>.</summary>
        public static string Cell(int row, int column) =>
            ColumnName(column) + row.ToString(CultureInfo.InvariantCulture);

        /// <summary>Converts a 1-based column index to letters (1 → A, 27 → AA).</summary>
        public static string ColumnName(int column)
        {
            if (column < 1)
                throw new ArgumentOutOfRangeException(nameof(column), column, "Column index is 1-based.");

            var sb = new StringBuilder(3);
            while (column > 0)
            {
                var rem = (column - 1) % 26;
                sb.Insert(0, (char)('A' + rem));
                column = (column - 1) / 26;
            }
            return sb.ToString();
        }

        /// <summary>Converts column letters to a 1-based index (A → 1, AA → 27).</summary>
        public static int ColumnIndex(string columnName)
        {
            if (string.IsNullOrWhiteSpace(columnName))
                throw new ArgumentException("Column name is required.", nameof(columnName));

            var index = 0;
            foreach (var ch in columnName)
            {
                var upper = char.ToUpperInvariant(ch);
                if (upper is < 'A' or > 'Z')
                    throw new ArgumentException($"'{columnName}' is not a column name.", nameof(columnName));

                index = index * 26 + (upper - 'A' + 1);
            }
            return index;
        }

        /// <summary>Parses a single cell reference such as <c>$B$2</c>.</summary>
        public static bool TryParseCell(string text, out int row, out int column)
        {
            row = 0;
            column = 0;
            if (string.IsNullOrWhiteSpace(text)) return false;

            var i = 0;
            var s = text.Trim();

            if (i < s.Length && s[i] == '$') i++;

            var letterStart = i;
            while (i < s.Length && char.IsLetter(s[i])) i++;
            var letterCount = i - letterStart;
            if (letterCount is < 1 or > 3) return false;

            if (i < s.Length && s[i] == '$') i++;

            var digitStart = i;
            while (i < s.Length && char.IsDigit(s[i])) i++;
            if (i - digitStart < 1) return false;

            // Trailing junk means this is not a cell reference.
            if (i != s.Length) return false;

            var letters = s.Substring(letterStart, letterCount);
            foreach (var ch in letters)
            {
                var upper = char.ToUpperInvariant(ch);
                if (upper is < 'A' or > 'Z') return false;
            }

            if (!int.TryParse(s.Substring(digitStart), NumberStyles.None, CultureInfo.InvariantCulture, out row))
                return false;

            column = ColumnIndex(letters);
            return row is >= 1 and <= MaxRow && column is >= 1 and <= MaxColumn;
        }

        /// <summary>
        /// Finds the sheet/reference separator, ignoring '!' inside a quoted
        /// sheet name. Returns -1 when there is no prefix.
        /// </summary>
        private static int LastSheetSeparator(string s)
        {
            var inQuotes = false;
            var last = -1;
            for (var i = 0; i < s.Length; i++)
            {
                if (s[i] == '\'') inQuotes = !inQuotes;
                else if (s[i] == '!' && !inQuotes) last = i;
            }
            return last;
        }

        private static string UnquoteSheet(string s)
        {
            var t = s.Trim();
            if (t.Length >= 2 && t[0] == '\'' && t[^1] == '\'')
                t = t.Substring(1, t.Length - 2).Replace("''", "'");
            return t.Trim();
        }

        private static string QuoteSheet(string s) =>
            s.IndexOfAny(new[] { ' ', '\'', '!', '-', '(', ')' }) >= 0
                ? "'" + s.Replace("'", "''") + "'"
                : s;

        /// <summary>
        /// Defined names must start with a letter or underscore and contain no
        /// spaces. This is a light check purely to give better error messages.
        /// </summary>
        private static bool IsPlausibleName(string s)
        {
            if (s.Length == 0) return false;
            if (!char.IsLetter(s[0]) && s[0] != '_' && s[0] != '\\') return false;

            foreach (var ch in s)
            {
                if (!char.IsLetterOrDigit(ch) && ch != '_' && ch != '.' && ch != '\\')
                    return false;
            }
            return true;
        }
    }
}

