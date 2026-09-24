using System;
using GhDucky.Utils;
using Xunit;

namespace GhDucky.Tests
{
    public class ExcelAddressTests
    {
        [Theory]
        [InlineData(1, "A")]
        [InlineData(2, "B")]
        [InlineData(26, "Z")]
        [InlineData(27, "AA")]
        [InlineData(28, "AB")]
        [InlineData(52, "AZ")]
        [InlineData(53, "BA")]
        [InlineData(702, "ZZ")]
        [InlineData(703, "AAA")]
        [InlineData(16384, "XFD")]
        public void ColumnName_MatchesExcel(int index, string expected) =>
            Assert.Equal(expected, ExcelAddress.ColumnName(index));

        [Theory]
        [InlineData("A", 1)]
        [InlineData("B", 2)]
        [InlineData("Z", 26)]
        [InlineData("AA", 27)]
        [InlineData("XFD", 16384)]
        [InlineData("aa", 27)]
        public void ColumnIndex_IsInverseOfColumnName(string name, int expected) =>
            Assert.Equal(expected, ExcelAddress.ColumnIndex(name));

        [Fact]
        public void ColumnName_RejectsZeroAndNegative()
        {
            Assert.Throws<ArgumentOutOfRangeException>(() => ExcelAddress.ColumnName(0));
            Assert.Throws<ArgumentOutOfRangeException>(() => ExcelAddress.ColumnName(-1));
        }

        [Theory]
        [InlineData("A1", 1, 1)]
        [InlineData("B2", 2, 2)]
        [InlineData("$B$2", 2, 2)]
        [InlineData("AA10", 10, 27)]
        [InlineData("xfd1048576", 1048576, 16384)]
        public void TryParseCell_AcceptsValidReferences(string text, int row, int column)
        {
            Assert.True(ExcelAddress.TryParseCell(text, out var r, out var c));
            Assert.Equal(row, r);
            Assert.Equal(column, c);
        }

        [Theory]
        [InlineData("")]
        [InlineData("1A")]
        [InlineData("A")]
        [InlineData("1")]
        [InlineData("ABCD1")]
        [InlineData("A1B")]
        [InlineData("A0")]
        public void TryParseCell_RejectsInvalidReferences(string text) =>
            Assert.False(ExcelAddress.TryParseCell(text, out _, out _));

        [Fact]
        public void TryParse_SingleCell_IsAnchorOnly()
        {
            Assert.True(ExcelAddress.TryParse("B2", out var address, out _));
            Assert.Equal(2, address.StartRow);
            Assert.Equal(2, address.StartColumn);
            Assert.Null(address.EndRow);
            Assert.Null(address.EndColumn);
            Assert.True(address.IsAnchorOnly);
            Assert.Null(address.Sheet);
        }

        [Fact]
        public void TryParse_Range_CapturesExtent()
        {
            Assert.True(ExcelAddress.TryParse("B2:D10", out var address, out _));
            Assert.Equal(2, address.StartRow);
            Assert.Equal(2, address.StartColumn);
            Assert.Equal(10, address.EndRow);
            Assert.Equal(4, address.EndColumn);
            Assert.False(address.IsAnchorOnly);
        }

        [Fact]
        public void TryParse_ReversedRange_IsNormalised()
        {
            Assert.True(ExcelAddress.TryParse("D10:B2", out var address, out _));
            Assert.Equal(2, address.StartRow);
            Assert.Equal(2, address.StartColumn);
            Assert.Equal(10, address.EndRow);
            Assert.Equal(4, address.EndColumn);
        }

        [Theory]
        [InlineData("Data!B2:D10", "Data")]
        [InlineData("'My Sheet'!B2", "My Sheet")]
        [InlineData("'It''s Here'!A1", "It's Here")]
        public void TryParse_ExtractsSheetPrefix(string text, string sheet)
        {
            Assert.True(ExcelAddress.TryParse(text, out var address, out _));
            Assert.Equal(sheet, address.Sheet);
        }

        [Fact]
        public void TryParse_AbsoluteSheetQualifiedRange()
        {
            Assert.True(ExcelAddress.TryParse("Data!$B$2:$C$3", out var address, out _));
            Assert.Equal("Data", address.Sheet);
            Assert.Equal(2, address.StartRow);
            Assert.Equal(2, address.StartColumn);
            Assert.Equal(3, address.EndRow);
            Assert.Equal(3, address.EndColumn);
        }

        [Fact]
        public void TryParse_NonReference_IsTreatedAsName()
        {
            Assert.True(ExcelAddress.TryParse("MyRange", out var address, out _));
            Assert.True(address.IsName);
            Assert.Equal("MyRange", address.Name);
        }

        [Fact]
        public void TryParse_SheetQualifiedNonReference_IsRejected()
        {
            Assert.False(ExcelAddress.TryParse("Data!NotACell!", out _, out var error));
            Assert.False(string.IsNullOrEmpty(error));
        }

        [Theory]
        [InlineData(null)]
        [InlineData("")]
        [InlineData("   ")]
        public void TryParse_RejectsEmpty(string text)
        {
            Assert.False(ExcelAddress.TryParse(text, out _, out var error));
            Assert.False(string.IsNullOrEmpty(error));
        }

        [Fact]
        public void Format_SingleCell_OmitsRangeSuffix() =>
            Assert.Equal("Data!B2", ExcelAddress.Format("Data", 2, 2, 2, 2));

        [Fact]
        public void Format_Range_IncludesBothCorners() =>
            Assert.Equal("Data!B2:D11", ExcelAddress.Format("Data", 2, 2, 11, 4));

        [Fact]
        public void Format_WithoutSheet_OmitsPrefix() =>
            Assert.Equal("B2:D11", ExcelAddress.Format(null, 2, 2, 11, 4));

        [Fact]
        public void Format_QuotesSheetNamesNeedingIt() =>
            Assert.Equal("'My Sheet'!A1", ExcelAddress.Format("My Sheet", 1, 1, 1, 1));

        [Fact]
        public void Format_RoundTripsThroughTryParse()
        {
            var formatted = ExcelAddress.Format("My Sheet", 2, 2, 11, 4);
            Assert.True(ExcelAddress.TryParse(formatted, out var address, out _));
            Assert.Equal("My Sheet", address.Sheet);
            Assert.Equal(2, address.StartRow);
            Assert.Equal(2, address.StartColumn);
            Assert.Equal(11, address.EndRow);
            Assert.Equal(4, address.EndColumn);
        }

        [Fact]
        public void WithSheet_FillsInMissingSheet()
        {
            Assert.True(ExcelAddress.TryParse("B2:D10", out var address, out _));
            var qualified = address.WithSheet("Data");
            Assert.Equal("Data", qualified.Sheet);
            Assert.Equal(address.StartRow, qualified.StartRow);
            Assert.Equal(address.EndColumn, qualified.EndColumn);
        }
    }
}

