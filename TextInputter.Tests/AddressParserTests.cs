using TextInputter.Services;
using Xunit;
using System.Reflection;

namespace TextInputter.Tests;

public class AddressParserTests
{
    [Theory]
    [InlineData("123 Đường D1, Long Thạnh Mỹ, Thủ Đức", "9")]
    [InlineData("12 Đường Số 2, Thủ Thiêm, Q2", "2")]
    [InlineData("99 An Phú, Q2", "2")]
    [InlineData("S503 Vinhome Grand Park, đường Nguyễn Xiển, Long Thạnh Mỹ, Quận Thủ Đức, TP.HCM", "9")]
    [InlineData("s503 vinhome gradpark đường nguyễn xiển Long thạnh mỹ Quận Thủ đức Tp Thủ Đức Hồ Chí Minh", "9")]
    [InlineData("Căn 11, khu phố 1, Thủ Thiêm, Q2, TP.HCM", "2")]
    [InlineData("123 Đường Lê Văn Khương, xã Đức Nhuận, Củ Chi, HCM", "cu chi")]
    [InlineData("123 Đường A, Phường Đức Nhuận, Phú Nhuận", "phu nhuan")]
    public void Parse_ShouldPreferWardMappingOverGenericDistrictName(string rawAddress, string expectedQuan)
    {
        var parsed = AddressParser.Parse(rawAddress);
        Assert.Equal(expectedQuan, parsed.Quan);
    }

    [Theory]
    [InlineData("COD", 1545, 25, 1520)]
    [InlineData("SHIP_ONLY_PAID", 0, 25, 25)]
    [InlineData("SHIP_ONLY_FREE", 0, 25, -25)]
    public void ComputeTienHang_ShouldFollowBusinessRule(string invoiceType, long tienThu, long tienShip, long expectedTienHang)
    {
        var result = OCRTextParsingService.ComputeTienHang(invoiceType, tienThu, tienShip);
        Assert.Equal(expectedTienHang, result);
    }

    [Fact]
    public void StripDistrictAndWard_ShouldKeepWardInAddress()
    {
        var method = typeof(OCRTextParsingService).GetMethod(
            "StripDistrictAndWard",
            BindingFlags.NonPublic | BindingFlags.Static
        );

        var result = (string)method.Invoke(null, new object[] { "123 Nguyễn Văn A, Phường 3, Quận 5" });

        Assert.Contains("Phường 3", result);
        Assert.DoesNotContain("Quận 5", result);
    }
}
