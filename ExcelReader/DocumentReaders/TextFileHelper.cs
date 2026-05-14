using System.Text;

namespace ExcelReader;

internal static class TextFileHelper
{
    private static readonly Encoding Utf8NoBom = new UTF8Encoding(false);

    public static Encoding DetectEncoding(string path)
    {
        var bytes = File.ReadAllBytes(path);
        if (bytes.Length >= 3 && bytes[0] == 0xEF && bytes[1] == 0xBB && bytes[2] == 0xBF)
            return new UTF8Encoding(true);

        if (bytes.Length >= 2)
        {
            if (bytes[0] == 0xFF && bytes[1] == 0xFE)
                return Encoding.Unicode;
            if (bytes[0] == 0xFE && bytes[1] == 0xFF)
                return Encoding.BigEndianUnicode;
        }

        if (bytes.Length >= 4)
        {
            if (bytes[0] == 0xFF && bytes[1] == 0xFE && bytes[2] == 0x00 && bytes[3] == 0x00)
                return Encoding.UTF32;
        }

        if (LooksLikeUtf8(bytes))
            return Utf8NoBom;

        try
        {
            return Encoding.GetEncoding("GB18030");
        }
        catch
        {
            return Encoding.Default;
        }
    }

    public static string GetEncodingDisplayName(Encoding encoding)
    {
        return encoding.WebName.ToUpperInvariant();
    }

    private static bool LooksLikeUtf8(byte[] bytes)
    {
        int i = 0;
        while (i < bytes.Length)
        {
            byte b = bytes[i];
            if (b <= 0x7F)
            {
                i++;
                continue;
            }

            int additionalBytes = b switch
            {
                >= 0xC2 and <= 0xDF => 1,
                >= 0xE0 and <= 0xEF => 2,
                >= 0xF0 and <= 0xF4 => 3,
                _ => -1
            };

            if (additionalBytes < 0 || i + additionalBytes >= bytes.Length)
                return false;

            for (int j = 1; j <= additionalBytes; j++)
            {
                if ((bytes[i + j] & 0xC0) != 0x80)
                    return false;
            }

            i += additionalBytes + 1;
        }

        return true;
    }
}
