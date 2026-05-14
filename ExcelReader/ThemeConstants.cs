namespace ExcelReader;

/// <summary>
/// Centralized theme colors, fonts, and layout constants.
/// Fonts are cached statically for the application lifetime.
/// </summary>
public static class ThemeConstants
{
    // ── Primary palette ────────────────────────────────
    public static readonly Color Primary = Color.FromArgb(44, 95, 138);
    public static readonly Color PrimaryDark = Color.FromArgb(32, 75, 115);
    public static readonly Color PrimaryLight = Color.FromArgb(185, 207, 234);

    // ── Backgrounds ────────────────────────────────────
    public static readonly Color BackgroundWhite = Color.White;
    public static readonly Color BackgroundLight = Color.FromArgb(248, 248, 248);
    public static readonly Color BackgroundAltRow = Color.FromArgb(248, 249, 250);
    public static readonly Color BackgroundPaper = Color.FromArgb(234, 229, 222);

    // ── Text ───────────────────────────────────────────
    public static readonly Color TextPrimary = Color.FromArgb(50, 50, 50);
    public static readonly Color TextSecondary = Color.FromArgb(100, 100, 100);
    public static readonly Color TextMuted = Color.FromArgb(180, 180, 180);
    public static readonly Color TextOnPrimary = Color.White;
    public static readonly Color TextTableBorder = Color.FromArgb(80, 80, 80);

    // ── Selection ──────────────────────────────────────
    public static readonly Color SelectionBack = Color.FromArgb(225, 238, 250);
    public static readonly Color SelectionFore = Color.FromArgb(50, 50, 50);

    // ── Grid ───────────────────────────────────────────
    public static readonly Color GridLine = Color.FromArgb(230, 230, 230);

    // ── Buttons ────────────────────────────────────────
    public static readonly Color ButtonSave = Color.FromArgb(46, 125, 50);

    // ── Fonts (cached for the app lifetime) ────────────

    private const string FontFamily = "Microsoft YaHei UI";
    private const string FontFamilyMono = "Consolas";

    public static readonly Font FontTitle = new(FontFamily, 13f, FontStyle.Bold);
    public static readonly Font FontSubtitle = new(FontFamily, 9.75f);
    public static readonly Font FontBody = new(FontFamily, 9.75f);
    public static readonly Font FontBodyBold = new(FontFamily, 9.75f, FontStyle.Bold);
    public static readonly Font FontStatus = new(FontFamily, 9f);
    public static readonly Font FontHeading1 = new(FontFamily, 18f, FontStyle.Bold);
    public static readonly Font FontHeading2 = new(FontFamily, 14f, FontStyle.Bold);
    public static readonly Font FontHeading3 = new(FontFamily, 12f, FontStyle.Bold);
    public static readonly Font FontParagraph = new(FontFamily, 11f);
    public static readonly Font FontParagraphBold = new(FontFamily, 11f, FontStyle.Bold);
    public static readonly Font FontPageNumber = new(FontFamily, 9.75f, FontStyle.Bold);
    public static readonly Font FontTableMono = new(FontFamilyMono, 9.75f);

    // ── Layout ─────────────────────────────────────────
    public const int TitleBarHeight = 56;
    public const int ToolbarHeight = 50;
    public const int ColumnHeaderHeight = 38;
    public const int RowHeight = 32;
    public const int RowHeaderWidth = 55;
    public const int CellPadding = 10;
    public const int MaxColumnPixelWidth = 400;

    // ── Limits ─────────────────────────────────────────
    public const int LargeSheetRowThreshold = 100_000;
}
