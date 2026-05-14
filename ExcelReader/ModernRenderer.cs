namespace ExcelReader;

internal class ModernColorTable : ProfessionalColorTable
{
    public override Color MenuItemSelected => ThemeConstants.SelectionBack;
    public override Color MenuItemBorder => Color.Transparent;
    public override Color MenuBorder => ThemeConstants.GridLine;
    public override Color MenuItemPressedGradientBegin => ThemeConstants.SelectionBack;
    public override Color MenuItemPressedGradientEnd => ThemeConstants.SelectionBack;
    public override Color MenuStripGradientBegin => ThemeConstants.BackgroundWhite;
    public override Color MenuStripGradientEnd => ThemeConstants.BackgroundWhite;
    public override Color ToolStripBorder => ThemeConstants.GridLine;
    public override Color ToolStripDropDownBackground => ThemeConstants.BackgroundWhite;
    public override Color ToolStripGradientBegin => ThemeConstants.BackgroundWhite;
    public override Color ToolStripGradientEnd => ThemeConstants.BackgroundWhite;
    public override Color ToolStripGradientMiddle => ThemeConstants.BackgroundWhite;
    public override Color ImageMarginGradientBegin => ThemeConstants.BackgroundWhite;
    public override Color ImageMarginGradientEnd => ThemeConstants.BackgroundWhite;
    public override Color ButtonSelectedHighlight => ThemeConstants.SelectionBack;
    public override Color ButtonSelectedHighlightBorder => ThemeConstants.Primary;
    public override Color ButtonPressedHighlight => ThemeConstants.SelectionBack;
    public override Color ButtonPressedHighlightBorder => ThemeConstants.Primary;
    public override Color ButtonSelectedGradientBegin => ThemeConstants.SelectionBack;
    public override Color ButtonSelectedGradientEnd => ThemeConstants.SelectionBack;
    public override Color ButtonPressedGradientBegin => ThemeConstants.SelectionBack;
    public override Color ButtonPressedGradientEnd => ThemeConstants.SelectionBack;
    public override Color ButtonCheckedGradientBegin => ThemeConstants.SelectionBack;
    public override Color ButtonCheckedGradientEnd => ThemeConstants.SelectionBack;
    public override Color ButtonPressedBorder => ThemeConstants.Primary;
    public override Color ButtonSelectedBorder => ThemeConstants.Primary;
    public override Color ButtonCheckedHighlightBorder => ThemeConstants.Primary;
    public override Color MenuItemSelectedGradientBegin => ThemeConstants.SelectionBack;
    public override Color MenuItemSelectedGradientEnd => ThemeConstants.SelectionBack;
}

internal class ModernRenderer : ToolStripProfessionalRenderer
{
    public ModernRenderer() : base(new ModernColorTable()) { }

    protected override void OnRenderToolStripBorder(ToolStripRenderEventArgs e)
    {
        if (e.ToolStrip is not StatusStrip)
            base.OnRenderToolStripBorder(e);
    }
}