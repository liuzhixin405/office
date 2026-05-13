namespace ExcelReader;

internal class ModernColorTable : ProfessionalColorTable
{
    private static readonly Color Primary = Color.FromArgb(43, 87, 154);
    private static readonly Color Hover = Color.FromArgb(220, 235, 252);
    private static readonly Color Pressed = Color.FromArgb(185, 210, 240);

    public override Color MenuItemSelected => Hover;
    public override Color MenuItemBorder => Color.Transparent;
    public override Color MenuBorder => Color.FromArgb(210, 210, 210);
    public override Color MenuItemPressedGradientBegin => Pressed;
    public override Color MenuItemPressedGradientEnd => Pressed;
    public override Color MenuStripGradientBegin => Color.White;
    public override Color MenuStripGradientEnd => Color.White;
    public override Color ToolStripBorder => Color.FromArgb(230, 230, 230);
    public override Color ToolStripDropDownBackground => Color.White;
    public override Color ToolStripGradientBegin => Color.White;
    public override Color ToolStripGradientEnd => Color.White;
    public override Color ToolStripGradientMiddle => Color.White;
    public override Color ImageMarginGradientBegin => Color.White;
    public override Color ImageMarginGradientEnd => Color.White;
    public override Color ButtonSelectedHighlight => Hover;
    public override Color ButtonSelectedHighlightBorder => Primary;
    public override Color ButtonPressedHighlight => Pressed;
    public override Color ButtonPressedHighlightBorder => Primary;
    public override Color ButtonSelectedGradientBegin => Hover;
    public override Color ButtonSelectedGradientEnd => Hover;
    public override Color ButtonPressedGradientBegin => Pressed;
    public override Color ButtonPressedGradientEnd => Pressed;
    public override Color ButtonCheckedGradientBegin => Pressed;
    public override Color ButtonCheckedGradientEnd => Pressed;
    public override Color ButtonPressedBorder => Primary;
    public override Color ButtonSelectedBorder => Primary;
    public override Color ButtonCheckedHighlightBorder => Primary;
    public override Color MenuItemSelectedGradientBegin => Hover;
    public override Color MenuItemSelectedGradientEnd => Hover;
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
