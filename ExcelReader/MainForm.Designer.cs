namespace ExcelReader;

partial class MainForm
{
    private System.ComponentModel.IContainer components = null;

    private Panel titleBar;
    private Label lblTitle;
    private Label lblSubtitle;
    private TableLayoutPanel toolbarPanel;
    private Button btnOpen;
    private Button btnSave;
    private Label lblSheetLabel;
    private ComboBox cmbSheet;
    private DataGridView dataGrid;
    private Panel wordContainer;
    private RichTextBox wordView;
    private StatusStrip statusBar;
    private ToolStripStatusLabel statusLabel;
    private ToolStripStatusLabel statusStats;

    protected override void Dispose(bool disposing)
    {
        if (disposing && (components != null))
        {
            components.Dispose();
        }
        base.Dispose(disposing);
    }

    private void InitializeComponent()
    {
        SuspendLayout();

        // ── Title bar ──────────────────────────────────────────
        titleBar = new Panel
        {
            BackColor = ThemeConstants.Primary,
            Dock = DockStyle.Top,
            Height = ThemeConstants.TitleBarHeight,
            Padding = new Padding(24, 0, 24, 0)
        };

        lblTitle = new Label
        {
            Text = "Office 阅读器",
            ForeColor = ThemeConstants.TextOnPrimary,
            Font = ThemeConstants.FontTitle,
            Location = new Point(24, 14),
            Size = new Size(160, 26),
            BackColor = Color.Transparent
        };

        lblSubtitle = new Label
        {
            Text = "·  支持 Excel, Word, PDF, TXT, CSV",
            ForeColor = ThemeConstants.PrimaryLight,
            Font = ThemeConstants.FontSubtitle,
            Location = new Point(184, 18),
            Size = new Size(500, 20),
            BackColor = Color.Transparent
        };

        titleBar.Controls.Add(lblTitle);
        titleBar.Controls.Add(lblSubtitle);

        // ── Toolbar ────────────────────────────────────────────
        toolbarPanel = new TableLayoutPanel
        {
            Dock = DockStyle.Top,
            Height = ThemeConstants.ToolbarHeight,
            Padding = new Padding(20, 8, 20, 8),
            BackColor = ThemeConstants.BackgroundWhite,
            ColumnCount = 5,
            RowCount = 1
        };
        toolbarPanel.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, 115));
        toolbarPanel.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, 115));
        toolbarPanel.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, 20));
        toolbarPanel.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, 65));
        toolbarPanel.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100));

        btnOpen = new Button
        {
            Text = "打开文件",
            FlatStyle = FlatStyle.Flat,
            BackColor = ThemeConstants.Primary,
            ForeColor = ThemeConstants.TextOnPrimary,
            Font = ThemeConstants.FontBody,
            Size = new Size(110, 34),
            Cursor = Cursors.Hand,
            TextAlign = ContentAlignment.MiddleCenter
        };
        btnOpen.FlatAppearance.BorderSize = 0;
        btnOpen.MouseEnter += (_, _) => btnOpen.BackColor = ThemeConstants.PrimaryDark;
        btnOpen.MouseLeave += (_, _) => btnOpen.BackColor = ThemeConstants.Primary;
        btnOpen.Click += BtnOpen_Click;

        btnSave = new Button
        {
            Text = "保存更改",
            FlatStyle = FlatStyle.Flat,
            BackColor = ThemeConstants.ButtonSave,
            ForeColor = ThemeConstants.TextOnPrimary,
            Font = ThemeConstants.FontBody,
            Size = new Size(110, 34),
            Cursor = Cursors.Hand,
            TextAlign = ContentAlignment.MiddleCenter,
            Enabled = false
        };
        btnSave.FlatAppearance.BorderSize = 0;
        btnSave.Click += BtnSave_Click;

        lblSheetLabel = new Label
        {
            Text = "工作表",
            Font = ThemeConstants.FontBody,
            ForeColor = ThemeConstants.TextSecondary,
            Size = new Size(60, 24),
            TextAlign = ContentAlignment.MiddleLeft,
            AutoSize = true
        };

        cmbSheet = new ComboBox
        {
            DropDownStyle = ComboBoxStyle.DropDownList,
            FlatStyle = FlatStyle.Flat,
            Font = ThemeConstants.FontBody,
            Size = new Size(170, 30),
            Anchor = AnchorStyles.Left,
            Enabled = false
        };
        cmbSheet.SelectedIndexChanged += CmbSheet_SelectedIndexChanged;

        toolbarPanel.Controls.Add(btnOpen, 0, 0);
        toolbarPanel.Controls.Add(btnSave, 1, 0);
        toolbarPanel.Controls.Add(lblSheetLabel, 3, 0);
        toolbarPanel.Controls.Add(cmbSheet, 4, 0);
        var spacer = new Panel { Size = new Size(1, 1), BackColor = Color.Transparent };
        toolbarPanel.Controls.Add(spacer, 2, 0);

        // ── Excel DataGridView ─────────────────────────────────
        dataGrid = new DataGridView
        {
            Dock = DockStyle.Fill,
            BackgroundColor = ThemeConstants.BackgroundWhite,
            BorderStyle = BorderStyle.None,
            AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill,
            AllowUserToAddRows = false,
            AllowUserToDeleteRows = false,
            AllowUserToResizeRows = false,
            ReadOnly = false,
            RowHeadersVisible = true,
            RowHeadersWidth = ThemeConstants.RowHeaderWidth,
            SelectionMode = DataGridViewSelectionMode.FullRowSelect,
            Font = ThemeConstants.FontBody,
            GridColor = ThemeConstants.GridLine,
            Visible = false
        };

        // ── Word/PDF Container (Paper look) ───────────────────
        wordContainer = new Panel
        {
            Dock = DockStyle.Fill,
            BackColor = ThemeConstants.BackgroundPaper,
            Padding = new Padding(40, 15, 40, 15),
            Visible = false
        };

        wordView = new RichTextBox
        {
            Dock = DockStyle.Fill,
            BorderStyle = BorderStyle.None,
            BackColor = ThemeConstants.BackgroundWhite,
            Font = ThemeConstants.FontParagraph,
            ReadOnly = true,
            DetectUrls = true,
            SelectionIndent = 10,
            SelectionRightIndent = 10
        };
        wordContainer.Controls.Add(wordView);

        // ── Status bar ─────────────────────────────────────────
        statusBar = new StatusStrip
        {
            BackColor = ThemeConstants.BackgroundLight,
            SizingGrip = false,
            Padding = new Padding(0, 4, 12, 4),
            Renderer = new ModernRenderer()
        };
        statusLabel = new ToolStripStatusLabel
        {
            Text = "就绪 — 请打开文档文件",
            ForeColor = ThemeConstants.TextSecondary,
            Font = ThemeConstants.FontStatus
        };
        statusStats = new ToolStripStatusLabel
        {
            ForeColor = ThemeConstants.TextSecondary,
            Font = ThemeConstants.FontStatus,
            Alignment = ToolStripItemAlignment.Right
        };
        statusBar.Items.Add(statusLabel);
        statusBar.Items.Add(statusStats);

        // ── Form ───────────────────────────────────────────────
        AllowDrop = true;
        AutoScaleDimensions = new SizeF(96, 96);
        AutoScaleMode = AutoScaleMode.Dpi;
        BackColor = ThemeConstants.BackgroundWhite;
        ClientSize = new Size(1200, 720);
        Controls.Add(wordContainer);
        Controls.Add(dataGrid);
        Controls.Add(toolbarPanel);
        Controls.Add(titleBar);
        Controls.Add(statusBar);
        MinimumSize = new Size(640, 400);
        StartPosition = FormStartPosition.CenterScreen;
        Text = "Office 阅读器";
        DragEnter += Form_DragEnter;
        DragDrop += Form_DragDrop;

        ResumeLayout(false);
        PerformLayout();
    }
}