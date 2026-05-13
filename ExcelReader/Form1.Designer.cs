namespace ExcelReader;

partial class Form1
{
    private System.ComponentModel.IContainer components = null;

    private Panel titleBar;
    private Label lblTitle;
    private Label lblSubtitle;
    private TableLayoutPanel toolbarPanel;
    private Button btnOpen;
    private Label lblSheetLabel;
    private ComboBox cmbSheet;
    private DataGridView dataGrid;
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
        var primary = Color.FromArgb(43, 87, 154);
        var primaryDark = Color.FromArgb(30, 65, 120);
        var textSecondary = Color.FromArgb(100, 100, 100);

        SuspendLayout();

        // ── Title bar ──────────────────────────────────────────
        titleBar = new Panel
        {
            BackColor = primary,
            Dock = DockStyle.Top,
            Height = 60,
            Padding = new Padding(24, 0, 24, 0)
        };

        lblTitle = new Label
        {
            Text = "Office 阅读器",
            ForeColor = Color.White,
            Font = new Font("Microsoft YaHei UI", 16f, FontStyle.Bold),
            Location = new Point(24, 8),
            Size = new Size(400, 30),
            BackColor = Color.Transparent
        };

        lblSubtitle = new Label
        {
            Text = "拖拽文件到窗口，或点击按钮打开  ·  支持 Excel 和 Word",
            ForeColor = Color.FromArgb(180, 200, 230),
            Font = new Font("Microsoft YaHei UI", 9f),
            Location = new Point(24, 36),
            Size = new Size(500, 18),
            BackColor = Color.Transparent
        };

        titleBar.Controls.Add(lblTitle);
        titleBar.Controls.Add(lblSubtitle);

        // ── Toolbar ────────────────────────────────────────────
        toolbarPanel = new TableLayoutPanel
        {
            Dock = DockStyle.Top,
            Height = 52,
            Padding = new Padding(20, 10, 20, 10),
            BackColor = Color.White,
            ColumnCount = 4,
            RowCount = 1
        };
        toolbarPanel.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, 100));
        toolbarPanel.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, 20));
        toolbarPanel.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, 65));
        toolbarPanel.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100));

        btnOpen = new Button
        {
            Text = "  📂  打开文件",
            FlatStyle = FlatStyle.Flat,
            BackColor = primary,
            ForeColor = Color.White,
            Font = new Font("Microsoft YaHei UI", 9f),
            Size = new Size(100, 32),
            Cursor = Cursors.Hand,
            TextAlign = ContentAlignment.MiddleCenter
        };
        btnOpen.FlatAppearance.BorderSize = 0;
        btnOpen.MouseEnter += (_, _) => btnOpen.BackColor = primaryDark;
        btnOpen.MouseLeave += (_, _) => btnOpen.BackColor = primary;
        btnOpen.Click += BtnOpen_Click;

        lblSheetLabel = new Label
        {
            Text = "工作表",
            Font = new Font("Microsoft YaHei UI", 9f),
            ForeColor = textSecondary,
            Size = new Size(60, 24),
            TextAlign = ContentAlignment.MiddleLeft,
            AutoSize = true
        };

        cmbSheet = new ComboBox
        {
            DropDownStyle = ComboBoxStyle.DropDownList,
            FlatStyle = FlatStyle.Flat,
            Font = new Font("Microsoft YaHei UI", 9f),
            Size = new Size(160, 28),
            Anchor = AnchorStyles.Left,
            Enabled = false
        };
        cmbSheet.SelectedIndexChanged += CmbSheet_SelectedIndexChanged;

        toolbarPanel.Controls.Add(btnOpen, 0, 0);
        toolbarPanel.Controls.Add(lblSheetLabel, 2, 0);
        toolbarPanel.Controls.Add(cmbSheet, 3, 0);
        var spacer = new Panel { Size = new Size(1, 1), BackColor = Color.Transparent };
        toolbarPanel.Controls.Add(spacer, 1, 0);

        // ── Excel DataGridView ─────────────────────────────────
        dataGrid = new DataGridView
        {
            Dock = DockStyle.Fill,
            BackgroundColor = Color.FromArgb(245, 245, 245),
            BorderStyle = BorderStyle.None,
            AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.DisplayedCells,
            AllowUserToAddRows = false,
            AllowUserToDeleteRows = false,
            AllowUserToResizeRows = false,
            ReadOnly = true,
            RowHeadersVisible = false,
            SelectionMode = DataGridViewSelectionMode.FullRowSelect,
            Font = new Font("Microsoft YaHei UI", 9f),
            GridColor = Color.FromArgb(230, 230, 230),
            Visible = false
        };

        // ── Word RichTextBox ───────────────────────────────────
        wordView = new RichTextBox
        {
            Dock = DockStyle.Fill,
            BorderStyle = BorderStyle.None,
            BackColor = Color.White,
            Font = new Font("Microsoft YaHei UI", 10.5f),
            ReadOnly = true,
            Visible = false,
            DetectUrls = true
        };

        // ── Status bar ─────────────────────────────────────────
        statusBar = new StatusStrip
        {
            BackColor = Color.FromArgb(248, 248, 248),
            SizingGrip = false,
            Padding = new Padding(0, 3, 10, 3),
            Renderer = new ModernRenderer()
        };
        statusLabel = new ToolStripStatusLabel
        {
            Text = "就绪 — 请打开 Excel 或 Word 文件",
            ForeColor = textSecondary,
            Font = new Font("Microsoft YaHei UI", 8.5f)
        };
        statusStats = new ToolStripStatusLabel
        {
            ForeColor = textSecondary,
            Font = new Font("Microsoft YaHei UI", 8.5f),
            Alignment = ToolStripItemAlignment.Right
        };
        statusBar.Items.Add(statusLabel);
        statusBar.Items.Add(statusStats);

        // ── Form ───────────────────────────────────────────────
        AllowDrop = true;
        AutoScaleDimensions = new SizeF(96, 96);
        AutoScaleMode = AutoScaleMode.Dpi;
        BackColor = Color.White;
        ClientSize = new Size(1200, 720);
        Controls.Add(wordView);
        Controls.Add(dataGrid);
        Controls.Add(toolbarPanel);
        Controls.Add(titleBar);
        Controls.Add(statusBar);
        MinimumSize = new Size(640, 400);
        StartPosition = FormStartPosition.CenterScreen;
        Text = "Office 阅读器";
        DragEnter += Form_DragEnter;
        DragDrop += Form_DragDrop;
        Resize += (_, _) => Invalidate();

        ResumeLayout(false);
        PerformLayout();
    }
}
