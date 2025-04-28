namespace WarrantAssistant
{
    partial class FrmRename
    {
        /// <summary>
        /// Required designer variable.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        /// <summary>
        /// Clean up any resources being used.
        /// </summary>
        /// <param name="disposing">true if managed resources should be disposed; otherwise, false.</param>
        protected override void Dispose(bool disposing) {
            if (disposing && (components != null)) {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        #region Windows Form Designer generated code

        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InitializeComponent() {
            this.components = new System.ComponentModel.Container();
            Infragistics.Win.Appearance appearance1 = new Infragistics.Win.Appearance();
            Infragistics.Win.Appearance appearance2 = new Infragistics.Win.Appearance();
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(FrmRename));
            Infragistics.Win.Appearance appearance3 = new Infragistics.Win.Appearance();
            Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand1 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
            Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn1 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("WName");
            Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn2 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("SerialNumber");
            Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn3 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Market");
            Infragistics.Win.Appearance appearance4 = new Infragistics.Win.Appearance();
            Infragistics.Win.Appearance appearance5 = new Infragistics.Win.Appearance();
            Infragistics.Win.Appearance appearance6 = new Infragistics.Win.Appearance();
            Infragistics.Win.Appearance appearance7 = new Infragistics.Win.Appearance();
            Infragistics.Win.Appearance appearance8 = new Infragistics.Win.Appearance();
            Infragistics.Win.Appearance appearance9 = new Infragistics.Win.Appearance();
            Infragistics.Win.Appearance appearance10 = new Infragistics.Win.Appearance();
            Infragistics.Win.Appearance appearance11 = new Infragistics.Win.Appearance();
            Infragistics.Win.Appearance appearance12 = new Infragistics.Win.Appearance();
            Infragistics.Win.Appearance appearance13 = new Infragistics.Win.Appearance();
            Infragistics.Win.Appearance appearance14 = new Infragistics.Win.Appearance();
            Infragistics.Win.Appearance appearance15 = new Infragistics.Win.Appearance();
            Infragistics.Win.Appearance appearance16 = new Infragistics.Win.Appearance();
            Infragistics.Win.Appearance appearance17 = new Infragistics.Win.Appearance();
            Infragistics.Win.Appearance appearance18 = new Infragistics.Win.Appearance();
            Infragistics.Win.Appearance appearance19 = new Infragistics.Win.Appearance();
            Infragistics.Win.Appearance appearance20 = new Infragistics.Win.Appearance();
            Infragistics.Win.Appearance appearance21 = new Infragistics.Win.Appearance();
            Infragistics.Win.Appearance appearance22 = new Infragistics.Win.Appearance();
            Infragistics.Win.UltraWinScrollBar.ScrollBarLook scrollBarLook1 = new Infragistics.Win.UltraWinScrollBar.ScrollBarLook();
            Infragistics.Win.Appearance appearance23 = new Infragistics.Win.Appearance();
            Infragistics.Win.Appearance appearance24 = new Infragistics.Win.Appearance();
            Infragistics.Win.Appearance appearance25 = new Infragistics.Win.Appearance();
            Infragistics.Win.Appearance appearance26 = new Infragistics.Win.Appearance();
            Infragistics.Win.Appearance appearance27 = new Infragistics.Win.Appearance();
            Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn4 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("WName");
            Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn5 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("SerialNumber");
            Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn6 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Market");
            this.ultraGrid1 = new Infragistics.Win.UltraWinGrid.UltraGrid();
            this.ultraDataSource1 = new Infragistics.Win.UltraWinDataSource.UltraDataSource(this.components);
            this.RenameFiles = new System.Windows.Forms.ToolStripButton();
            this.RenameXML = new System.Windows.Forms.ToolStripButton();
            this.toolStrip1 = new System.Windows.Forms.ToolStrip();
            this.RenameFilesB = new System.Windows.Forms.ToolStripButton();
            this.toolStripButton1 = new System.Windows.Forms.ToolStripButton();
            this.toolStripButton2 = new System.Windows.Forms.ToolStripButton();
            this.toolStripTextBox1 = new System.Windows.Forms.ToolStripTextBox();
            ((System.ComponentModel.ISupportInitialize)(this.ultraGrid1)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.ultraDataSource1)).BeginInit();
            this.toolStrip1.SuspendLayout();
            this.SuspendLayout();
            // 
            // ultraGrid1
            // 
            this.ultraGrid1.DataSource = this.ultraDataSource1;
            appearance1.BackColor = System.Drawing.Color.White;
            appearance1.BackColor2 = System.Drawing.Color.FromArgb(((int)(((byte)(221)))), ((int)(((byte)(221)))), ((int)(((byte)(221)))));
            appearance1.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
            appearance1.ForeColor = System.Drawing.Color.FromArgb(((int)(((byte)(21)))), ((int)(((byte)(27)))), ((int)(((byte)(85)))));
            this.ultraGrid1.DisplayLayout.AddNewBox.Appearance = appearance1;
            this.ultraGrid1.DisplayLayout.AddNewBox.BorderStyle = Infragistics.Win.UIElementBorderStyle.None;
            appearance2.ForeColor = System.Drawing.Color.FromArgb(((int)(((byte)(21)))), ((int)(((byte)(27)))), ((int)(((byte)(85)))));
            appearance2.ImageBackground = ((System.Drawing.Image)(resources.GetObject("appearance2.ImageBackground")));
            appearance2.ImageBackgroundAlpha = Infragistics.Win.Alpha.UseAlphaLevel;
            appearance2.ImageBackgroundStretchMargins = new Infragistics.Win.ImageBackgroundStretchMargins(6, 3, 6, 3);
            appearance2.ImageBackgroundStyle = Infragistics.Win.ImageBackgroundStyle.Stretched;
            this.ultraGrid1.DisplayLayout.AddNewBox.ButtonAppearance = appearance2;
            this.ultraGrid1.DisplayLayout.AddNewBox.ButtonConnectorColor = System.Drawing.Color.FromArgb(((int)(((byte)(21)))), ((int)(((byte)(27)))), ((int)(((byte)(85)))));
            this.ultraGrid1.DisplayLayout.AddNewBox.ButtonStyle = Infragistics.Win.UIElementButtonStyle.FlatBorderless;
            appearance3.BackColor = System.Drawing.Color.White;
            this.ultraGrid1.DisplayLayout.Appearance = appearance3;
            this.ultraGrid1.DisplayLayout.AutoFitStyle = Infragistics.Win.UltraWinGrid.AutoFitStyle.ResizeAllColumns;
            ultraGridColumn1.Header.VisiblePosition = 0;
            ultraGridColumn1.Width = 236;
            ultraGridColumn2.Header.VisiblePosition = 1;
            ultraGridColumn2.Width = 220;
            ultraGridColumn3.Header.VisiblePosition = 2;
            ultraGridColumn3.Width = 224;
            ultraGridBand1.Columns.AddRange(new object[] {
            ultraGridColumn1,
            ultraGridColumn2,
            ultraGridColumn3});
            this.ultraGrid1.DisplayLayout.BandsSerializer.Add(ultraGridBand1);
            this.ultraGrid1.DisplayLayout.BorderStyle = Infragistics.Win.UIElementBorderStyle.Solid;
            appearance4.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(21)))), ((int)(((byte)(27)))), ((int)(((byte)(85)))));
            appearance4.FontData.Name = "Trebuchet MS";
            appearance4.FontData.SizeInPoints = 9F;
            appearance4.ForeColor = System.Drawing.Color.White;
            appearance4.TextHAlignAsString = "Right";
            this.ultraGrid1.DisplayLayout.CaptionAppearance = appearance4;
            this.ultraGrid1.DisplayLayout.CaptionVisible = Infragistics.Win.DefaultableBoolean.False;
            this.ultraGrid1.DisplayLayout.FixedHeaderOffImage = ((System.Drawing.Image)(resources.GetObject("ultraGrid1.DisplayLayout.FixedHeaderOffImage")));
            this.ultraGrid1.DisplayLayout.FixedHeaderOnImage = ((System.Drawing.Image)(resources.GetObject("ultraGrid1.DisplayLayout.FixedHeaderOnImage")));
            this.ultraGrid1.DisplayLayout.FixedRowOffImage = ((System.Drawing.Image)(resources.GetObject("ultraGrid1.DisplayLayout.FixedRowOffImage")));
            this.ultraGrid1.DisplayLayout.FixedRowOnImage = ((System.Drawing.Image)(resources.GetObject("ultraGrid1.DisplayLayout.FixedRowOnImage")));
            appearance5.FontData.BoldAsString = "True";
            appearance5.FontData.Name = "Trebuchet MS";
            appearance5.FontData.SizeInPoints = 10F;
            appearance5.ForeColor = System.Drawing.Color.FromArgb(((int)(((byte)(60)))), ((int)(((byte)(127)))), ((int)(((byte)(177)))));
            appearance5.ImageBackground = ((System.Drawing.Image)(resources.GetObject("appearance5.ImageBackground")));
            appearance5.ImageBackgroundStretchMargins = new Infragistics.Win.ImageBackgroundStretchMargins(0, 2, 0, 3);
            appearance5.ImageBackgroundStyle = Infragistics.Win.ImageBackgroundStyle.Stretched;
            this.ultraGrid1.DisplayLayout.GroupByBox.Appearance = appearance5;
            appearance6.ForeColor = System.Drawing.SystemColors.GrayText;
            this.ultraGrid1.DisplayLayout.GroupByBox.BandLabelAppearance = appearance6;
            this.ultraGrid1.DisplayLayout.GroupByBox.BorderStyle = Infragistics.Win.UIElementBorderStyle.Solid;
            this.ultraGrid1.DisplayLayout.GroupByBox.ButtonBorderStyle = Infragistics.Win.UIElementBorderStyle.None;
            appearance7.BackColor = System.Drawing.SystemColors.ControlLightLight;
            appearance7.BackColor2 = System.Drawing.SystemColors.Control;
            appearance7.BackGradientStyle = Infragistics.Win.GradientStyle.Horizontal;
            appearance7.ForeColor = System.Drawing.SystemColors.GrayText;
            this.ultraGrid1.DisplayLayout.GroupByBox.PromptAppearance = appearance7;
            this.ultraGrid1.DisplayLayout.MaxColScrollRegions = 1;
            this.ultraGrid1.DisplayLayout.MaxRowScrollRegions = 1;
            appearance8.BackColor = System.Drawing.SystemColors.Window;
            appearance8.ForeColor = System.Drawing.SystemColors.ControlText;
            this.ultraGrid1.DisplayLayout.Override.ActiveCellAppearance = appearance8;
            appearance9.BackColor = System.Drawing.SystemColors.Highlight;
            appearance9.ForeColor = System.Drawing.SystemColors.HighlightText;
            this.ultraGrid1.DisplayLayout.Override.ActiveRowAppearance = appearance9;
            this.ultraGrid1.DisplayLayout.Override.AllowAddNew = Infragistics.Win.UltraWinGrid.AllowAddNew.TemplateOnBottom;
            this.ultraGrid1.DisplayLayout.Override.AllowDelete = Infragistics.Win.DefaultableBoolean.True;
            this.ultraGrid1.DisplayLayout.Override.AllowUpdate = Infragistics.Win.DefaultableBoolean.True;
            this.ultraGrid1.DisplayLayout.Override.BorderStyleCell = Infragistics.Win.UIElementBorderStyle.None;
            this.ultraGrid1.DisplayLayout.Override.BorderStyleHeader = Infragistics.Win.UIElementBorderStyle.None;
            this.ultraGrid1.DisplayLayout.Override.BorderStyleRow = Infragistics.Win.UIElementBorderStyle.None;
            this.ultraGrid1.DisplayLayout.Override.ButtonStyle = Infragistics.Win.UIElementButtonStyle.FlatBorderless;
            appearance10.BackColor = System.Drawing.Color.Transparent;
            this.ultraGrid1.DisplayLayout.Override.CardAreaAppearance = appearance10;
            appearance11.BorderColor = System.Drawing.Color.Transparent;
            appearance11.FontData.Name = "Verdana";
            this.ultraGrid1.DisplayLayout.Override.CellAppearance = appearance11;
            appearance12.ForeColor = System.Drawing.Color.FromArgb(((int)(((byte)(21)))), ((int)(((byte)(27)))), ((int)(((byte)(85)))));
            appearance12.ImageBackground = ((System.Drawing.Image)(resources.GetObject("appearance12.ImageBackground")));
            appearance12.ImageBackgroundStretchMargins = new Infragistics.Win.ImageBackgroundStretchMargins(6, 3, 6, 3);
            appearance12.ImageBackgroundStyle = Infragistics.Win.ImageBackgroundStyle.Stretched;
            this.ultraGrid1.DisplayLayout.Override.CellButtonAppearance = appearance12;
            this.ultraGrid1.DisplayLayout.Override.CellClickAction = Infragistics.Win.UltraWinGrid.CellClickAction.EditAndSelectText;
            this.ultraGrid1.DisplayLayout.Override.CellPadding = 0;
            appearance13.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(240)))), ((int)(((byte)(248)))), ((int)(((byte)(251)))));
            this.ultraGrid1.DisplayLayout.Override.FilterCellAppearance = appearance13;
            appearance14.ForeColor = System.Drawing.Color.FromArgb(((int)(((byte)(21)))), ((int)(((byte)(27)))), ((int)(((byte)(85)))));
            appearance14.ImageBackground = ((System.Drawing.Image)(resources.GetObject("appearance14.ImageBackground")));
            appearance14.ImageBackgroundStretchMargins = new Infragistics.Win.ImageBackgroundStretchMargins(6, 3, 6, 3);
            this.ultraGrid1.DisplayLayout.Override.FilterClearButtonAppearance = appearance14;
            appearance15.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(251)))), ((int)(((byte)(251)))), ((int)(((byte)(251)))));
            appearance15.BackColorAlpha = Infragistics.Win.Alpha.Opaque;
            this.ultraGrid1.DisplayLayout.Override.FilterRowPromptAppearance = appearance15;
            appearance16.BackColor = System.Drawing.SystemColors.Control;
            appearance16.BackColor2 = System.Drawing.SystemColors.ControlDark;
            appearance16.BackGradientAlignment = Infragistics.Win.GradientAlignment.Element;
            appearance16.BackGradientStyle = Infragistics.Win.GradientStyle.Horizontal;
            appearance16.BorderColor = System.Drawing.SystemColors.Window;
            this.ultraGrid1.DisplayLayout.Override.GroupByRowAppearance = appearance16;
            appearance17.BackGradientStyle = Infragistics.Win.GradientStyle.None;
            appearance17.FontData.BoldAsString = "True";
            appearance17.FontData.Name = "Trebuchet MS";
            appearance17.FontData.SizeInPoints = 10F;
            appearance17.ForeColor = System.Drawing.Color.FromArgb(((int)(((byte)(85)))), ((int)(((byte)(85)))), ((int)(((byte)(85)))));
            appearance17.ImageBackground = ((System.Drawing.Image)(resources.GetObject("appearance17.ImageBackground")));
            appearance17.ImageBackgroundStyle = Infragistics.Win.ImageBackgroundStyle.Tiled;
            appearance17.TextHAlignAsString = "Left";
            appearance17.ThemedElementAlpha = Infragistics.Win.Alpha.Transparent;
            this.ultraGrid1.DisplayLayout.Override.HeaderAppearance = appearance17;
            this.ultraGrid1.DisplayLayout.Override.HeaderClickAction = Infragistics.Win.UltraWinGrid.HeaderClickAction.SortMulti;
            this.ultraGrid1.DisplayLayout.Override.HeaderStyle = Infragistics.Win.HeaderStyle.XPThemed;
            appearance18.BorderColor = System.Drawing.Color.Transparent;
            this.ultraGrid1.DisplayLayout.Override.RowAppearance = appearance18;
            appearance19.BackColor = System.Drawing.Color.White;
            this.ultraGrid1.DisplayLayout.Override.RowSelectorAppearance = appearance19;
            this.ultraGrid1.DisplayLayout.Override.RowSelectors = Infragistics.Win.DefaultableBoolean.False;
            appearance20.BorderColor = System.Drawing.Color.Transparent;
            appearance20.ForeColor = System.Drawing.Color.Black;
            this.ultraGrid1.DisplayLayout.Override.SelectedCellAppearance = appearance20;
            appearance21.BorderColor = System.Drawing.Color.Transparent;
            appearance21.ForeColor = System.Drawing.Color.FromArgb(((int)(((byte)(21)))), ((int)(((byte)(27)))), ((int)(((byte)(85)))));
            appearance21.ImageBackground = ((System.Drawing.Image)(resources.GetObject("appearance21.ImageBackground")));
            appearance21.ImageBackgroundStretchMargins = new Infragistics.Win.ImageBackgroundStretchMargins(1, 1, 1, 4);
            appearance21.ImageBackgroundStyle = Infragistics.Win.ImageBackgroundStyle.Stretched;
            this.ultraGrid1.DisplayLayout.Override.SelectedRowAppearance = appearance21;
            appearance22.BackColor = System.Drawing.SystemColors.ControlLight;
            this.ultraGrid1.DisplayLayout.Override.TemplateAddRowAppearance = appearance22;
            appearance23.ImageBackgroundStretchMargins = new Infragistics.Win.ImageBackgroundStretchMargins(2, 4, 2, 4);
            appearance23.ImageBackgroundStyle = Infragistics.Win.ImageBackgroundStyle.Stretched;
            scrollBarLook1.Appearance = appearance23;
            appearance24.ImageBackground = ((System.Drawing.Image)(resources.GetObject("appearance24.ImageBackground")));
            appearance24.ImageBackgroundStretchMargins = new Infragistics.Win.ImageBackgroundStretchMargins(3, 2, 3, 2);
            scrollBarLook1.AppearanceHorizontal = appearance24;
            appearance25.ImageBackground = ((System.Drawing.Image)(resources.GetObject("appearance25.ImageBackground")));
            appearance25.ImageBackgroundStretchMargins = new Infragistics.Win.ImageBackgroundStretchMargins(2, 3, 2, 3);
            appearance25.ImageBackgroundStyle = Infragistics.Win.ImageBackgroundStyle.Stretched;
            scrollBarLook1.AppearanceVertical = appearance25;
            appearance26.ImageBackground = ((System.Drawing.Image)(resources.GetObject("appearance26.ImageBackground")));
            appearance26.ImageBackgroundStretchMargins = new Infragistics.Win.ImageBackgroundStretchMargins(0, 2, 0, 1);
            scrollBarLook1.TrackAppearanceHorizontal = appearance26;
            appearance27.ImageBackground = ((System.Drawing.Image)(resources.GetObject("appearance27.ImageBackground")));
            appearance27.ImageBackgroundStretchMargins = new Infragistics.Win.ImageBackgroundStretchMargins(2, 0, 1, 0);
            appearance27.ImageBackgroundStyle = Infragistics.Win.ImageBackgroundStyle.Stretched;
            scrollBarLook1.TrackAppearanceVertical = appearance27;
            this.ultraGrid1.DisplayLayout.ScrollBarLook = scrollBarLook1;
            this.ultraGrid1.DisplayLayout.ScrollBounds = Infragistics.Win.UltraWinGrid.ScrollBounds.ScrollToFill;
            this.ultraGrid1.DisplayLayout.ScrollStyle = Infragistics.Win.UltraWinGrid.ScrollStyle.Immediate;
            this.ultraGrid1.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F);
            this.ultraGrid1.Location = new System.Drawing.Point(0, 28);
            this.ultraGrid1.Name = "ultraGrid1";
            this.ultraGrid1.Size = new System.Drawing.Size(682, 242);
            this.ultraGrid1.TabIndex = 1;
            this.ultraGrid1.Text = "Grid Caption Area";
            this.ultraGrid1.UseOsThemes = Infragistics.Win.DefaultableBoolean.False;
            this.ultraGrid1.InitializeLayout += new Infragistics.Win.UltraWinGrid.InitializeLayoutEventHandler(this.ultraGrid1_InitializeLayout);
            // 
            // ultraDataSource1
            // 
            this.ultraDataSource1.Band.Columns.AddRange(new object[] {
            ultraDataColumn4,
            ultraDataColumn5,
            ultraDataColumn6});
            // 
            // RenameFiles
            // 
            this.RenameFiles.DisplayStyle = System.Windows.Forms.ToolStripItemDisplayStyle.Text;
            this.RenameFiles.Image = ((System.Drawing.Image)(resources.GetObject("RenameFiles.Image")));
            this.RenameFiles.ImageTransparentColor = System.Drawing.Color.Magenta;
            this.RenameFiles.Name = "RenameFiles";
            this.RenameFiles.Size = new System.Drawing.Size(92, 22);
            this.RenameFiles.Text = "修改發行檔名A";
            this.RenameFiles.ToolTipText = "修改全部發行檔名";
            this.RenameFiles.Click += new System.EventHandler(this.RenameFiles_Click);
            // 
            // RenameXML
            // 
            this.RenameXML.DisplayStyle = System.Windows.Forms.ToolStripItemDisplayStyle.Text;
            this.RenameXML.Image = ((System.Drawing.Image)(resources.GetObject("RenameXML.Image")));
            this.RenameXML.ImageTransparentColor = System.Drawing.Color.Magenta;
            this.RenameXML.Name = "RenameXML";
            this.RenameXML.Size = new System.Drawing.Size(86, 22);
            this.RenameXML.Text = "修改XML檔名";
            this.RenameXML.Click += new System.EventHandler(this.RenameXML_Click);
            // 
            // toolStrip1
            // 
            this.toolStrip1.Items.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.RenameFiles,
            this.RenameFilesB,
            this.RenameXML,
            this.toolStripButton1,
            this.toolStripButton2,
            this.toolStripTextBox1});
            this.toolStrip1.Location = new System.Drawing.Point(0, 0);
            this.toolStrip1.Name = "toolStrip1";
            this.toolStrip1.Size = new System.Drawing.Size(682, 25);
            this.toolStrip1.TabIndex = 4;
            this.toolStrip1.Text = "toolStrip1";
            // 
            // RenameFilesB
            // 
            this.RenameFilesB.DisplayStyle = System.Windows.Forms.ToolStripItemDisplayStyle.Text;
            this.RenameFilesB.Image = ((System.Drawing.Image)(resources.GetObject("RenameFilesB.Image")));
            this.RenameFilesB.ImageTransparentColor = System.Drawing.Color.Magenta;
            this.RenameFilesB.Name = "RenameFilesB";
            this.RenameFilesB.Size = new System.Drawing.Size(91, 22);
            this.RenameFilesB.Text = "修改發行檔名B";
            this.RenameFilesB.ToolTipText = "修改下午產製發行檔名，即除了附表6~10與附件3, 5, 8 以外之所有檔案";
            this.RenameFilesB.Click += new System.EventHandler(this.RenameFilesB_Click);
            // 
            // toolStripButton1
            // 
            this.toolStripButton1.DisplayStyle = System.Windows.Forms.ToolStripItemDisplayStyle.Text;
            this.toolStripButton1.Image = ((System.Drawing.Image)(resources.GetObject("toolStripButton1.Image")));
            this.toolStripButton1.ImageTransparentColor = System.Drawing.Color.Magenta;
            this.toolStripButton1.Name = "toolStripButton1";
            this.toolStripButton1.Size = new System.Drawing.Size(84, 22);
            this.toolStripButton1.Text = "修改銷售公告";
            this.toolStripButton1.Click += new System.EventHandler(this.SalesAnnounce_Click);
            // 
            // toolStripButton2
            // 
            this.toolStripButton2.DisplayStyle = System.Windows.Forms.ToolStripItemDisplayStyle.Text;
            this.toolStripButton2.Image = ((System.Drawing.Image)(resources.GetObject("toolStripButton2.Image")));
            this.toolStripButton2.ImageTransparentColor = System.Drawing.Color.Magenta;
            this.toolStripButton2.Name = "toolStripButton2";
            this.toolStripButton2.Size = new System.Drawing.Size(84, 22);
            this.toolStripButton2.Text = "修改上市公告";
            this.toolStripButton2.Click += new System.EventHandler(this.ListedAnnounce_Click);
            // 
            // toolStripTextBox1
            // 
            this.toolStripTextBox1.Name = "toolStripTextBox1";
            this.toolStripTextBox1.Size = new System.Drawing.Size(150, 25);
            // 
            // FrmRename
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 12F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(682, 270);
            this.Controls.Add(this.toolStrip1);
            this.Controls.Add(this.ultraGrid1);
            this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedSingle;
            this.Icon = ((System.Drawing.Icon)(resources.GetObject("$this.Icon")));
            this.Name = "FrmRename";
            this.Text = "修改檔案名稱";
            this.Load += new System.EventHandler(this.FrmRename_Load);
            ((System.ComponentModel.ISupportInitialize)(this.ultraGrid1)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.ultraDataSource1)).EndInit();
            this.toolStrip1.ResumeLayout(false);
            this.toolStrip1.PerformLayout();
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion
        private Infragistics.Win.UltraWinGrid.UltraGrid ultraGrid1;
        private System.Windows.Forms.ToolStripButton RenameFiles;
        private System.Windows.Forms.ToolStripButton RenameXML;
        private System.Windows.Forms.ToolStrip toolStrip1;
        private System.Windows.Forms.ToolStripButton RenameFilesB;
        private Infragistics.Win.UltraWinDataSource.UltraDataSource ultraDataSource1;
        private System.Windows.Forms.ToolStripButton toolStripButton1;
        private System.Windows.Forms.ToolStripTextBox toolStripTextBox1;
        private System.Windows.Forms.ToolStripButton toolStripButton2;
    }
}