namespace SheetsAddIn {
    partial class SheetsPanel {
        /// <summary> 
        /// 必要なデザイナー変数です。
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        /// <summary> 
        /// 使用中のリソースをすべてクリーンアップします。
        /// </summary>
        /// <param name="disposing">マネージド リソースを破棄する場合は true を指定し、その他の場合は false を指定します。</param>
        protected override void Dispose( bool disposing ) {
            if( disposing && (components != null) ) {
                components.Dispose();
            }
            base.Dispose( disposing );
        }

        #region コンポーネント デザイナーで生成されたコード

        /// <summary> 
        /// デザイナー サポートに必要なメソッドです。このメソッドの内容を 
        /// コード エディターで変更しないでください。
        /// </summary>
        private void InitializeComponent() {
            this.components = new System.ComponentModel.Container();
            System.Windows.Forms.Panel panel1;
            this.TxtFilter = new System.Windows.Forms.TextBox();
            this.ChkShowHidden = new System.Windows.Forms.CheckBox();
            this.LstSheets = new System.Windows.Forms.ListBox();
            this.stbInfo = new System.Windows.Forms.StatusStrip();
            this.lblInfo = new System.Windows.Forms.ToolStripStatusLabel();
            this.cmPopup = new System.Windows.Forms.ContextMenuStrip(this.components);
            this.tmrFilter = new System.Windows.Forms.Timer(this.components);
            panel1 = new System.Windows.Forms.Panel();
            panel1.SuspendLayout();
            this.stbInfo.SuspendLayout();
            this.SuspendLayout();
            // 
            // TxtFilter
            // 
            this.TxtFilter.Dock = System.Windows.Forms.DockStyle.Top;
            this.TxtFilter.Location = new System.Drawing.Point(5, 5);
            this.TxtFilter.Margin = new System.Windows.Forms.Padding(4);
            this.TxtFilter.Name = "TxtFilter";
            this.TxtFilter.Size = new System.Drawing.Size(906, 31);
            this.TxtFilter.TabIndex = 0;
            this.TxtFilter.Text = "textBox1";
            this.TxtFilter.TextChanged += new System.EventHandler(this.TxtFilter_TextChanged);
            // 
            // panel1
            // 
            panel1.Controls.Add(this.ChkShowHidden);
            panel1.Dock = System.Windows.Forms.DockStyle.Top;
            panel1.Location = new System.Drawing.Point(5, 36);
            panel1.Margin = new System.Windows.Forms.Padding(4);
            panel1.Name = "panel1";
            panel1.Size = new System.Drawing.Size(906, 59);
            panel1.TabIndex = 1;
            // 
            // ChkShowHidden
            // 
            this.ChkShowHidden.AutoSize = true;
            this.ChkShowHidden.Dock = System.Windows.Forms.DockStyle.Right;
            this.ChkShowHidden.Location = new System.Drawing.Point(724, 0);
            this.ChkShowHidden.Margin = new System.Windows.Forms.Padding(4);
            this.ChkShowHidden.Name = "ChkShowHidden";
            this.ChkShowHidden.Size = new System.Drawing.Size(182, 59);
            this.ChkShowHidden.TabIndex = 0;
            this.ChkShowHidden.Text = "隠しシート表示";
            this.ChkShowHidden.UseVisualStyleBackColor = true;
            this.ChkShowHidden.CheckedChanged += new System.EventHandler(this.ThkShowHidden_CheckedChanged);
            // 
            // LstSheets
            // 
            this.LstSheets.Dock = System.Windows.Forms.DockStyle.Fill;
            this.LstSheets.DrawMode = System.Windows.Forms.DrawMode.OwnerDrawVariable;
            this.LstSheets.Font = new System.Drawing.Font("MS UI Gothic", 14F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(128)));
            this.LstSheets.FormattingEnabled = true;
            this.LstSheets.Items.AddRange(new object[] {
            "listBoxItem1",
            "listBoxItem2",
            "listBoxItem3"});
            this.LstSheets.Location = new System.Drawing.Point(5, 95);
            this.LstSheets.Margin = new System.Windows.Forms.Padding(4);
            this.LstSheets.Name = "LstSheets";
            this.LstSheets.Size = new System.Drawing.Size(906, 841);
            this.LstSheets.TabIndex = 2;
            this.LstSheets.DrawItem += new System.Windows.Forms.DrawItemEventHandler(this.LstSheets_DrawItem);
            this.LstSheets.MeasureItem += new System.Windows.Forms.MeasureItemEventHandler(this.LstSheets_MeasureItem);
            this.LstSheets.SelectedIndexChanged += new System.EventHandler(this.LstSheets_SelectedIndexChanged);
            // 
            // stbInfo
            // 
            this.stbInfo.Font = new System.Drawing.Font("Yu Gothic UI", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(128)));
            this.stbInfo.ImageScalingSize = new System.Drawing.Size(24, 24);
            this.stbInfo.Items.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.lblInfo});
            this.stbInfo.Location = new System.Drawing.Point(5, 897);
            this.stbInfo.Name = "stbInfo";
            this.stbInfo.Size = new System.Drawing.Size(906, 39);
            this.stbInfo.SizingGrip = false;
            this.stbInfo.TabIndex = 3;
            this.stbInfo.Text = "statusStrip1";
            // 
            // lblInfo
            // 
            this.lblInfo.Name = "lblInfo";
            this.lblInfo.Size = new System.Drawing.Size(845, 32);
            this.lblInfo.Spring = true;
            this.lblInfo.Text = "toolStripStatusLabel1";
            this.lblInfo.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
            // 
            // cmPopup
            // 
            this.cmPopup.ImageScalingSize = new System.Drawing.Size(24, 24);
            this.cmPopup.Name = "cmPopup";
            this.cmPopup.Size = new System.Drawing.Size(61, 4);
            // 
            // tmrFilter
            // 
            this.tmrFilter.Interval = 500;
            this.tmrFilter.Tick += new System.EventHandler(this.tmrFilter_Tick);
            // 
            // SheetsPanel
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(13F, 24F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.Controls.Add(this.stbInfo);
            this.Controls.Add(this.LstSheets);
            this.Controls.Add(panel1);
            this.Controls.Add(this.TxtFilter);
            this.Font = new System.Drawing.Font("MS UI Gothic", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(128)));
            this.Margin = new System.Windows.Forms.Padding(4);
            this.Name = "SheetsPanel";
            this.Padding = new System.Windows.Forms.Padding(5);
            this.Size = new System.Drawing.Size(916, 941);
            this.Load += new System.EventHandler(this.SheetsPanel_Load);
            this.VisibleChanged += new System.EventHandler(this.SheetsPanel_VisibleChanged);
            panel1.ResumeLayout(false);
            panel1.PerformLayout();
            this.stbInfo.ResumeLayout(false);
            this.stbInfo.PerformLayout();
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.TextBox TxtFilter;
        private System.Windows.Forms.CheckBox ChkShowHidden;
        private System.Windows.Forms.ListBox LstSheets;
        private System.Windows.Forms.StatusStrip stbInfo;
        private System.Windows.Forms.ContextMenuStrip cmPopup;
        private System.Windows.Forms.ToolStripStatusLabel lblInfo;
        private System.Windows.Forms.Timer tmrFilter;
    }
}
