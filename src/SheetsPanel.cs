using Microsoft.Office.Interop.Excel;

using stdole;

using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace SheetsAddIn {
    public partial class SheetsPanel : UserControl {

        private const int ListBoxItemMargin = 5;

        public SheetsPanel() {
            InitializeComponent();
        }

        #region events handlers

        /// <summary>
        ///
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void SheetsPanel_Load( object sender, EventArgs e ) {
            this.TxtFilter.Text = string.Empty;
            this.LstSheets.Items.Clear();
        }

        /// <summary>
        ///
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void SheetsPanel_VisibleChanged( object sender, EventArgs e ) {
            var ThisAddIn = Globals.ThisAddIn;
            var app = ThisAddIn.Application;
            var window = app.ActiveWindow;
            if( this.Visible ) {
                this.RefreshList();
                app.SheetActivate += this.Application_SheetActivate;
                app.SheetBeforeDelete += this.Application_SheetBeforeDelete;
                app.WorkbookNewSheet += this.Application_WorkbookNewSheet;
            }
            else {
                app.SheetActivate -= this.Application_SheetActivate;
                app.SheetBeforeDelete -= this.Application_SheetBeforeDelete;
                app.WorkbookNewSheet -= this.Application_WorkbookNewSheet;
            }
        }

        /// <summary>
        ///
        /// </summary>
        /// <param name="wb"></param>
        /// <param name="s"></param>
        private void Application_WorkbookNewSheet( Workbook wb, object s ) {
            this.RefreshList();
        }

        /// <summary>
        ///
        /// </summary>
        /// <param name="s"></param>
        private void Application_SheetBeforeDelete( object s ) {
            this.RefreshList();
        }

        /// <summary>
        ///
        /// </summary>
        /// <param name="sh"></param>
        private void Application_SheetActivate( object sh ) {
            var ThisAddIn = Globals.ThisAddIn;
            var app = ThisAddIn.Application;
            var window = app.ActiveWindow;
            var activeSheet = window.ActiveSheet as Worksheet;

            var pane = ThisAddIn.CustomTaskPanes.FirstOrDefault( p => (p.Window as Window).Hwnd == window.Hwnd && p.Control is SheetsPanel );
            if(pane?.Control == this ) {
                this.LstSheets.SelectedItem = activeSheet;
            }
        }

        /// <summary>
        ///
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void TxtFilter_TextChanged( object sender, EventArgs e ) {
            this.tmrFilter.Stop();
            this.tmrFilter.Start();
        }

        /// <summary>
        ///
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void ThkShowHidden_CheckedChanged( object sender, EventArgs e ) {
            this.RefreshList();
        }

        /// <summary>
        ///
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void LstSheets_SelectedIndexChanged( object sender, EventArgs e ) {
            if( this.LstSheets.SelectedItem is Worksheet sheet ) {
                if( sheet.Visible != XlSheetVisibility.xlSheetVisible ) {
                    sheet.Visible = XlSheetVisibility.xlSheetVisible;
                }
                sheet.Activate();
            }
        }

        /// <summary>
        ///
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void LstSheets_MeasureItem( object sender, MeasureItemEventArgs e ) {
            if( e.Index == -1 || !(this.LstSheets.Items[e.Index] is Worksheet sheet) )
                return;

            var text = this.GetDisplayText( sheet );

            var fontFamilly = this.LstSheets.Font.FontFamily;
            var fontSize = this.LstSheets.Font.Size;
            var fontStyle = this.LstSheets.Font.Style | FontStyle.Bold;
            using( var font = new System.Drawing.Font(fontFamilly, fontSize, fontStyle ) ) {
                var size = this.MeasureListBoxItemString( e.Graphics, text, font ).ToSize();
                e.ItemWidth = size.Width + ListBoxItemMargin * 4;
                e.ItemHeight = size.Height + ListBoxItemMargin * 4;
            }
        }

        /// <summary>
        ///
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void LstSheets_DrawItem( object sender, DrawItemEventArgs e ) {
            e.DrawBackground();

            var sheet = this.LstSheets.Items[e.Index] as Worksheet;
            var text = this.GetDisplayText( sheet );

            var fontFamilly = this.LstSheets.Font.FontFamily;
            var fontSize = this.LstSheets.Font.Size;
            var fontStyle = this.LstSheets.Font.Style;

            if( (e.State & DrawItemState.Selected) != DrawItemState.None ) {
                // 選択中
                fontFamilly = this.LstSheets.Font.FontFamily;
                //fontSize += 1;
                fontStyle |= FontStyle.Bold;
            }


            Color backColor = ExcelUtil.ConvertExcelColorToColor( sheet.Tab.Color, SystemColors.Window );
            Color foreColor = ExcelUtil.ChooseTextColor( backColor );

            // 背景描画
            if( (e.State & DrawItemState.Selected) != DrawItemState.None ) {
                // 選択中
                // 外枠(2px)背景色を設定
                using( var brush = new SolidBrush( SystemColors.Highlight ) ) {
                    e.Graphics.FillRectangle( brush, e.Bounds );
                }
            }

            // 内部にタブ色を設定
            var rect = e.Bounds;
            rect.X += 5;
            rect.Y += 5;
            rect.Width -= ListBoxItemMargin * 2;
            rect.Height -= ListBoxItemMargin * 2;
            using( var brush = new SolidBrush( backColor ) ) {
                e.Graphics.FillRectangle( brush, rect );
            }

            // 文字列描画
            using( var font = new System.Drawing.Font( fontFamilly, fontSize, fontStyle ) ) {
                var size = this.MeasureListBoxItemString( e.Graphics, text, font ).ToSize();

                rect = e.Bounds;
                rect.X += ListBoxItemMargin;
                rect.Y += (rect.Height - size.Height) / 2;
                rect.Width -= ListBoxItemMargin * 2;
                rect.Height -= ListBoxItemMargin * 2;

                using( var brush = new SolidBrush( foreColor ) ) {
                    e.Graphics.DrawString( text, font, brush, rect );
                }
            }

            //e.DrawFocusRectangle();
        }

        /// <summary>
        ///
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void tmrFilter_Tick( object sender, EventArgs e ) {
            this.tmrFilter.Stop();
            this.RefreshList();
        }

        #endregion

        #region private methods

        /// <summary>
        ///
        /// </summary>
        private void RefreshList() {
            var ThisAddIn = Globals.ThisAddIn;
            var app = ThisAddIn.Application;
            var wb = app.ActiveWorkbook;
            var window = app.ActiveWindow;
            var activeSheet = window?.ActiveSheet;

            if( activeSheet == null )
                return;

            this.LstSheets.Items.Clear();

            var visibleCount = 0;
            var hiddenCount = 0;

            foreach(var sheet in wb.Sheets.Cast<Worksheet>()) {
                switch( sheet.Visible ) {
                    case XlSheetVisibility.xlSheetVisible:
                        visibleCount += 1;
                        break;
                    case XlSheetVisibility.xlSheetVeryHidden:
                        continue;
                    case XlSheetVisibility.xlSheetHidden:
                        hiddenCount += 1;
                        if( !this.ChkShowHidden.Checked ) {
                            continue;
                        }
                        break;
                }

                var filter = this.TxtFilter.Text.Trim();
                if( sheet.Name.IndexOf( filter ) == -1 )
                    continue;

                var index = this.LstSheets.Items.Add( sheet );
                if(sheet == activeSheet ) {
                    this.LstSheets.SelectedIndex = index;
                }
            }

            this.lblInfo.Text = $"{visibleCount}シート(非表示{hiddenCount})";
        }

        /// <summary>
        ///
        /// </summary>
        /// <param name="sheet"></param>
        /// <returns></returns>
        private string GetDisplayText(Worksheet sheet) {
            var text = sheet.Name;
            if(sheet.Visible != XlSheetVisibility.xlSheetVisible ) {
                text = "(*)" + sheet.Name;
            }
            return text;
        }

        /// <summary>
        ///
        /// </summary>
        /// <param name="graphics"></param>
        /// <param name="text"></param>
        /// <param name="font"></param>
        /// <returns></returns>
        private SizeF MeasureListBoxItemString(Graphics graphics, string text, System.Drawing.Font font ) {
            var maxWidth = this.LstSheets.ClientRectangle.Width - ListBoxItemMargin * 2;
            var size = graphics.MeasureString( text, font, maxWidth );
            return size;
        }

        #endregion

    }
}
