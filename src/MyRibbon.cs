using Microsoft.Office.Tools.Ribbon;

using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace SheetsAddIn {
    public partial class MyRibbon {
        private void Ribbon1_Load( object sender, RibbonUIEventArgs e ) {

        }

        private void btnSheets_Click( object sender, RibbonControlEventArgs e ) {
            var ThisAddIn = Globals.ThisAddIn;
            var window = ThisAddIn.Application.ActiveWindow;

            var pane = ThisAddIn.CustomTaskPanes.FirstOrDefault( p => (p.Window as Microsoft.Office.Interop.Excel.Window).Hwnd == window.Hwnd && p.Control is SheetsPanel );
            if( pane != null ) {
                pane.Visible = !pane.Visible;
            }
            else {
                var uc = new SheetsPanel();
                pane = ThisAddIn.CustomTaskPanes.Add( uc, "シート一覧", window );
                pane.Visible = true;
            }
        }
    }
}
