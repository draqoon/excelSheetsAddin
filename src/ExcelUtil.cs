using Microsoft.Office.Interop.Excel;

using System;
using System.Collections.Generic;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace SheetsAddIn {
    static class ExcelUtil {

        /// <summary>
        /// Excel の色データから <see cref="Color" /> に変換する
        /// </summary>
        /// <param name="excelColor">Excel の色データ</param>
        /// <param name="defaultColor">Excel の色データが既定職の場合の色</param>
        /// <returns></returns>
        public static Color ConvertExcelColorToColor( dynamic excelColor, Color defaultColor ) {
            if( excelColor is bool b && !b ) {
                return defaultColor;
            }
            else if( excelColor is int bgr ) {
                var red = bgr & 0x0000FF;
                var green = (bgr & 0x00FF00) >> 8;
                var blue = (bgr & 0xFF0000) >> 16;
                var color = Color.FromArgb( red, green, blue );
                return color;
            }

            return defaultColor;
        }

        /// <summary>
        /// 背景色から白文字か黒文字を判定
        /// </summary>
        /// <param name="red">背景色（赤）</param>
        /// <param name="green">背景色（緑）</param>
        /// <param name="blue">背景色（青）</param>
        /// <returns></returns>
        public static Color ChooseTextColor( byte red, byte green, byte blue ) {

            #region 内部関数

            // RGB から相対輝度を算出（0.0 ～ 1.0）
            double RelativeLuminance( byte R, byte G, byte B ) {
                // RGB の各値を相対輝度算出用に変換
                double toRgb(byte rgb) {
                    double srgb = (double)rgb / 255;
                    return srgb <= 0.03928 ? srgb / 12.92 : Math.Pow( (srgb + 0.055) / 1.055, 2.4 );
                };

                return 0.2126 * toRgb( R ) + 0.7152 * toRgb( G ) + 0.0722 * toRgb( B );
            }

            // 2つの相対輝度値から、相対輝度比率を算出（0.0 ～ 21.0）
            // 相対輝度比率が 7.0 以上の値だと見やすい
            double RelativeLuminanceRatio( double relativeLuminance1, double relativeLuminance2 ) {
                // 相対輝度比率 = (大きい値 + 0.05) / (小さい値 + 0.05)
                return (Math.Max( relativeLuminance1, relativeLuminance2 ) + 0.05) / (Math.Min( relativeLuminance1, relativeLuminance2 ) + 0.05);
            }

            #endregion

            // 背景色の相対輝度
            double background = RelativeLuminance( red, green, blue );

            const double white = 1.0D;  // 白の相対輝度
            const double black = 0.0D;  // 黒の相対輝度

            // 文字色と背景色のコントラスト比を計算
            double whiteContrast = RelativeLuminanceRatio( white, background );   // 文字色：白との比
            double blackContrast = RelativeLuminanceRatio( black, background );   // 文字色：黒との比

            // コントラスト比が大きい文字色を採用
            return whiteContrast < blackContrast ? Color.Black : Color.White;
        }

        /// <summary>
        /// 背景色から白文字か黒文字を判定
        /// </summary>
        /// <param name="backColor">背景色</param>
        /// <returns></returns>
        public static Color ChooseTextColor( Color backColor ) {
            return ChooseTextColor( backColor.R, backColor.G, backColor.B );
        }
    }
}
