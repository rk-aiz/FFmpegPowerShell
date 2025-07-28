using System;
using System.Windows.Media;
using System.Collections.Concurrent;
using System.Collections.Generic;
using System.ComponentModel;
using System.Diagnostics;
using System.Globalization;
using System.Linq;
using System.Runtime.ConstrainedExecution;
using System.Runtime.CompilerServices;
using System.Runtime.InteropServices;
using System.Text.RegularExpressions;
using System.Threading;
using System.Threading.Tasks;
using System.Windows.Data;
using System.Windows.Input;
using Microsoft.Win32.SafeHandles;


namespace ThemeHelper
{
    public class ThemeBase : INotifyPropertyChanged
    {
        // dwmapi.dll から色情報を取得する
        [DllImport("dwmapi.dll", EntryPoint = "DwmGetColorizationColor",
            PreserveSig = true)]
        private static extern int DwmGetColorizationColor(
            out uint pcrColorization,
            out bool pfOpaqueBlend);


        // メインテキスト用フォントファミリー
        private FontFamily _mainFontFamily = new FontFamily("Segoe UI");
        public FontFamily MainFontFamily
        {
            get { return _mainFontFamily; }
            set
            {
                if (_mainFontFamily != value)
                {
                    _mainFontFamily = value;
                    OnPropertyChanged("MainFontFamily");
                }
            }
        }

        // メインテキスト用フォントサイズ
        private double _mainFontSize = 14;
        public double MainFontSize
        {
            get { return _mainFontSize; }
            set
            {
                if (_mainFontSize != value)
                {
                    _mainFontSize = value;
                    OnPropertyChanged("MainFontSize");
                }
            }
        }

        // 見出し用フォントファミリー
        private FontFamily _headingFontFamily = new FontFamily("Segoe UI Semibold");
        public FontFamily HeadingFontFamily
        {
            get { return _headingFontFamily; }
            set
            {
                if (_headingFontFamily != value)
                {
                    _headingFontFamily = value;
                    OnPropertyChanged("HeadingFontFamily");
                }
            }
        }

        // 見出し用フォントサイズ
        private double _headingFontSize = 18;
        public double HeadingFontSize
        {
            get { return _headingFontSize; }
            set
            {
                if (_headingFontSize != value)
                {
                    _headingFontSize = value;
                    OnPropertyChanged("HeadingFontSize");
                }
            }
        }

        // モノスペース用フォントファミリー
        private FontFamily _monospaceFontFamily = new FontFamily("Consolas");
        public FontFamily MonospaceFontFamily
        {
            get { return _monospaceFontFamily; }
            set
            {
                if (_monospaceFontFamily != value)
                {
                    _monospaceFontFamily = value;
                    OnPropertyChanged("MonospaceFontFamily");
                }
            }
        }

        // アイコン用フォントファミリー
        private FontFamily _iconFontFamily = new FontFamily("Segoe MDL2 Assets");
        public FontFamily IconFontFamily
        {
            get { return _iconFontFamily; }
            set
            {
                if (_iconFontFamily != value)
                {
                    _iconFontFamily = value;
                    OnPropertyChanged("IconFontFamily");
                }
            }
        }

        // ロゴ用フォントファミリー
        private FontFamily _logoFontFamily = new FontFamily("Calibri");
        public FontFamily LogoFontFamily
        {
            get { return _logoFontFamily; }
            set
            {
                if (_logoFontFamily != value)
                {
                    _logoFontFamily = value;
                    OnPropertyChanged("LogoFontFamily");
                }
            }
        }


        // INotifyPropertyChanged 実装部
        public event PropertyChangedEventHandler PropertyChanged;
        protected void OnPropertyChanged(string name)
        {
            PropertyChangedEventHandler handler = this.PropertyChanged;
            if (handler != null)
            {
                handler(this, new PropertyChangedEventArgs(name));
            }
        }

        /// <summary>
        /// Windowsのアクセントカラー（Colorization Color）を取得する
        /// </summary>
        public static Color GetAccentColor()
        {
            // ARGBを受け取る
            uint colorizationColor;
            bool opaque;
            try
            {
                // ここでDwmGetColorizationColorが失敗した場合、HRESULTを取得します
                int hr = DwmGetColorizationColor(out colorizationColor, out opaque);
                if (hr != 0) // HRESULTはエラーコードです。通常、0以外の値はエラーを示します
                {
                    throw new InvalidOperationException(
                        String.Format("DwmGetColorizationColor failed: 0x{0:X8}", hr));
                }
            }
            catch (Exception ex)
            {
                // 例外処理
                Console.WriteLine(ex.Message);
                colorizationColor = 0xFFFDF5D7;
            }

            // DWORD (AARRGGBB) → Color
            byte a = (byte)((colorizationColor >> 24) & 0xFF);
            byte r = (byte)((colorizationColor >> 16) & 0xFF);
            byte g = (byte)((colorizationColor >> 8) & 0xFF);
            byte b = (byte)(colorizationColor & 0xFF);
            return Color.FromArgb(a, r, g, b);
        }


        private Brush _accentColorBrush = null;

        public Brush AccentColorBrush
        {
            get
            {
                if (null == _accentColorBrush)
                {
                    try
                    {
                        _accentColorBrush = new SolidColorBrush(GetAccentColor());
                    }
                    catch (Exception e)
                    {
                        _accentColorBrush = new SolidColorBrush(
                            new Color { A = 255, R = 15, G = 15, B = 15 }
                        );
                        Console.WriteLine(e);
                    }
                }
                return _accentColorBrush;
            }
        }

        public static Geometry GetGeometry(string source)
        {
            var geometry = Geometry.Parse(source);
            return geometry;
        }

        public static SolidColorBrush GetFreezedBrush(byte R, byte G, byte B, byte A = 255)
        {
            var brush = new SolidColorBrush(
                new Color { A = A, R = R, G = G, B = B }
            );
            brush.Freeze();
            return brush;
        }

        public static SolidColorBrush GetFreezedBrush(Color color)
        {
            var brush = new SolidColorBrush(color);
            brush.Freeze();
            return brush;
        }
    }
}
