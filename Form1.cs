using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Drawing;
using System.Drawing.Drawing2D;
using System.Globalization;
using System.Linq;
using System.Windows.Forms;

namespace CSharpColorPicker
{
    /// <summary>
    /// The main form for the application.
    /// </summary>
    public partial class Form1 : Form
    {
        #region Constructor

        /// <summary>
        /// Initializes a new instance of the <see cref="Form1"/> class.
        /// </summary>
        public Form1()
        {
            InitializeComponent();
        }

        #endregion

        #region Event Handlers

        private void Form1_Load(object sender, EventArgs e)
        {
        }

        #endregion
    }

    /// <summary>
    /// A custom button control that displays a selected color and opens a color picker drop-down when clicked.
    /// </summary>
    public class ExcelColorPopupButton : Button
    {
        #region Fields

        private ExcelColorDropDown _dropDown;

        #endregion

        #region Properties and Events

        /// <summary>
        /// Gets or sets the currently selected color.
        /// </summary>
        public Color SelectedColor { get; set; } = Color.White;

        /// <summary>
        /// Occurs when the selected color is changed.
        /// </summary>
        public event EventHandler ColorChanged;

        #endregion

        #region Constructor

        /// <summary>
        /// Initializes a new instance of the <see cref="ExcelColorPopupButton"/> class.
        /// </summary>
        public ExcelColorPopupButton()
        {
            Text = "";
            Width = 32;
            Height = 25;
            FlatStyle = FlatStyle.Standard;
        }

        #endregion

        #region Protected Overrides

        /// <summary>
        /// Paints the control.
        /// </summary>
        /// <param name="pevent">A <see cref="PaintEventArgs"/> that contains the event data.</param>
        protected override void OnPaint(PaintEventArgs pevent)
        {
            base.OnPaint(pevent);
            const int ColorRectWidth = 14;
            const int ColorRectHeight = 10;
            Rectangle colorRect = new Rectangle(10, (Height - ColorRectHeight) / 2, ColorRectWidth, ColorRectHeight);
            using (SolidBrush brush = new SolidBrush(SelectedColor))
            {
                pevent.Graphics.FillRectangle(brush, colorRect);
            }
            pevent.Graphics.DrawRectangle(SystemPens.ControlDark, colorRect);
            int arrowX = Width - 15;
            int arrowY = Height / 2;
            pevent.Graphics.FillPolygon(SystemBrushes.ControlText,
                new[] { new Point(arrowX, arrowY), new Point(arrowX + 6, arrowY), new Point(arrowX + 3, arrowY + 3) });
        }

        /// <summary>
        /// Raises the <see cref="Control.Click"/> event.
        /// </summary>
        /// <param name="e">An <see cref="EventArgs"/> that contains the event data.</param>
        protected override void OnClick(EventArgs e)
        {
            base.OnClick(e);
            if (_dropDown != null && !_dropDown.IsDisposed)
            {
                _dropDown.Close();
                return;
            }
            _dropDown = new ExcelColorDropDown(SelectedColor);
            _dropDown.ColorSelected += (s, color) =>
            {
                SelectedColor = color;
                Invalidate();
                ColorChanged?.Invoke(this, EventArgs.Empty);
            };
            _dropDown.Closed += (s, ev) =>
            {
                _dropDown = null;
                Focus();
            };
            Point screenPoint = PointToScreen(new Point(0, Height));
            Rectangle workingArea = Screen.FromControl(this).WorkingArea;
            if (screenPoint.X + _dropDown.Width > workingArea.Right)
            {
                screenPoint.X = workingArea.Right - _dropDown.Width;
            }
            if (screenPoint.Y + _dropDown.Height > workingArea.Bottom)
            {
                int topY = PointToScreen(Point.Empty).Y - _dropDown.Height;
                screenPoint.Y = topY < workingArea.Top ? workingArea.Top : topY;
            }
            _dropDown.Show(screenPoint);
        }

        /// <summary>
        /// Releases the unmanaged resources used by the <see cref="Button"/> and optionally releases the managed resources.
        /// </summary>
        /// <param name="disposing">true to release both managed and unmanaged resources; false to release only unmanaged resources.</param>
        protected override void Dispose(bool disposing)
        {
            if (disposing)
            {
                _dropDown?.Dispose();
            }
            base.Dispose(disposing);
        }

        #endregion
    }

    /// <summary>
    /// A custom <see cref="ToolStripDropDown"/> that hosts the <see cref="ExcelColorDropDownControl"/>.
    /// </summary>
    public class ExcelColorDropDown : ToolStripDropDown
    {
        #region Fields

        private readonly ExcelColorDropDownControl _hostedControl;
        private readonly ToolStripControlHost _controlHost;

        #endregion

        #region Events

        /// <summary>
        /// Occurs when a color is selected in the hosted control.
        /// </summary>
        public event EventHandler<Color> ColorSelected;

        #endregion

        #region Constructor

        /// <summary>
        /// Initializes a new instance of the <see cref="ExcelColorDropDown"/> class.
        /// </summary>
        /// <param name="initialColor">The initial color to display in the color picker.</param>
        public ExcelColorDropDown(Color initialColor)
        {
            AutoSize = false;
            DoubleBuffered = true;
            ResizeRedraw = true;
            Padding = Padding.Empty;
            _hostedControl = new ExcelColorDropDownControl(initialColor);
            _hostedControl.ColorSelected += (s, c) =>
            {
                ColorSelected?.Invoke(this, c);
                Close();
            };
            _controlHost = new ToolStripControlHost(_hostedControl)
            {
                Margin = Padding.Empty,
                Padding = Padding.Empty,
                AutoSize = false,
                Size = _hostedControl.Size
            };
            Items.Add(_controlHost);
            Size = _hostedControl.Size;
        }

        #endregion

        #region Protected Overrides

        /// <summary>
        /// Raises the <see cref="ToolStripDropDown.Opening"/> event.
        /// </summary>
        /// <param name="e">A <see cref="CancelEventArgs"/> that contains the event data.</param>
        protected override void OnOpening(CancelEventArgs e)
        {
            base.OnOpening(e);
            _hostedControl.Focus();
        }

        #endregion
    }

    /// <summary>
    /// Provides static utility methods for color conversions and safety checks.
    /// </summary>
    public static class ColorUtils
    {
        #region Fields

        private const double Epsilon = 0.0001;

        #endregion

        #region Web-Safe Helpers

        /// <summary>
        /// Determines whether the specified color is a web-safe color.
        /// </summary>
        /// <param name="color">The color to check.</param>
        /// <returns>true if the color is web-safe; otherwise, false.</returns>
        public static bool IsWebSafe(Color color)
        {
            return (color.R % 51 == 0) && (color.G % 51 == 0) && (color.B % 51 == 0);
        }

        /// <summary>
        /// Converts a color to its nearest web-safe equivalent.
        /// </summary>
        /// <param name="color">The color to convert.</param>
        /// <returns>The nearest web-safe color.</returns>
        public static Color ToWebSafe(Color color)
        {
            int r = (int)(Math.Round(color.R / 51.0) * 51);
            int g = (int)(Math.Round(color.G / 51.0) * 51);
            int b = (int)(Math.Round(color.B / 51.0) * 51);
            return Color.FromArgb(255, r, g, b);
        }

        #endregion

        #region CMYK-Safe Helpers

        /// <summary>
        /// Determines whether a color is CMYK-safe, meaning it can be converted from RGB to CMYK and back without loss.
        /// </summary>
        /// <param name="color">The color to check.</param>
        /// <returns>true if the color is CMYK-safe; otherwise, false.</returns>
        public static bool IsCmykSafe(Color color)
        {
            (int c, int m, int y, int k) = RgbToCmyk(color);
            Color roundTrippedColor = CmykToColor(c, m, y, k);
            return color.ToArgb() == roundTrippedColor.ToArgb();
        }

        /// <summary>
        /// Converts a color to its nearest CMYK-safe equivalent by performing a round-trip conversion.
        /// </summary>
        /// <param name="color">The color to convert.</param>
        /// <returns>The nearest CMYK-safe color.</returns>
        public static Color ToCmykSafe(Color color)
        {
            (int c, int m, int y, int k) = RgbToCmyk(color);
            return CmykToColor(c, m, y, k);
        }

        #endregion

        #region HSV Converters

        /// <summary>
        /// Creates a <see cref="Color"/> from hue, saturation, and value components.
        /// </summary>
        /// <param name="hue">The hue component (0-360).</param>
        /// <param name="saturation">The saturation component (0-1).</param>
        /// <param name="value">The value (brightness) component (0-1).</param>
        /// <returns>A <see cref="Color"/> structure.</returns>
        public static Color ColorFromHsv(double hue, double saturation, double value)
        {
            int hi = Convert.ToInt32(Math.Floor(hue / 60)) % 6;
            double f = hue / 60 - Math.Floor(hue / 60);
            value *= 255;
            int v = Convert.ToInt32(value);
            int p = Convert.ToInt32(value * (1 - saturation));
            int q = Convert.ToInt32(value * (1 - f * saturation));
            int t = Convert.ToInt32(value * (1 - (1 - f) * saturation));
            switch (hi)
            {
                case 0: return Color.FromArgb(255, v, t, p);
                case 1: return Color.FromArgb(255, q, v, p);
                case 2: return Color.FromArgb(255, p, v, t);
                case 3: return Color.FromArgb(255, p, q, v);
                case 4: return Color.FromArgb(255, t, p, v);
                default: return Color.FromArgb(255, v, p, q);
            }
        }

        /// <summary>
        /// Converts a <see cref="Color"/> to its hue, saturation, and value (HSV) components.
        /// </summary>
        /// <param name="color">The color to convert.</param>
        /// <returns>A tuple containing the H (0-360), S (0-1), and V (0-1) components.</returns>
        public static (double H, double S, double V) ColorToHsv(Color color)
        {
            double r = color.R / 255d;
            double g = color.G / 255d;
            double b = color.B / 255d;
            double max = Math.Max(r, Math.Max(g, b));
            double min = Math.Min(r, Math.Min(g, b));
            double delta = max - min;
            double h = 0;
            if (delta < Epsilon)
            {
                h = 0;
            }
            else if (Math.Abs(max - r) < Epsilon)
            {
                h = (60 * ((g - b) / delta) + 360) % 360;
            }
            else if (Math.Abs(max - g) < Epsilon)
            {
                h = (60 * ((b - r) / delta) + 120) % 360;
            }
            else if (Math.Abs(max - b) < Epsilon)
            {
                h = (60 * ((r - g) / delta) + 240) % 360;
            }
            double s = (max < Epsilon) ? 0 : delta / max;
            double v = max;
            return (h, s, v);
        }

        #endregion

        #region HSL Converters

        /// <summary>
        /// Converts an RGB color to its Hue, Saturation, and Lightness (HSL) components.
        /// </summary>
        /// <param name="color">The color to convert.</param>
        /// <returns>A tuple containing the H (0-360), S (0-1), and L (0-1) components.</returns>
        public static (double H, double S, double L) RgbToHsl(Color color)
        {
            double r = color.R / 255.0;
            double g = color.G / 255.0;
            double b = color.B / 255.0;
            double max = Math.Max(r, Math.Max(g, b));
            double min = Math.Min(r, Math.Min(g, b));
            double h = 0, s;
            double l = (max + min) / 2.0;
            if (Math.Abs(max - min) < Epsilon)
            {
                h = 0;
                s = 0;
            }
            else
            {
                double d = max - min;
                s = l > 0.5 ? d / (2.0 - max - min) : d / (max + min);
                if (Math.Abs(max - r) < Epsilon)
                {
                    h = (g - b) / d + (g < b ? 6.0 : 0.0);
                }
                else if (Math.Abs(max - g) < Epsilon)
                {
                    h = (b - r) / d + 2.0;
                }
                else if (Math.Abs(max - b) < Epsilon)
                {
                    h = (r - g) / d + 4.0;
                }
                h /= 6.0;
            }
            return (h * 360, s, l);
        }

        /// <summary>
        /// Creates a <see cref="Color"/> from Hue, Saturation, and Lightness (HSL) components.
        /// </summary>
        /// <param name="h">The hue component (0-360).</param>
        /// <param name="s">The saturation component (0-100).</param>
        /// <param name="l">The lightness component (0-100).</param>
        /// <returns>A <see cref="Color"/> structure.</returns>
        public static Color HslToColor(int h, int s, int l)
        {
            float h_f = h / 360.0f;
            float s_f = s / 100.0f;
            float l_f = l / 100.0f;
            float r, g, b;
            if (Math.Abs(s_f) < Epsilon)
            {
                r = g = b = l_f;
            }
            else
            {
                float q = l_f < 0.5f ? l_f * (1 + s_f) : l_f + s_f - l_f * s_f;
                float p = 2 * l_f - q;
                r = HueToRgb(p, q, h_f + 1.0f / 3.0f);
                g = HueToRgb(p, q, h_f);
                b = HueToRgb(p, q, h_f - 1.0f / 3.0f);
            }
            return Color.FromArgb(255, (int)Math.Round(r * 255), (int)Math.Round(g * 255), (int)Math.Round(b * 255));
        }

        /// <summary>
        /// Helper function for HSL to RGB conversion.
        /// </summary>
        private static float HueToRgb(float p, float q, float t)
        {
            if (t < 0)
            {
                t += 1;
            }
            if (t > 1)
            {
                t -= 1;
            }
            if (t < 1.0f / 6.0f)
            {
                return p + (q - p) * 6 * t;
            }
            if (t < 1.0f / 2.0f)
            {
                return q;
            }
            if (t < 2.0f / 3.0f)
            {
                return p + (q - p) * (2.0f / 3.0f - t) * 6;
            }
            return p;
        }

        #endregion

        #region CMYK Converters

        /// <summary>
        /// Converts an RGB color to its CMYK (Cyan, Magenta, Yellow, Key/Black) components.
        /// </summary>
        /// <param name="color">The color to convert.</param>
        /// <returns>A tuple containing the C, M, Y, and K components (each 0-100).</returns>
        public static (int C, int M, int Y, int K) RgbToCmyk(Color color)
        {
            double r = color.R / 255.0;
            double g = color.G / 255.0;
            double b = color.B / 255.0;
            double k = 1.0 - Math.Max(r, Math.Max(g, b));
            if (Math.Abs(1.0 - k) < Epsilon)
            {
                return (0, 0, 0, 100);
            }
            double c = (1.0 - r - k) / (1.0 - k);
            double m = (1.0 - g - k) / (1.0 - k);
            double y = (1.0 - b - k) / (1.0 - k);
            return (
                (int)Math.Round(c * 100),
                (int)Math.Round(m * 100),
                (int)Math.Round(y * 100),
                (int)Math.Round(k * 100)
            );
        }

        /// <summary>
        /// Creates a <see cref="Color"/> from CMYK components.
        /// </summary>
        /// <param name="c">The cyan component (0-100).</param>
        /// <param name="m">The magenta component (0-100).</param>
        /// <param name="y">The yellow component (0-100).</param>
        /// <param name="k">The key/black component (0-100).</param>
        /// <returns>A <see cref="Color"/> structure.</returns>
        public static Color CmykToColor(int c, int m, int y, int k)
        {
            float c_f = c / 100.0f;
            float m_f = m / 100.0f;
            float y_f = y / 100.0f;
            float k_f = k / 100.0f;
            int r = (int)Math.Round(255 * (1 - c_f) * (1 - k_f));
            int g = (int)Math.Round(255 * (1 - m_f) * (1 - k_f));
            int b = (int)Math.Round(255 * (1 - y_f) * (1 - k_f));
            return Color.FromArgb(255, r, g, b);
        }

        #endregion
    }

    /// <summary>
    /// A circular color picker control for selecting hue and saturation.
    /// </summary>
    public class CircularColorPicker : Control
    {
        #region Fields

        private Color _finalColor = Color.White;
        private Color _positionColor = Color.White;
        private bool _isMouseDown = false;
        private double _idealHue;
        private double _idealSat;
        private bool _isUpdatingFromSelf = false;
        private int _selectorRadius = 7;
        private int _radius;
        private int _centerX;
        private int _centerY;
        private Bitmap _colorWheelBitmap;
        private bool _isKeyDown;

        #endregion

        #region Events

        /// <summary>
        /// Occurs when a color is picked by the user.
        /// </summary>
        public event EventHandler<Color> ColorPicked;

        #endregion

        #region Properties

        /// <summary>
        /// Gets or sets the final color, which includes the brightness component. This is the color shown in the selector circle.
        /// </summary>
        public Color FinalColor
        {
            get => _finalColor;
            set
            {
                if (_finalColor.ToArgb() != value.ToArgb())
                {
                    _finalColor = value;
                    Invalidate();
                }
            }
        }

        /// <summary>
        /// Gets or sets the color that determines the selector's position (hue and saturation) on the wheel.
        /// This is the "pure" color without the brightness component.
        /// </summary>
        public Color PositionColor
        {
            get => _positionColor;
            set
            {
                if (_positionColor.ToArgb() != value.ToArgb())
                {
                    _positionColor = value;
                    if (!_isUpdatingFromSelf)
                    {
                        (_idealHue, _idealSat, _) = ColorUtils.ColorToHsv(value);
                    }
                    Invalidate();
                }
            }
        }

        /// <summary>
        /// Gets or sets a value indicating whether to snap to web-safe colors.
        /// </summary>
        public bool ShowWebSafeOnly { get; set; }

        /// <summary>
        /// Gets or sets a value indicating whether to snap to print-safe (CMYK) colors.
        /// </summary>
        public bool ShowPrintSafeOnly { get; set; }

        #endregion

        #region Constructor

        /// <summary>
        /// Initializes a new instance of the <see cref="CircularColorPicker"/> class.
        /// </summary>
        public CircularColorPicker()
        {
            DoubleBuffered = true;
            SetStyle(ControlStyles.AllPaintingInWmPaint | ControlStyles.OptimizedDoubleBuffer | ControlStyles.ResizeRedraw | ControlStyles.UserPaint | ControlStyles.Selectable, true);
            TabStop = true;
            MouseDown += (s, e) => { Focus(); _isMouseDown = true; _selectorRadius = 10; UpdateColorFromMouse(e.Location); Invalidate(); };
            MouseUp += (s, e) =>
            {
                _isMouseDown = false;
                _selectorRadius = 7;
                Invalidate();
            };
            MouseMove += (s, e) => { if (_isMouseDown) { UpdateColorFromMouse(e.Location); } };
        }

        #endregion

        #region Public Methods

        /// <summary>
        /// Regenerates the color wheel bitmap. This should be called when safety modes change.
        /// </summary>
        public void RegenerateBitmap()
        {
            _colorWheelBitmap?.Dispose();
            _colorWheelBitmap = GenerateColorWheelBitmap();
            Invalidate();
        }

        #endregion

        #region Protected Overrides

        /// <summary>
        /// Determines whether the specified key is a regular input key or a special key that requires preprocessing.
        /// </summary>
        /// <param name="keyData">One of the <see cref="Keys"/> values.</param>
        /// <returns>true if the specified key is an input key; otherwise, false.</returns>
        protected override bool IsInputKey(Keys keyData)
        {
            switch (keyData)
            {
                case Keys.Right:
                case Keys.Left:
                case Keys.Up:
                case Keys.Down:
                    return true;
            }
            return base.IsInputKey(keyData);
        }

        /// <summary>
        /// Raises the <see cref="Control.KeyDown"/> event.
        /// </summary>
        /// <param name="e">A <see cref="KeyEventArgs"/> that contains the event data.</param>
        protected override void OnKeyDown(KeyEventArgs e)
        {
            base.OnKeyDown(e);
            if (e.KeyCode < Keys.Left || e.KeyCode > Keys.Down)
            {
                return;
            }
            if (!_isKeyDown)
            {
                _isKeyDown = true;
            }
            if (ShowWebSafeOnly || ShowPrintSafeOnly)
            {
                Color startColor = this.FinalColor;
                double currentHue = _idealHue;
                double currentSat = _idealSat;
                const double SearchStep = 0.01;
                const int MaxIterations = 500;
                for (int i = 0; i < MaxIterations; i++)
                {
                    double angleRad = currentHue * Math.PI / 180.0;
                    double x = currentSat * Math.Cos(angleRad);
                    double y = currentSat * Math.Sin(angleRad);
                    switch (e.KeyCode)
                    {
                        case Keys.Left: x -= SearchStep; break;
                        case Keys.Right: x += SearchStep; break;
                        case Keys.Up: y += SearchStep; break;
                        case Keys.Down: y -= SearchStep; break;
                    }
                    currentSat = Math.Sqrt(x * x + y * y);
                    if (currentSat > 1.0)
                    {
                        currentSat = 1.0;
                    }
                    if (currentSat > 0.001)
                    {
                        currentHue = Math.Atan2(y, x) * 180.0 / Math.PI;
                        if (currentHue < 0) currentHue += 360.0;
                    }
                    Color rawColor = ColorUtils.ColorFromHsv(currentHue, currentSat, 1.0);
                    Color newSnappedColor = ShowWebSafeOnly
                        ? ColorUtils.ToWebSafe(rawColor)
                        : ColorUtils.ToCmykSafe(rawColor);
                    if (newSnappedColor.ToArgb() != startColor.ToArgb())
                    {
                        _idealHue = currentHue;
                        _idealSat = currentSat;
                        _isUpdatingFromSelf = true;
                        ColorPicked?.Invoke(this, newSnappedColor);
                        _isUpdatingFromSelf = false;
                        e.Handled = true;
                        return;
                    }
                }
                e.Handled = true;
                return;
            }
            double original_angleRad = _idealHue * Math.PI / 180.0;
            double original_x = _idealSat * Math.Cos(original_angleRad);
            double original_y = _idealSat * Math.Sin(original_angleRad);
            const double Step = 0.05;
            switch (e.KeyCode)
            {
                case Keys.Left: original_x -= Step; break;
                case Keys.Right: original_x += Step; break;
                case Keys.Up: original_y += Step; break;
                case Keys.Down: original_y -= Step; break;
            }
            _idealSat = Math.Min(1.0, Math.Sqrt(original_x * original_x + original_y * original_y));
            _idealHue = Math.Atan2(original_y, original_x) * 180.0 / Math.PI;
            if (_idealHue < 0)
            {
                _idealHue += 360.0;
            }
            Color newColor = ColorUtils.ColorFromHsv(_idealHue, _idealSat, 1.0);
            _isUpdatingFromSelf = true;
            ColorPicked?.Invoke(this, newColor);
            _isUpdatingFromSelf = false;
            e.Handled = true;
        }

        /// <summary>
        /// Raises the <see cref="Control.KeyUp"/> event.
        /// </summary>
        /// <param name="e">A <see cref="KeyEventArgs"/> that contains the event data.</param>
        protected override void OnKeyUp(KeyEventArgs e)
        {
            base.OnKeyUp(e);
            if (e.KeyCode >= Keys.Left && e.KeyCode <= Keys.Down)
            {
                _isKeyDown = false;
                Invalidate();
            }
        }

        /// <summary>
        /// Raises the <see cref="Control.GotFocus"/> event.
        /// </summary>
        /// <param name="e">An <see cref="EventArgs"/> that contains the event data.</param>
        protected override void OnGotFocus(EventArgs e) { base.OnGotFocus(e); Invalidate(); }

        /// <summary>
        /// Raises the <see cref="Control.LostFocus"/> event.
        /// </summary>
        /// <param name="e">An <see cref="EventArgs"/> that contains the event data.</param>
        protected override void OnLostFocus(EventArgs e) { base.OnLostFocus(e); Invalidate(); }

        /// <summary>
        /// Raises the <see cref="Control.Resize"/> event.
        /// </summary>
        /// <param name="e">An <see cref="EventArgs"/> that contains the event data.</param>
        protected override void OnResize(EventArgs e)
        {
            base.OnResize(e);
            int minDim = Math.Min(Width, Height);
            _radius = Math.Max(0, (minDim - 20) / 2);
            _centerX = Width / 2;
            _centerY = Height / 2;
            RegenerateBitmap();
        }

        /// <summary>
        /// Paints the background of the control.
        /// </summary>
        /// <param name="e">A <see cref="PaintEventArgs"/> that contains the event data.</param>
        protected override void OnPaint(PaintEventArgs e)
        {
            base.OnPaint(e);
            e.Graphics.SmoothingMode = SmoothingMode.AntiAlias;
            if (_colorWheelBitmap != null)
            {
                e.Graphics.DrawImage(_colorWheelBitmap, 0, 0);
            }
            using (Pen borderPen = new Pen(Color.Gray, 2))
            {
                e.Graphics.DrawEllipse(borderPen, _centerX - _radius, _centerY - _radius, 2 * _radius, 2 * _radius);
            }
            double selHue, selSat;
            if (_isMouseDown || _isKeyDown)
            {
                selHue = _idealHue;
                selSat = _idealSat;
            }
            else
            {
                (selHue, selSat, _) = ColorUtils.ColorToHsv(PositionColor);
            }
            double selAngle = selHue * Math.PI / 180.0;
            double selRadius = _radius * selSat;
            int selX = _centerX + (int)(selRadius * Math.Cos(selAngle));
            int selY = _centerY - (int)(selRadius * Math.Sin(selAngle));
            Rectangle selectorRect = new Rectangle(selX - _selectorRadius, selY - _selectorRadius, _selectorRadius * 2, _selectorRadius * 2);
            using (SolidBrush selBrush = new SolidBrush(FinalColor))
            {
                e.Graphics.FillEllipse(selBrush, selectorRect);
            }
            using (Pen whitePen = new Pen(Color.White, 1))
            {
                e.Graphics.DrawEllipse(whitePen, selectorRect.X - 1, selectorRect.Y - 1, selectorRect.Width + 2, selectorRect.Height + 2);
            }
            using (Pen blackPen = new Pen(Color.Black, 1))
            {
                e.Graphics.DrawEllipse(blackPen, selectorRect);
            }
        }

        /// <summary>
        /// Releases the unmanaged resources used by the <see cref="Control"/> and optionally releases the managed resources.
        /// </summary>
        /// <param name="disposing">true to release both managed and unmanaged resources; false to release only unmanaged resources.</param>
        protected override void Dispose(bool disposing)
        {
            if (disposing)
            {
                _colorWheelBitmap?.Dispose();
            }
            base.Dispose(disposing);
        }

        #endregion

        #region Private Methods

        /// <summary>
        /// Generates the color wheel bitmap based on current size and safety settings.
        /// </summary>
        /// <returns>A new <see cref="Bitmap"/> of the color wheel.</returns>
        private Bitmap GenerateColorWheelBitmap()
        {
            if (Width <= 0 || Height <= 0 || _radius <= 0)
            {
                return null;
            }
            const int Scale = 2;
            int largeWidth = Width * Scale;
            int largeHeight = Height * Scale;
            int largeCenterX = _centerX * Scale;
            int largeCenterY = _centerY * Scale;
            int largeRadius = _radius * Scale;
            using (Bitmap largeBmp = new Bitmap(largeWidth, largeHeight, System.Drawing.Imaging.PixelFormat.Format32bppArgb))
            {
                for (int y = 0; y < largeHeight; ++y)
                {
                    for (int x = 0; x < largeWidth; ++x)
                    {
                        int dx = x - largeCenterX;
                        int dy = y - largeCenterY;
                        double dist = Math.Sqrt(dx * dx + dy * dy);
                        if (dist > largeRadius)
                        {
                            largeBmp.SetPixel(x, y, Color.Transparent);
                            continue;
                        }
                        double sat = dist / largeRadius;
                        double hue = Math.Atan2(-dy, dx) * 180.0 / Math.PI;
                        if (hue < 0)
                        {
                            hue += 360.0;
                        }
                        Color c = ColorUtils.ColorFromHsv(hue, sat, 1.0);
                        if (ShowPrintSafeOnly)
                        {
                            c = ColorUtils.ToCmykSafe(c);
                        }
                        else if (ShowWebSafeOnly)
                        {
                            c = ColorUtils.ToWebSafe(c);
                        }
                        largeBmp.SetPixel(x, y, c);
                    }
                }
                Bitmap finalBmp = new Bitmap(Width, Height, System.Drawing.Imaging.PixelFormat.Format32bppArgb);
                using (Graphics g = Graphics.FromImage(finalBmp))
                {
                    g.SmoothingMode = SmoothingMode.AntiAlias;
                    g.InterpolationMode = InterpolationMode.HighQualityBicubic;
                    g.DrawImage(largeBmp, new Rectangle(0, 0, Width, Height));
                }
                return finalBmp;
            }
        }

        /// <summary>
        /// Updates the selected color based on the mouse pointer's location.
        /// </summary>
        /// <param name="p">The location of the mouse pointer.</param>
        private void UpdateColorFromMouse(Point p)
        {
            int dx = p.X - _centerX;
            int dy = p.Y - _centerY;
            double dist = Math.Sqrt(dx * dx + dy * dy);
            dist = Math.Min(dist, _radius);
            _idealSat = dist / _radius;
            _idealHue = Math.Atan2(-dy, dx) * 180.0 / Math.PI;
            if (_idealHue < 0)
            {
                _idealHue += 360.0;
            }
            Color rawColor = ColorUtils.ColorFromHsv(_idealHue, _idealSat, 1.0);
            if (ShowWebSafeOnly)
            {
                rawColor = ColorUtils.ToWebSafe(rawColor);
            }
            else if (ShowPrintSafeOnly)
            {
                rawColor = ColorUtils.ToCmykSafe(rawColor);
            }
            _isUpdatingFromSelf = true;
            ColorPicked?.Invoke(this, rawColor);
            _isUpdatingFromSelf = false;
        }

        #endregion
    }

    /// <summary>
    /// A vertical slider control for adjusting color brightness (value in HSV model).
    /// </summary>
    public class BrightnessSlider : Control
    {
        #region Fields

        private double _value = 1.0;
        private Color _baseColor = Color.Red;
        private bool _isWebSafeMode;
        private List<Color> _webSafeColorSteps;
        private bool _isMouseDown = false;
        private readonly ToolTip _toolTip;
        private double _valueOnInteractionStart;
        private bool _isKeyDown;

        #endregion

        #region Events

        /// <summary>
        /// Occurs when the <see cref="Value"/> property changes.
        /// </summary>
        public event EventHandler ValueChanged;

        #endregion

        #region Properties

        /// <summary>
        /// Gets or sets the current brightness value of the slider, from 0.0 (black) to 1.0 (full color/white).
        /// </summary>
        public double Value
        {
            get => _value;
            set
            {
                double newValue = Math.Max(0, Math.Min(1, value));
                if (IsWebSafeMode && _webSafeColorSteps != null && _webSafeColorSteps.Count > 1)
                {
                    int numSteps = _webSafeColorSteps.Count;
                    int stepIndex = (int)Math.Round(newValue * (numSteps - 1));
                    newValue = (double)stepIndex / (numSteps - 1);
                }
                if (Math.Abs(_value - newValue) > 0.0001)
                {
                    _value = newValue;
                    Invalidate();
                    ValueChanged?.Invoke(this, EventArgs.Empty);
                }
            }
        }

        /// <summary>
        /// Gets or sets the base color (with full brightness) used to create the gradient.
        /// </summary>
        public Color BaseColor
        {
            get => _baseColor;
            set
            {
                if (_baseColor.ToArgb() == value.ToArgb())
                {
                    return;
                }
                _baseColor = value;
                if (IsWebSafeMode)
                {
                    RecalculateWebSafeSteps();
                    Value = _value;
                }
                Invalidate();
            }
        }

        /// <summary>
        /// Gets or sets a value indicating whether the slider should snap to web-safe color values.
        /// </summary>
        public bool IsWebSafeMode
        {
            get => _isWebSafeMode;
            set
            {
                if (_isWebSafeMode == value)
                {
                    return;
                }
                _isWebSafeMode = value;
                RecalculateWebSafeSteps();
                Value = _value;
                Invalidate();
            }
        }

        /// <summary>
        /// Gets or sets a value indicating whether the slider should snap to print-safe color values.
        /// </summary>
        public bool IsPrintSafeMode { get; set; }

        #endregion

        #region Constructor

        /// <summary>
        /// Initializes a new instance of the <see cref="BrightnessSlider"/> class.
        /// </summary>
        public BrightnessSlider()
        {
            DoubleBuffered = true;
            SetStyle(ControlStyles.AllPaintingInWmPaint | ControlStyles.UserPaint | ControlStyles.ResizeRedraw | ControlStyles.Selectable, true);
            TabStop = true;
            _toolTip = new ToolTip();
            RecalculateWebSafeSteps();
        }

        #endregion

        #region Public Methods

        /// <summary>
        /// Gets the final selected color based on the current <see cref="BaseColor"/> and <see cref="Value"/>.
        /// </summary>
        /// <returns>The final calculated <see cref="Color"/>.</returns>
        public Color GetSelectedColor()
        {
            if (IsWebSafeMode && _webSafeColorSteps != null && _webSafeColorSteps.Count > 0)
            {
                int numSteps = _webSafeColorSteps.Count;
                if (numSteps == 1)
                {
                    return _webSafeColorSteps[0];
                }
                int stepIndex = (int)Math.Round((1.0 - _value) * (numSteps - 1));
                stepIndex = Math.Max(0, Math.Min(numSteps - 1, stepIndex));
                return _webSafeColorSteps[stepIndex];
            }
            (double h, double s, _) = ColorUtils.ColorToHsv(BaseColor);
            Color result = ColorUtils.ColorFromHsv(h, s, _value);
            if (IsPrintSafeMode)
            {
                result = ColorUtils.ToCmykSafe(result);
            }
            return result;
        }

        #endregion

        #region Private Helper Methods

        /// <summary>
        /// Recalculates the discrete color steps available when in web-safe mode.
        /// </summary>
        private void RecalculateWebSafeSteps()
        {
            if (!IsWebSafeMode)
            {
                _webSafeColorSteps = null;
                return;
            }
            HashSet<Color> webColors = new HashSet<Color>();
            (double baseHue, double baseSat, _) = ColorUtils.ColorToHsv(BaseColor);
            for (int i = 0; i <= 100; i++)
            {
                double v = i / 100.0;
                Color testColor = ColorUtils.ColorFromHsv(baseHue, baseSat, v);
                webColors.Add(ColorUtils.ToWebSafe(testColor));
            }
            List<Color> candidateColors;
            if (baseSat < 0.1)
            {
                candidateColors = webColors.Where(c => c.R == c.G && c.G == c.B).ToList();
            }
            else
            {
                const double HueTolerance = 45.0;
                double GetHueDiff(double h1, double h2)
                {
                    double diff = Math.Abs(h1 - h2);
                    return Math.Min(diff, 360 - diff);
                }
                candidateColors = webColors.Where(c =>
                {
                    (double cH, double cS, _) = ColorUtils.ColorToHsv(c);
                    if (cS < 0.15)
                    {
                        return true;
                    }
                    return GetHueDiff(cH, baseHue) < HueTolerance;
                }).ToList();
            }
            _webSafeColorSteps = candidateColors
                .OrderByDescending(c => ColorUtils.RgbToHsl(c).L)
                .ToList();
        }

        /// <summary>
        /// Snaps the current slider value to the nearest print-safe color value after user interaction.
        /// </summary>
        private void SnapToPrintSafeValue()
        {
            if (!IsPrintSafeMode)
            {
                return;
            }
            (double h, double s, _) = ColorUtils.ColorToHsv(BaseColor);
            double currentValue = Value;
            Color colorFromValue = ColorUtils.ColorFromHsv(h, s, currentValue);
            if (ColorUtils.IsCmykSafe(colorFromValue))
            {
                return;
            }
            double direction = currentValue - _valueOnInteractionStart;
            if (Math.Abs(direction) < 0.001)
            {
                Color snappedColor = ColorUtils.ToCmykSafe(colorFromValue);
                (_, _, double finalValue) = ColorUtils.ColorToHsv(snappedColor);
                if (Math.Abs(Value - finalValue) > 0.0001)
                {
                    Value = finalValue;
                }
                return;
            }
            double step = direction > 0 ? 0.001 : -0.001;
            double searchValue = currentValue;
            for (int i = 0; i < 1000 && searchValue >= 0.0 && searchValue <= 1.0; i++)
            {
                Color testColor = ColorUtils.ColorFromHsv(h, s, searchValue);
                if (ColorUtils.IsCmykSafe(testColor))
                {
                    if (Math.Abs(Value - searchValue) > 0.0001)
                    {
                        Value = searchValue;
                    }
                    return;
                }
                searchValue += step;
            }
            Color fallbackSnappedColor = ColorUtils.ToCmykSafe(colorFromValue);
            (_, _, double fallbackFinalValue) = ColorUtils.ColorToHsv(fallbackSnappedColor);
            if (Math.Abs(Value - fallbackFinalValue) > 0.0001)
            {
                Value = fallbackFinalValue;
            }
        }

        /// <summary>
        /// Shows a tooltip with the current brightness value (0-100).
        /// </summary>
        private void ShowValueToolTip()
        {
            if (IsWebSafeMode)
            {
                return;
            }
            string tipText = $"{Math.Round(Value * 100)}";
            const int Margin = 5;
            int availableHeight = Height - Margin * 2;
            int selectorY = Margin + (int)((1.0 - Value) * availableHeight);
            _toolTip.Show(tipText, this, Width + 5, selectorY - 10, 2000);
        }

        /// <summary>
        /// Updates the slider's value based on the mouse's Y coordinate.
        /// </summary>
        /// <param name="y">The Y coordinate of the mouse.</param>
        private void UpdateValueFromMouse(int y)
        {
            const int Margin = 5;
            int availableHeight = Height - Margin * 2;
            int clampedY = Math.Max(Margin, Math.Min(y, Height - Margin));
            double newValue = 1.0 - ((double)(clampedY - Margin) / availableHeight);
            Value = newValue;
            ShowValueToolTip();
        }

        #endregion

        #region Overridden Events

        /// <summary>
        /// Raises the <see cref="Control.MouseDown"/> event.
        /// </summary>
        /// <param name="e">A <see cref="MouseEventArgs"/> that contains the event data.</param>
        protected override void OnMouseDown(MouseEventArgs e)
        {
            base.OnMouseDown(e);
            if (e.Button == MouseButtons.Left)
            {
                Focus();
                _isMouseDown = true;
                _valueOnInteractionStart = Value;
                UpdateValueFromMouse(e.Y);
            }
        }

        /// <summary>
        /// Raises the <see cref="Control.MouseMove"/> event.
        /// </summary>
        /// <param name="e">A <see cref="MouseEventArgs"/> that contains the event data.</param>
        protected override void OnMouseMove(MouseEventArgs e)
        {
            base.OnMouseMove(e);
            if (_isMouseDown)
            {
                UpdateValueFromMouse(e.Y);
            }
        }

        /// <summary>
        /// Raises the <see cref="Control.MouseUp"/> event.
        /// </summary>
        /// <param name="e">A <see cref="MouseEventArgs"/> that contains the event data.</param>
        protected override void OnMouseUp(MouseEventArgs e)
        {
            base.OnMouseUp(e);
            if (e.Button == MouseButtons.Left && _isMouseDown)
            {
                _isMouseDown = false;
                _toolTip.Hide(this);
                SnapToPrintSafeValue();
            }
        }

        /// <summary>
        /// Raises the <see cref="Control.KeyUp"/> event.
        /// </summary>
        /// <param name="e">A <see cref="KeyEventArgs"/> that contains the event data.</param>
        protected override void OnKeyUp(KeyEventArgs e)
        {
            base.OnKeyUp(e);
            if (IsInputKey(e.KeyData) && _isKeyDown)
            {
                _isKeyDown = false;
                SnapToPrintSafeValue();
            }
        }

        /// <summary>
        /// Determines whether the specified key is a regular input key or a special key that requires preprocessing.
        /// </summary>
        /// <param name="keyData">One of the <see cref="Keys"/> values.</param>
        /// <returns>true if the specified key is an input key; otherwise, false.</returns>
        protected override bool IsInputKey(Keys keyData)
        {
            switch (keyData)
            {
                case Keys.Right:
                case Keys.Left:
                case Keys.Up:
                case Keys.Down:
                case Keys.PageUp:
                case Keys.PageDown:
                case Keys.Home:
                case Keys.End:
                    return true;
            }
            return base.IsInputKey(keyData);
        }

        /// <summary>
        /// Raises the <see cref="Control.KeyDown"/> event.
        /// </summary>
        /// <param name="e">A <see cref="KeyEventArgs"/> that contains the event data.</param>
        protected override void OnKeyDown(KeyEventArgs e)
        {
            if (IsInputKey(e.KeyData) && !_isKeyDown)
            {
                _isKeyDown = true;
                _valueOnInteractionStart = Value;
            }
            base.OnKeyDown(e);
            double step = 0.01;
            double pageStep = 0.1;
            if (IsWebSafeMode && _webSafeColorSteps != null && _webSafeColorSteps.Count > 1)
            {
                step = 1.0 / (_webSafeColorSteps.Count - 1);
                pageStep = step;
            }
            switch (e.KeyCode)
            {
                case Keys.Up: case Keys.Right: Value += step; break;
                case Keys.Down: case Keys.Left: Value -= step; break;
                case Keys.PageUp: Value += pageStep; break;
                case Keys.PageDown: Value -= pageStep; break;
                case Keys.Home: Value = 1.0; break;
                case Keys.End: Value = 0.0; break;
                default: return;
            }
            ShowValueToolTip();
            e.Handled = true;
        }

        /// <summary>
        /// Raises the <see cref="Control.GotFocus"/> event.
        /// </summary>
        /// <param name="e">An <see cref="EventArgs"/> that contains the event data.</param>
        protected override void OnGotFocus(EventArgs e) { base.OnGotFocus(e); Invalidate(); }

        /// <summary>
        /// Raises the <see cref="Control.LostFocus"/> event.
        /// </summary>
        /// <param name="e">An <see cref="EventArgs"/> that contains the event data.</param>
        protected override void OnLostFocus(EventArgs e)
        {
            base.OnLostFocus(e);
            _toolTip.Hide(this);
            if (_isMouseDown || _isKeyDown)
            {
                _isMouseDown = false;
                _isKeyDown = false;
                SnapToPrintSafeValue();
            }
            Invalidate();
        }

        /// <summary>
        /// Raises the <see cref="Control.Paint"/> event.
        /// </summary>
        /// <param name="e">A <see cref="PaintEventArgs"/> that contains the event data.</param>
        protected override void OnPaint(PaintEventArgs e)
        {
            base.OnPaint(e);
            e.Graphics.SmoothingMode = SmoothingMode.AntiAlias;
            const int TrackWidth = 14;
            const int Margin = 5;
            int centerX = Width / 2;
            Rectangle trackRect = new Rectangle(centerX - (TrackWidth / 2), Margin, TrackWidth, Height - Margin * 2);
            if (IsWebSafeMode && _webSafeColorSteps != null && _webSafeColorSteps.Count > 0)
            {
                int numColors = _webSafeColorSteps.Count;
                float blockHeight = (float)trackRect.Height / numColors;
                for (int i = 0; i < numColors; i++)
                {
                    Color c = _webSafeColorSteps[i];
                    float yPos = trackRect.Top + (i * blockHeight);
                    float rectHeight = (float)Math.Ceiling(blockHeight);
                    if (yPos + rectHeight > trackRect.Bottom)
                    {
                        rectHeight = trackRect.Bottom - yPos;
                    }
                    if (rectHeight <= 0)
                    {
                        continue;
                    }
                    using (SolidBrush b = new SolidBrush(c))
                    {
                        e.Graphics.FillRectangle(b, trackRect.X, yPos, trackRect.Width, rectHeight);
                    }
                }
            }
            else
            {
                using (LinearGradientBrush brush = new LinearGradientBrush(trackRect, Color.White, Color.Black, LinearGradientMode.Vertical))
                {
                    brush.InterpolationColors = new ColorBlend(3)
                    {
                        Colors = new[] { Color.White, BaseColor, Color.Black },
                        Positions = new[] { 0f, 0.5f, 1f }
                    };
                    e.Graphics.FillRectangle(brush, trackRect);
                }
            }
            e.Graphics.DrawRectangle(SystemPens.ControlDark, trackRect);
            int selectorY;
            if (IsWebSafeMode && _webSafeColorSteps != null && _webSafeColorSteps.Count > 0)
            {
                int numSteps = _webSafeColorSteps.Count;
                int stepIndex = (int)Math.Round((1.0 - Value) * (numSteps - 1));
                stepIndex = Math.Max(0, Math.Min(numSteps - 1, stepIndex));
                float blockHeight = (float)trackRect.Height / numSteps;
                selectorY = trackRect.Top + (int)Math.Round(stepIndex * blockHeight + blockHeight / 2.0f);
            }
            else
            {
                int availableHeight = Height - Margin * 2;
                selectorY = Margin + (int)Math.Round((1.0 - Value) * availableHeight);
            }
            const int TriangleHeight = 8;
            const int TriangleWidth = 4;
            const int TriangleHalfHeight = TriangleHeight / 2;
            Point[] leftTriangle = { new Point(trackRect.Left, selectorY - TriangleHalfHeight), new Point(trackRect.Left, selectorY + TriangleHalfHeight), new Point(trackRect.Left + TriangleWidth, selectorY) };
            Point[] rightTriangle = { new Point(trackRect.Right, selectorY - TriangleHalfHeight), new Point(trackRect.Right, selectorY + TriangleHalfHeight), new Point(trackRect.Right - TriangleWidth, selectorY) };
            using (SolidBrush whiteBrush = new SolidBrush(Color.White))
            {
                e.Graphics.FillPolygon(whiteBrush, leftTriangle);
                e.Graphics.FillPolygon(whiteBrush, rightTriangle);
            }
            using (Pen blackPen = new Pen(Color.Black))
            {
                e.Graphics.DrawPolygon(blackPen, leftTriangle);
                e.Graphics.DrawPolygon(blackPen, rightTriangle);
            }
        }

        /// <summary>
        /// Releases the unmanaged resources used by the <see cref="Control"/> and optionally releases the managed resources.
        /// </summary>
        /// <param name="disposing">true to release both managed and unmanaged resources; false to release only unmanaged resources.</param>
        protected override void Dispose(bool disposing)
        {
            if (disposing)
            {
                _toolTip?.Dispose();
            }
            base.Dispose(disposing);
        }

        #endregion
    }

    /// <summary>
    /// A custom <see cref="TextBox"/> that vertically centers its text content for single-line text boxes.
    /// </summary>
    public class VerticallyCenteredTextBox : TextBox
    {
        #region Win32 API

        private const int EM_SETRECT = 0xB3;

        [System.Runtime.InteropServices.DllImport("user32.dll")]
        private static extern IntPtr SendMessage(IntPtr hWnd, int msg, int wParam, ref RECT lParam);

        [System.Runtime.InteropServices.StructLayout(System.Runtime.InteropServices.LayoutKind.Sequential)]
        private struct RECT { public int Left, Top, Right, Bottom; }

        #endregion

        #region Protected Overrides

        /// <summary>
        /// Raises the <see cref="Control.HandleCreated"/> event.
        /// </summary>
        /// <param name="e">An <see cref="EventArgs"/> that contains the event data.</param>
        protected override void OnHandleCreated(EventArgs e)
        {
            base.OnHandleCreated(e);
            SetVerticalPadding();
        }

        /// <summary>
        /// Raises the <see cref="Control.Resize"/> event.
        /// </summary>
        /// <param name="e">An <see cref="EventArgs"/> that contains the event data.</param>
        protected override void OnResize(EventArgs e)
        {
            base.OnResize(e);
            SetVerticalPadding();
        }

        /// <summary>
        /// Raises the <see cref="Control.FontChanged"/> event.
        /// </summary>
        /// <param name="e">An <see cref="EventArgs"/> that contains the event data.</param>
        protected override void OnFontChanged(EventArgs e)
        {
            base.OnFontChanged(e);
            SetVerticalPadding();
        }

        #endregion

        #region Private Methods

        /// <summary>
        /// Sets the internal formatting rectangle of the TextBox to center the text vertically.
        /// </summary>
        private void SetVerticalPadding()
        {
            if (!IsDisposed && IsHandleCreated && !Multiline)
            {
                int textHeight = TextRenderer.MeasureText("X", Font).Height;
                int paddingTop = Math.Max(1, (ClientSize.Height - textHeight) / 2);
                RECT rect = new RECT { Left = 3, Top = paddingTop, Right = ClientSize.Width - 3, Bottom = ClientSize.Height };
                SendMessage(Handle, EM_SETRECT, 0, ref rect);
            }
        }

        #endregion
    }

    /// <summary>
    /// The main user control for the Excel-style color picker. It combines the circular picker,
    /// brightness slider, color grid, and value text boxes into a single control.
    /// </summary>
    public class ExcelColorDropDownControl : UserControl
    {
        #region Fields and Constants

        private static readonly Color[][] GroupedPalette =
        {
            new[] { Color.FromArgb(0,0,0), Color.FromArgb(64,64,64), Color.FromArgb(128,128,128), Color.FromArgb(192,192,192), Color.FromArgb(224,224,224), Color.FromArgb(255,255,255) },
            new[] { Color.FromArgb(183,28,28), Color.FromArgb(198,40,40), Color.FromArgb(229,57,53), Color.FromArgb(239,154,154), Color.FromArgb(255,205,210), Color.FromArgb(255,235,238) },
            new[] { Color.FromArgb(230,81,0), Color.FromArgb(239,108,0), Color.FromArgb(251,140,0), Color.FromArgb(255,183,77), Color.FromArgb(255,224,178), Color.FromArgb(255,243,224) },
            new[] { Color.FromArgb(245,127,23), Color.FromArgb(251,192,45), Color.FromArgb(253,216,53), Color.FromArgb(255,241,118), Color.FromArgb(255,249,196), Color.FromArgb(255,253,231) },
            new[] { Color.FromArgb(27,94,32), Color.FromArgb(46,125,50), Color.FromArgb(67,160,71), Color.FromArgb(129,199,132), Color.FromArgb(200,230,201), Color.FromArgb(232,245,233) },
            new[] { Color.FromArgb(0,96,100), Color.FromArgb(0,131,143), Color.FromArgb(0,172,193), Color.FromArgb(77,208,225), Color.FromArgb(178,235,242), Color.FromArgb(224,247,250) },
            new[] { Color.FromArgb(13,71,161), Color.FromArgb(21,101,192), Color.FromArgb(30,136,229), Color.FromArgb(100,181,246), Color.FromArgb(187,222,251), Color.FromArgb(227,242,253) },
            new[] { Color.FromArgb(74,20,140), Color.FromArgb(106,27,154), Color.FromArgb(142,36,170), Color.FromArgb(186,104,200), Color.FromArgb(225,190,231), Color.FromArgb(243,229,245) }
        };
        private const int PaletteGroupCount = 8;
        private const int ColorsPerGroup = 6;
        private const int CellSize = 14;
        private const int GridPadding = 3;
        private const int GridCellPadding = 1;
        private readonly IContainer _components = null;
        private readonly Font _textBoxFont;
        private readonly CircularColorPicker _circularColorPicker;
        private readonly BrightnessSlider _brightnessSlider;
        private readonly Panel _originalColorPanel;
        private readonly Panel _currentColorPanel;
        private readonly Button _moreButton;
        private readonly Button _okButton;
        private readonly ToolTip _toolTip;
        private readonly DoubleBufferedPanel _gridPanel;
        private readonly Label _labelRgb, _labelCmyk, _labelHex;
        private readonly TextBox _textBoxHex;
        private readonly TextBox[] _textBoxesRgb, _textBoxesCmyk;
        private readonly Label _webSafeLabel;
        private readonly Label _printSafeLabel;
        private readonly CheckBox _webSafeCheckBox;
        private readonly CheckBox _printSafeCheckBox;
        private Color _selectedColor;
        private readonly Color _originalColor;
        private int _hoveredGroup = -1, _hoveredIndex = -1;
        private int _selectedGroup = -1, _selectedIndex = -1;
        private bool _isUpdatingInternally = false;
        private bool _isHoveringOriginal = false;

        #endregion

        #region Events

        /// <summary>
        /// Occurs when the user clicks the OK button, finalizing the color selection.
        /// </summary>
        public event EventHandler<Color> ColorSelected;

        #endregion

        #region Constructor

        /// <summary>
        /// Initializes a new instance of the <see cref="ExcelColorDropDownControl"/> class.
        /// </summary>
        /// <param name="initialColor">The color to initialize the control with.</param>
        public ExcelColorDropDownControl(Color initialColor)
        {
            _components = new Container();
            SetStyle(ControlStyles.AllPaintingInWmPaint | ControlStyles.OptimizedDoubleBuffer | ControlStyles.UserPaint, true);
            _originalColor = initialColor;
            const int Spacing = 6;
            const int CircularHeight = 195;
            const int SliderWidth = 15;
            const int PickerSliderPadding = 8;
            _textBoxFont = new Font("Segoe UI", 7f);
            _circularColorPicker = new CircularColorPicker { Left = 0, Top = 0, Width = CircularHeight + (Spacing * 2), Height = CircularHeight + (Spacing * 2) };
            _webSafeLabel = new Label { Font = _textBoxFont, AutoSize = true, BackColor = Color.Transparent, Location = new Point(Spacing, Spacing + 1) };
            _printSafeLabel = new Label { Font = _textBoxFont, AutoSize = true, BackColor = Color.Transparent };
            Controls.Add(_webSafeLabel);
            Controls.Add(_printSafeLabel);
            _brightnessSlider = new BrightnessSlider { Left = _circularColorPicker.Right + PickerSliderPadding - 7, Top = Spacing - 4, Width = SliderWidth, Height = CircularHeight + (Spacing) };
            Controls.Add(_circularColorPicker);
            Controls.Add(_brightnessSlider);
            int gridLeft = _brightnessSlider.Right + PickerSliderPadding;
            int gridWidth = (PaletteGroupCount * CellSize) + ((PaletteGroupCount - 1) * GridPadding) + (GridCellPadding * 2);
            int gridHeight = (ColorsPerGroup * CellSize) + ((ColorsPerGroup - 1) * GridPadding) + (GridCellPadding * 2);
            _gridPanel = new DoubleBufferedPanel { Left = gridLeft, Top = Spacing, Width = gridWidth, Height = gridHeight, TabStop = true };
            Controls.Add(_gridPanel);
            _webSafeCheckBox = new CheckBox { Text = "Web-safe", Font = _textBoxFont, AutoSize = true };
            _printSafeCheckBox = new CheckBox { Text = "Print-safe", Font = _textBoxFont, AutoSize = true };
            int checkY = _gridPanel.Bottom + Spacing / 2;
            _webSafeCheckBox.Location = new Point(_gridPanel.Left + 1, checkY);
            _printSafeCheckBox.Location = new Point(_gridPanel.Left + 1 + (4 * (CellSize + GridPadding)), checkY);
            Controls.Add(_webSafeCheckBox);
            Controls.Add(_printSafeCheckBox);
            int currentY = _webSafeCheckBox.Bottom;
            const int TextBoxHeight = 16;
            const int RowHeight = TextBoxHeight;
            const int VerticalPadding = 2;
            int labelWidth = (CellSize * 2) + Spacing + 1;
            const int TextBoxPadding = 2;
            int availableWidth = gridWidth - labelWidth;
            _labelHex = new Label { Text = "HEX", Left = gridLeft, Top = currentY, Size = new Size(labelWidth, TextBoxHeight), Font = _textBoxFont, TextAlign = ContentAlignment.MiddleLeft };
            int txtBoxLeft = _labelHex.Right;
            _textBoxHex = new VerticallyCenteredTextBox { Left = txtBoxLeft, Top = currentY, Width = availableWidth, Font = _textBoxFont, Height = TextBoxHeight, AutoSize = false };
            Controls.Add(_labelHex);
            Controls.Add(_textBoxHex);
            currentY += RowHeight + VerticalPadding;
            _labelRgb = new Label { Text = "RGB", Left = gridLeft, Top = currentY, Size = new Size(labelWidth, TextBoxHeight), Font = _textBoxFont, TextAlign = ContentAlignment.MiddleLeft };
            int rgbTxtWidth = (availableWidth - (2 * TextBoxPadding)) / 3;
            _textBoxesRgb = new TextBox[3];
            for (int i = 0; i < 3; i++) { _textBoxesRgb[i] = new VerticallyCenteredTextBox { Left = txtBoxLeft + i * (rgbTxtWidth + TextBoxPadding), Top = currentY, Width = rgbTxtWidth, Font = _textBoxFont, Height = TextBoxHeight, AutoSize = false }; }
            _textBoxesRgb[2].Width = availableWidth - (2 * (rgbTxtWidth + TextBoxPadding));
            Controls.Add(_labelRgb);
            Controls.AddRange(_textBoxesRgb);
            currentY += RowHeight + VerticalPadding;
            _labelCmyk = new Label { Text = "CMYK", Left = gridLeft, Top = currentY, Size = new Size(labelWidth, TextBoxHeight), Font = _textBoxFont, TextAlign = ContentAlignment.MiddleLeft };
            int cmykTxtWidth = (availableWidth - (3 * TextBoxPadding)) / 4;
            _textBoxesCmyk = new TextBox[4];
            for (int i = 0; i < 4; i++) { _textBoxesCmyk[i] = new VerticallyCenteredTextBox { Left = txtBoxLeft + i * (cmykTxtWidth + TextBoxPadding), Top = currentY, Width = cmykTxtWidth, Font = _textBoxFont, Height = TextBoxHeight, AutoSize = false }; }
            _textBoxesCmyk[3].Width = availableWidth - (3 * (cmykTxtWidth + TextBoxPadding));
            Controls.Add(_labelCmyk);
            Controls.AddRange(_textBoxesCmyk);
            _currentColorPanel = new Panel { BackColor = initialColor, TabStop = true };
            _originalColorPanel = new Panel { BackColor = _originalColor, TabStop = true };
            _moreButton = new Button { Text = "&More", Font = _textBoxFont, Padding = new Padding(0), FlatStyle = FlatStyle.System };
            _okButton = new Button { Text = "&OK", Font = _textBoxFont, Padding = new Padding(0), FlatStyle = FlatStyle.System };
            Controls.Add(_originalColorPanel);
            Controls.Add(_currentColorPanel);
            Controls.Add(_moreButton);
            Controls.Add(_okButton);
            Width = _gridPanel.Right + Spacing + 1;
            BackColor = SystemColors.Window;
            _toolTip = new ToolTip(_components) { AutomaticDelay = 200, AutoPopDelay = 10000 };
            WireEvents();
            ArrangeLayout();
            UpdateControlsFromColor(initialColor);
        }

        #endregion

        #region Overrides

        /// <summary>
        /// Clean up any resources being used.
        /// </summary>
        /// <param name="disposing">true if managed resources should be disposed; otherwise, false.</param>
        protected override void Dispose(bool disposing)
        {
            if (disposing && (_components != null))
            {
                _components.Dispose();
                _textBoxFont?.Dispose();
            }
            base.Dispose(disposing);
        }

        #endregion

        #region Layout and Initialization

        /// <summary>
        /// Wires up all the event handlers for the child controls.
        /// </summary>
        private void WireEvents()
        {
            _webSafeLabel.Click += WebSafeLabel_Click;
            _printSafeLabel.Click += PrintSafeLabel_Click;
            _toolTip.SetToolTip(_currentColorPanel, "New Color");
            _toolTip.SetToolTip(_originalColorPanel, "Current Color (Click to restore)");
            _circularColorPicker.ColorPicked += (s, c) => { _brightnessSlider.BaseColor = c; UpdateFromPickerInteraction(); };
            _brightnessSlider.ValueChanged += (s, e) => UpdateFromPickerInteraction();
            _gridPanel.Paint += Grid_Paint;
            _gridPanel.MouseClick += Grid_MouseClick;
            _gridPanel.MouseMove += Grid_MouseMove;
            _gridPanel.KeyDown += Grid_KeyDown;
            _gridPanel.GotFocus += (s, e) => _gridPanel.Invalidate();
            _gridPanel.LostFocus += (s, e) => _gridPanel.Invalidate();
            _gridPanel.MouseLeave += (s, e) => { _hoveredGroup = -1; _hoveredIndex = -1; _toolTip.SetToolTip(_gridPanel, null); _gridPanel.Invalidate(); };
            _currentColorPanel.Paint += ColorPanel_Paint;
            _currentColorPanel.GotFocus += (s, e) => _currentColorPanel.Invalidate();
            _currentColorPanel.LostFocus += (s, e) => _currentColorPanel.Invalidate();
            _currentColorPanel.Click += (s, e) => _currentColorPanel.Focus();
            _originalColorPanel.Paint += ColorPanel_Paint;
            _originalColorPanel.Click += (s, e) => { _originalColorPanel.Focus(); UpdateControlsFromColor(_originalColor); };
            _originalColorPanel.KeyDown += (s, e) => { if (e.KeyCode == Keys.Enter || e.KeyCode == Keys.Space) { UpdateControlsFromColor(_originalColor); e.Handled = true; } };
            _originalColorPanel.MouseEnter += (s, e) => { _isHoveringOriginal = true; _originalColorPanel.Invalidate(); };
            _originalColorPanel.MouseLeave += (s, e) => { _isHoveringOriginal = false; _originalColorPanel.Invalidate(); };
            _originalColorPanel.GotFocus += (s, e) => _originalColorPanel.Invalidate();
            _originalColorPanel.LostFocus += (s, e) => _originalColorPanel.Invalidate();
            _moreButton.Click += MoreButton_Click;
            _okButton.Click += (s, e) => ColorSelected?.Invoke(this, _selectedColor);
            _webSafeCheckBox.CheckedChanged += CheckBox_CheckedChanged;
            _printSafeCheckBox.CheckedChanged += CheckBox_CheckedChanged;
            _brightnessSlider.IsWebSafeMode = _webSafeCheckBox.Checked;
            _textBoxHex.Validated += TextBox_Validated;
            foreach (TextBox txt in _textBoxesRgb) { txt.Validated += TextBox_Validated; txt.KeyDown += TextBox_Value_KeyDown; }
            foreach (TextBox txt in _textBoxesCmyk) { txt.Validated += TextBox_Validated; txt.KeyDown += TextBox_Value_KeyDown; }
        }

        /// <summary>
        /// Arranges the final row of controls (color panels and buttons).
        /// </summary>
        private void ArrangeLayout()
        {
            const int Spacing = 6;
            int bottomRowTop = Spacing + _labelCmyk.Bottom;
            int buttonHeight = CellSize + 1;
            _currentColorPanel.Top = bottomRowTop;
            _currentColorPanel.Left = _gridPanel.Left + 1;
            _currentColorPanel.Width = CellSize + 1;
            _currentColorPanel.Height = buttonHeight;
            _originalColorPanel.Top = bottomRowTop;
            _originalColorPanel.Left = _currentColorPanel.Right + GridPadding - 1;
            _originalColorPanel.Width = CellSize + 1;
            _originalColorPanel.Height = buttonHeight;
            int buttonsLeft = _originalColorPanel.Right + GridCellPadding;
            int buttonsRightEdge = _textBoxesRgb.Last().Right;
            int buttonsAvailableWidth = buttonsRightEdge - buttonsLeft;
            int gap = GridCellPadding - 1;
            int moreButtonWidth = (buttonsAvailableWidth - gap) / 2;
            int buttonTop = bottomRowTop - 1;
            int btnHeight = buttonHeight + 2;
            _moreButton.Top = buttonTop;
            _moreButton.Left = buttonsLeft;
            _moreButton.Width = moreButtonWidth;
            _moreButton.Height = btnHeight;
            _okButton.Top = buttonTop;
            _okButton.Left = _moreButton.Right + gap;
            _okButton.Width = buttonsRightEdge - _okButton.Left + 1;
            _okButton.Height = btnHeight;
            Height = _okButton.Bottom + Spacing;
        }

        #endregion

        #region Event Handlers

        /// <summary>
        /// Handles the CheckedChanged event for the web-safe and print-safe check boxes.
        /// </summary>
        private void CheckBox_CheckedChanged(object sender, EventArgs e)
        {
            if (!(sender is CheckBox changedCheckBox) || _isUpdatingInternally)
            {
                return;
            }
            if (changedCheckBox.Checked)
            {
                _isUpdatingInternally = true;
                if (changedCheckBox == _webSafeCheckBox && _printSafeCheckBox.Checked)
                {
                    _printSafeCheckBox.Checked = false;
                }
                else if (changedCheckBox == _printSafeCheckBox && _webSafeCheckBox.Checked)
                {
                    _webSafeCheckBox.Checked = false;
                }
                _isUpdatingInternally = false;
            }
            if (_webSafeCheckBox.Checked) { _printSafeLabel.Visible = false; _webSafeLabel.Visible = true; }
            else if (_printSafeCheckBox.Checked) { _webSafeLabel.Visible = false; _printSafeLabel.Visible = true; }
            else { _webSafeLabel.Visible = true; _printSafeLabel.Visible = true; }
            _circularColorPicker.ShowWebSafeOnly = _webSafeCheckBox.Checked;
            _circularColorPicker.ShowPrintSafeOnly = _printSafeCheckBox.Checked;
            _brightnessSlider.IsWebSafeMode = _webSafeCheckBox.Checked;
            _brightnessSlider.IsPrintSafeMode = _printSafeCheckBox.Checked;
            _circularColorPicker.RegenerateBitmap();
            _gridPanel.Invalidate();
            Color currentColor = _selectedColor;
            if (_webSafeCheckBox.Checked)
            {
                UpdateControlsFromColor(ColorUtils.ToWebSafe(currentColor));
            }
            else if (_printSafeCheckBox.Checked)
            {
                UpdateControlsFromColor(ColorUtils.ToCmykSafe(currentColor));
            }
            else
            {
                UpdateControlsFromColor(currentColor);
            }
        }

        /// <summary>
        /// Handles clicks on the web-safe warning label to snap the color.
        /// </summary>
        private void WebSafeLabel_Click(object sender, EventArgs e)
        {
            if (!ColorUtils.IsWebSafe(_selectedColor))
            {
                UpdateControlsFromColor(ColorUtils.ToWebSafe(_selectedColor));
            }
        }

        /// <summary>
        /// Handles clicks on the print-safe warning label to snap the color.
        /// </summary>
        private void PrintSafeLabel_Click(object sender, EventArgs e)
        {
            if (!ColorUtils.IsCmykSafe(_selectedColor))
            {
                UpdateControlsFromColor(ColorUtils.ToCmykSafe(_selectedColor));
            }
        }

        /// <summary>
        /// Handles the Validated event for the color value text boxes.
        /// </summary>
        private void TextBox_Validated(object sender, EventArgs e)
        {
            if (_isUpdatingInternally)
            {
                return;
            }
            try
            {
                Color newColor = _selectedColor;
                if (sender == _textBoxHex)
                {
                    string hex = _textBoxHex.Text.TrimStart('#');
                    if (hex.Length != 6)
                    {
                        throw new FormatException("Hex code must be 6 characters long.");
                    }
                    newColor = Color.FromArgb(255,
                        int.Parse(hex.Substring(0, 2), NumberStyles.HexNumber),
                        int.Parse(hex.Substring(2, 2), NumberStyles.HexNumber),
                        int.Parse(hex.Substring(4, 2), NumberStyles.HexNumber));
                }
                else if (_textBoxesRgb.Contains(sender))
                {
                    newColor = Color.FromArgb(
                        ParseTextBoxValue(_textBoxesRgb[0].Text, 0, 255),
                        ParseTextBoxValue(_textBoxesRgb[1].Text, 0, 255),
                        ParseTextBoxValue(_textBoxesRgb[2].Text, 0, 255));
                }
                else if (_textBoxesCmyk.Contains(sender))
                {
                    newColor = ColorUtils.CmykToColor(
                        ParseTextBoxValue(_textBoxesCmyk[0].Text, 0, 100),
                        ParseTextBoxValue(_textBoxesCmyk[1].Text, 0, 100),
                        ParseTextBoxValue(_textBoxesCmyk[2].Text, 0, 100),
                        ParseTextBoxValue(_textBoxesCmyk[3].Text, 0, 100));
                }
                if (_printSafeCheckBox.Checked)
                {
                    newColor = ColorUtils.ToCmykSafe(newColor);
                }
                else if (_webSafeCheckBox.Checked)
                {
                    newColor = ColorUtils.ToWebSafe(newColor);
                }
                UpdateControlsFromColor(newColor);
            }
            catch (FormatException) { UpdateValueTextBoxes(); }
            catch (OverflowException) { UpdateValueTextBoxes(); }
        }

        /// <summary>
        /// Handles the KeyDown event for numeric text boxes to increment/decrement values with arrow keys.
        /// </summary>
        private void TextBox_Value_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode != Keys.Up && e.KeyCode != Keys.Down)
            {
                return;
            }
            if (_webSafeCheckBox.Checked || _printSafeCheckBox.Checked)
            {
                return;
            }
            e.Handled = true;
            e.SuppressKeyPress = true;
            if (!(sender is TextBox txt))
            {
                return;
            }
            if (!int.TryParse(txt.Text, out int currentValue))
            {
                currentValue = 0;
            }
            int change = e.KeyCode == Keys.Up ? 1 : -1;
            int newValue = currentValue + change;
            int max = _textBoxesRgb.Contains(txt) ? 255 : 100;
            newValue = Math.Max(0, Math.Min(max, newValue));
            if (newValue != currentValue || txt.Text != newValue.ToString())
            {
                txt.Text = newValue.ToString();
                txt.SelectionStart = txt.Text.Length;
                TextBox_Validated(sender, EventArgs.Empty);
            }
        }

        /// <summary>
        /// Handles the MouseClick event for the color grid panel.
        /// </summary>
        private void Grid_MouseClick(object sender, MouseEventArgs e)
        {
            if (sender is Control panel)
            {
                panel.Focus();
            }
            (int group, int idx) = HitTest(new Point(e.X - GridCellPadding, e.Y - GridCellPadding));
            if (group >= 0 && idx >= 0)
            {
                UpdateControlsFromColor(GetGridColor(group, idx));
            }
        }

        /// <summary>
        /// Handles the KeyDown event for the color grid panel for keyboard navigation.
        /// </summary>
        private void Grid_KeyDown(object sender, KeyEventArgs e)
        {
            switch (e.KeyCode)
            {
                case Keys.Left:
                case Keys.Right:
                case Keys.Up:
                case Keys.Down:
                case Keys.PageUp:
                case Keys.PageDown:
                case Keys.Home:
                case Keys.End:
                    e.Handled = true; break;
                default: return;
            }
            if (_selectedGroup == -1 || _selectedIndex == -1)
            {
                _selectedGroup = 0; _selectedIndex = 0;
            }
            else
            {
                switch (e.KeyCode)
                {
                    case Keys.Right: _selectedGroup++; break;
                    case Keys.Left: _selectedGroup--; break;
                    case Keys.Down: _selectedIndex++; break;
                    case Keys.Up: _selectedIndex--; break;
                    case Keys.Home: _selectedGroup = 0; break;
                    case Keys.End: _selectedGroup = PaletteGroupCount - 1; break;
                    case Keys.PageUp: _selectedIndex = 0; break;
                    case Keys.PageDown: _selectedIndex = ColorsPerGroup - 1; break;
                }
            }
            _selectedGroup = (_selectedGroup + PaletteGroupCount) % PaletteGroupCount;
            _selectedIndex = (_selectedIndex + ColorsPerGroup) % ColorsPerGroup;
            UpdateControlsFromColor(GetGridColor(_selectedGroup, _selectedIndex));
        }

        /// <summary>
        /// Handles the Click event for the "More" button, which opens the system ColorDialog.
        /// </summary>
        private void MoreButton_Click(object sender, EventArgs e)
        {
            using (ColorDialog dlg = new ColorDialog { Color = _selectedColor, FullOpen = true })
            {
                if (dlg.ShowDialog(this) == DialogResult.OK)
                {
                    Color selected = dlg.Color;
                    if (_printSafeCheckBox.Checked)
                    {
                        selected = ColorUtils.ToCmykSafe(selected);
                    }
                    else if (_webSafeCheckBox.Checked)
                    {
                        selected = ColorUtils.ToWebSafe(selected);
                    }
                    UpdateControlsFromColor(selected);
                }
            }
        }

        /// <summary>
        /// Handles the MouseMove event for the color grid panel to update hover effects and tooltips.
        /// </summary>
        private void Grid_MouseMove(object sender, MouseEventArgs e)
        {
            (int group, int idx) = HitTest(new Point(e.X - GridCellPadding, e.Y - GridCellPadding));
            if (group != _hoveredGroup || idx != _hoveredIndex)
            {
                _hoveredGroup = group; _hoveredIndex = idx;
                if (_hoveredGroup >= 0 && _hoveredIndex >= 0)
                {
                    Color c = GetGridColor(_hoveredGroup, _hoveredIndex);
                    (int ck, int mk, int yk, int kk) = ColorUtils.RgbToCmyk(c);
                    string tip = $"HEX: #{c.R:X2}{c.G:X2}{c.B:X2}\n" +
                                 $"RGB: {c.R}, {c.G}, {c.B}\n" +
                                 $"CMYK: {ck}, {mk}, {yk}, {kk}";
                    _toolTip.SetToolTip(_gridPanel, tip);
                }
                else
                {
                    _toolTip.SetToolTip(_gridPanel, null);
                }
                _gridPanel.Invalidate();
            }
        }

        #endregion

        #region Color Update Logic

        /// <summary>
        /// Gets the color from the grid palette at the specified coordinates, applying safety modes if active.
        /// </summary>
        /// <param name="group">The column index (0-7).</param>
        /// <param name="index">The row index (0-5).</param>
        /// <returns>The color at the specified grid location.</returns>
        private Color GetGridColor(int group, int index)
        {
            Color baseColor = GroupedPalette[group][index];
            if (_webSafeCheckBox.Checked)
            {
                return ColorUtils.ToWebSafe(baseColor);
            }
            if (_printSafeCheckBox.Checked)
            {
                return ColorUtils.ToCmykSafe(baseColor);
            }
            return baseColor;
        }

        /// <summary>
        /// Updates the state of the web-safe and print-safe warning labels.
        /// </summary>
        private void UpdateSafetyWarnings()
        {
            if (ColorUtils.IsWebSafe(_selectedColor)) { _webSafeLabel.Text = "✔ Web"; _webSafeLabel.ForeColor = SystemColors.ControlText; _webSafeLabel.Cursor = Cursors.Default; _toolTip.SetToolTip(_webSafeLabel, "This color is web-safe."); }
            else { _webSafeLabel.Text = "⚠ Web"; _webSafeLabel.ForeColor = Color.Red; _webSafeLabel.Cursor = Cursors.Hand; _toolTip.SetToolTip(_webSafeLabel, "This color is not web-safe. Click to fix."); }
            if (ColorUtils.IsCmykSafe(_selectedColor)) { _printSafeLabel.Text = "✔ Print"; _printSafeLabel.ForeColor = SystemColors.ControlText; _printSafeLabel.Cursor = Cursors.Default; _toolTip.SetToolTip(_printSafeLabel, "This color is print-safe."); }
            else { _printSafeLabel.Text = "⚠ Print"; _printSafeLabel.ForeColor = Color.Red; _printSafeLabel.Cursor = Cursors.Hand; _toolTip.SetToolTip(_printSafeLabel, "This color is not print-safe. Click to fix."); }
            _printSafeLabel.Location = new Point(6, _circularColorPicker.Height - (_printSafeLabel.Height * 3 / 2) - 2);
        }

        /// <summary>
        /// Updates all child controls to reflect a new color.
        /// </summary>
        /// <param name="c">The new color to display.</param>
        private void UpdateControlsFromColor(Color c)
        {
            _isUpdatingInternally = true;
            _selectedColor = c;
            (double hue, double sat, double val) = ColorUtils.ColorToHsv(c);
            _brightnessSlider.Value = val;
            Color pureColor = ColorUtils.ColorFromHsv(hue, sat, 1.0);
            Color snappedPureColor = pureColor;
            if (_webSafeCheckBox.Checked)
            {
                snappedPureColor = ColorUtils.ToWebSafe(pureColor);
            }
            else if (_printSafeCheckBox.Checked)
            {
                snappedPureColor = ColorUtils.ToCmykSafe(pureColor);
            }
            if (_printSafeCheckBox.Checked && !ColorUtils.IsCmykSafe(c))
            {
                c = ColorUtils.ToCmykSafe(c);
                _selectedColor = c;
                (hue, sat, val) = ColorUtils.ColorToHsv(c);
                snappedPureColor = ColorUtils.ColorFromHsv(hue, sat, 1.0);
                snappedPureColor = ColorUtils.ToCmykSafe(snappedPureColor);
                _brightnessSlider.Value = val;
            }
            _brightnessSlider.BaseColor = snappedPureColor;
            _circularColorPicker.PositionColor = snappedPureColor;
            _circularColorPicker.FinalColor = c;
            if (_currentColorPanel != null)
            {
                _currentColorPanel.BackColor = c;
            }
            _selectedGroup = -1; _selectedIndex = -1;
            for (int g = 0; g < PaletteGroupCount; g++)
            {
                for (int i = 0; i < ColorsPerGroup; i++)
                {
                    if (GetGridColor(g, i).ToArgb() == c.ToArgb())
                    {
                        _selectedGroup = g; _selectedIndex = i;
                        break;
                    }
                }
                if (_selectedGroup != -1)
                {
                    break;
                }
            }
            _gridPanel.Invalidate();
            UpdateValueTextBoxes();
            UpdateSafetyWarnings();
            _isUpdatingInternally = false;
        }

        /// <summary>
        /// Updates controls based on interaction with the circular picker or brightness slider.
        /// </summary>
        private void UpdateFromPickerInteraction()
        {
            if (_isUpdatingInternally)
            {
                return;
            }
            _selectedColor = _brightnessSlider.GetSelectedColor();
            if (_printSafeCheckBox.Checked && !ColorUtils.IsCmykSafe(_selectedColor))
            {
                _selectedColor = ColorUtils.ToCmykSafe(_selectedColor);
            }
            _currentColorPanel.BackColor = _selectedColor;
            _circularColorPicker.PositionColor = _brightnessSlider.BaseColor;
            _circularColorPicker.FinalColor = _selectedColor;
            _selectedGroup = -1;
            _selectedIndex = -1;
            _gridPanel.Invalidate();
            UpdateValueTextBoxes();
            UpdateSafetyWarnings();
        }

        /// <summary>
        /// Updates the HEX, RGB, and CMYK text boxes with the current selected color's values.
        /// </summary>
        private void UpdateValueTextBoxes()
        {
            Color c = _selectedColor;
            _textBoxHex.Text = $"#{c.R:X2}{c.G:X2}{c.B:X2}";
            _textBoxesRgb[0].Text = c.R.ToString();
            _textBoxesRgb[1].Text = c.G.ToString();
            _textBoxesRgb[2].Text = c.B.ToString();
            (int ck, int mk, int yk, int kk) = ColorUtils.RgbToCmyk(c);
            _textBoxesCmyk[0].Text = ck.ToString();
            _textBoxesCmyk[1].Text = mk.ToString();
            _textBoxesCmyk[2].Text = yk.ToString();
            _textBoxesCmyk[3].Text = kk.ToString();
        }

        /// <summary>
        /// Parses text from a TextBox into an integer, clamping it within a specified range.
        /// </summary>
        private int ParseTextBoxValue(string text, int min, int max)
        {
            return Math.Max(min, Math.Min(max, int.Parse(text)));
        }

        #endregion

        #region Painting

        /// <summary>
        /// Handles the Paint event for the current and original color panels.
        /// </summary>
        private void ColorPanel_Paint(object sender, PaintEventArgs e)
        {
            if (!(sender is Panel p))
            {
                return;
            }
            e.Graphics.Clear(p.Parent.BackColor);
            Rectangle clientRect = new Rectangle(0, 0, p.Width - 1, p.Height - 1);
            using (SolidBrush colorBrush = new SolidBrush(p.BackColor))
            {
                e.Graphics.FillRectangle(colorBrush, clientRect);
            }
            if (p.Focused || (p == _originalColorPanel && _isHoveringOriginal))
            {
                using (Pen borderPen = new Pen(SystemColors.Highlight, 2))
                {
                    e.Graphics.DrawRectangle(borderPen, clientRect);
                }
            }
            else
            {
                e.Graphics.DrawRectangle(SystemPens.ControlDark, clientRect);
            }
        }

        /// <summary>
        /// Handles the Paint event for the color grid panel.
        /// </summary>
        private void Grid_Paint(object sender, PaintEventArgs e)
        {
            e.Graphics.Clear(SystemColors.Window);
            e.Graphics.TranslateTransform(GridCellPadding, GridCellPadding);
            for (int g = 0; g < PaletteGroupCount; ++g)
            {
                for (int i = 0; i < ColorsPerGroup; ++i)
                {
                    Rectangle rect = new Rectangle(g * (CellSize + GridPadding), i * (CellSize + GridPadding), CellSize, CellSize);
                    Color color = GetGridColor(g, i);
                    using (SolidBrush b = new SolidBrush(color))
                    {
                        e.Graphics.FillRectangle(b, rect);
                    }
                    if (g == _selectedGroup && i == _selectedIndex)
                    {
                        using (Pen p = new Pen(SystemColors.Highlight, 2)) { e.Graphics.DrawRectangle(p, rect); }
                    }
                    else if (g == _hoveredGroup && i == _hoveredIndex)
                    {
                        using (Pen p = new Pen(SystemColors.ControlDark, 2)) { e.Graphics.DrawRectangle(p, rect); }
                    }
                    else
                    {
                        e.Graphics.DrawRectangle(SystemPens.ControlDark, rect);
                    }
                }
            }
            e.Graphics.ResetTransform();
        }

        #endregion

        #region Hit Testing

        /// <summary>
        /// Determines which color cell in the grid corresponds to a given point.
        /// </summary>
        /// <param name="pt">The point to test.</param>
        /// <returns>A tuple containing the group (column) and index (row) of the cell, or (-1, -1) if no cell is hit.</returns>
        private (int group, int idx) HitTest(Point pt)
        {
            int g = pt.X / (CellSize + GridPadding);
            int i = pt.Y / (CellSize + GridPadding);
            if (g >= 0 && g < PaletteGroupCount && i >= 0 && i < ColorsPerGroup)
            {
                Rectangle rect = new Rectangle(g * (CellSize + GridPadding), i * (CellSize + GridPadding), CellSize, CellSize);
                if (rect.Contains(pt))
                {
                    return (g, i);
                }
            }
            return (-1, -1);
        }

        #endregion

        #region Nested Classes

        /// <summary>
        /// A simple double-buffered panel that also allows keyboard input.
        /// </summary>
        private class DoubleBufferedPanel : Panel
        {
            /// <summary>
            /// Initializes a new instance of the <see cref="DoubleBufferedPanel"/> class.
            /// </summary>
            public DoubleBufferedPanel() { DoubleBuffered = true; SetStyle(ControlStyles.Selectable, true); }

            /// <summary>
            /// Determines if a key is an input key.
            /// </summary>
            /// <param name="keyData">The key data.</param>
            /// <returns>true if the key is an input key; otherwise, false.</returns>
            protected override bool IsInputKey(Keys keyData)
            {
                switch (keyData)
                {
                    case Keys.Left:
                    case Keys.Right:
                    case Keys.Up:
                    case Keys.Down:
                    case Keys.Home:
                    case Keys.End:
                    case Keys.PageUp:
                    case Keys.PageDown:
                        return true;
                }
                return base.IsInputKey(keyData);
            }
        }

        #endregion
    }
}