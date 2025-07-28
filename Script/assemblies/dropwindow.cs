using System;
using System.Collections.Generic;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Input;
using System.Windows.Media;
using ViewModelHelper;


namespace DropWindow
{
    internal class Theme
    {
        internal static Color backgroundColor = new Color { A = 255, R = 40, G = 40, B = 40 };
        internal static Brush backgroundBrush = new SolidColorBrush(backgroundColor);
        internal static Brush foregroundBrush = new SolidColorBrush(Colors.WhiteSmoke);

        internal static Color dashedBorderColor = new Color { A = 255, R = 60, G = 60, B = 60 };
        internal static Brush dashedBorderBrush = new SolidColorBrush(dashedBorderColor);

        static Theme()
        { }
    }

    public class MainWindow : Window
    {
        public MainWindow(DropWindowViewModel viewModel)
        {
            DataContext = viewModel;

            SetBinding(TitleProperty, new OneWayBinding("Title"));

            Width = 480.0;
            Height = 270.0;
            ShowActivated = true;
            WindowStartupLocation = WindowStartupLocation.Manual;
            Top = 150.0;
            Left = 150.0;
            Topmost = true;
            Background = Theme.backgroundBrush;
            ResizeMode = ResizeMode.NoResize;
            //AllowDrop = true;

            InitializeComponent();
        }

        private void InitializeComponent()
        {
            var dragDropText = new DragDropText();
            dragDropText.SetBinding(DragDropText.DragDropCommandProperty, new OneWayBinding("DropFilesCommand"));

            var mainGrid = new Grid();
            mainGrid.Children.Add(dragDropText);
            Content = mainGrid;
        }
    }

    public class DropWindowViewModel : ViewModelBase
    {
        private string _title = "Drag & Drop";
        public string Title
        {
            get { return _title; }
            set { SetProperty(ref _title, value); }
        }

        private ICommand _dropFilesCommand;
        public ICommand DropFilesCommand
        {
            get { return _dropFilesCommand; }
            set { SetProperty(ref _dropFilesCommand, value); }
        }

        public DropWindowViewModel()
        { }
    }

    internal class DragDropText : UserControl
    {
        private TextBlock _textBlock;

        private Brush _stroke;
        private double _strokeThickness;
        private double _strokeDashLine;
        private double _strokeDashSpace;
        private Brush _Fill;

        public ICommand DragDropCommand
        {
            get { return (ICommand)GetValue(DragDropCommandProperty); }
            set { SetValue(DragDropCommandProperty, value); }
        }
        public static readonly DependencyProperty DragDropCommandProperty =
            DependencyProperty.Register("DragDropCommand", typeof(ICommand), typeof(DragDropText),
                                        new PropertyMetadata(null));

        public DragDropText()
        {
            AllowDrop = true;

            _textBlock = new TextBlock
            {
                Text = "Drag & Drop here",
                Background = Brushes.Transparent,
                Foreground = Theme.foregroundBrush,
                HorizontalAlignment = HorizontalAlignment.Center,
                VerticalAlignment = VerticalAlignment.Center,
                TextAlignment = TextAlignment.Center,
                FontSize = 23.0,
                Margin = new Thickness(5.0),
            };

            SnapsToDevicePixels = true;
            UseLayoutRounding = true;
            Margin = new Thickness(15.0);

            _stroke = Theme.dashedBorderBrush;
            _strokeThickness = 4.0;
            _strokeDashLine = 10.0;
            _strokeDashSpace = 10.0;
            _Fill = Brushes.Transparent;

            Content = _textBlock;
        }

        protected override void OnDrop(System.Windows.DragEventArgs e)
        {
            if (DragDropCommand != null && DragDropCommand.CanExecute(null))
                DragDropCommand.Execute(e);
        }

        protected override void OnRender(DrawingContext drawingContext)
        {
            double w = ActualWidth;
            double h = ActualHeight;
            double x = _strokeThickness / 2.0;

            Pen horizontalPen = GetPen(ActualWidth - 2.0 * x);
            Pen verticalPen = GetPen(ActualHeight - 2.0 * x);

            drawingContext.DrawRectangle(_Fill, null, new Rect(new Point(0, 0), new Size(w, h)));

            drawingContext.DrawLine(horizontalPen, new Point(x, x), new Point(w - x, x));
            drawingContext.DrawLine(horizontalPen, new Point(x, h - x), new Point(w - x, h - x));

            drawingContext.DrawLine(verticalPen, new Point(x, x), new Point(x, h - x));
            drawingContext.DrawLine(verticalPen, new Point(w - x, x), new Point(w - x, h - x));
        }

        private Pen GetPen(double length)
        {
            IEnumerable<double> dashArray = GetDashArray(length);
            return new Pen(_stroke, _strokeThickness)
            {
                DashStyle = new DashStyle(dashArray, 0),
                EndLineCap = PenLineCap.Square,
                StartLineCap = PenLineCap.Square,
                DashCap = PenLineCap.Flat
            };
        }

        private IEnumerable<double> GetDashArray(double length)
        {
            double useableLength = length - _strokeDashLine;
            int lines = (int)Math.Round(useableLength / (_strokeDashLine + _strokeDashSpace));
            useableLength -= lines * _strokeDashLine;
            double actualSpacing = useableLength / lines;

            yield return _strokeDashLine / _strokeThickness;
            yield return actualSpacing / _strokeThickness;
        }
    }
}
