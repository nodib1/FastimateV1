using System;
using System.Threading;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Media;
using System.Windows.Media.Animation;
using System.Windows.Shapes;
using System.Windows.Threading;

namespace FASTIMATE
{
    /// <summary>
    /// Prozorčić "Učitavanje u toku..." sa krugom koji se okreće.
    /// Radi na zasebnoj niti (svoj Dispatcher), pa se krug okreće i dok je glavna nit zauzeta čitanjem iz SolidWorksa.
    /// Upotreba: using (UcitavanjeProzor.Prikazi(this, "Čitanje strukture...")) { ... }
    /// </summary>
    public sealed class UcitavanjeProzor : IDisposable
    {
        private Window _window;
        private Dispatcher _dispatcher;
        private bool _closed;

        private UcitavanjeProzor() { }

        public static UcitavanjeProzor Prikazi(Window owner, string poruka = "")
        {
            var p = new UcitavanjeProzor();
            try
            {
                Rect area = OwnerArea(owner);
                var ready = new ManualResetEventSlim(false);

                var thread = new Thread(() =>
                {
                    try
                    {
                        p._dispatcher = Dispatcher.CurrentDispatcher;
                        p._window = Build(poruka);
                        p._window.Left = area.Left + (area.Width - p._window.Width) / 2;
                        p._window.Top = area.Top + (area.Height - p._window.Height) / 2;
                        p._window.Show();
                    }
                    catch { }
                    finally { ready.Set(); }
                    try { Dispatcher.Run(); } catch { }
                });
                thread.SetApartmentState(ApartmentState.STA);
                thread.IsBackground = true;
                thread.Start();
                ready.Wait(3000);
            }
            catch { }
            return p;
        }

        /// <summary>Područje (u DIP) iznad kojeg se centrira prozorčić.</summary>
        private static Rect OwnerArea(Window owner)
        {
            try
            {
                if (owner != null && owner.IsVisible && owner.WindowState == WindowState.Normal && owner.ActualWidth > 0)
                    return new Rect(owner.Left, owner.Top, owner.ActualWidth, owner.ActualHeight);
            }
            catch { }
            return SystemParameters.WorkArea;
        }

        private static Window Build(string poruka)
        {
            var w = new Window
            {
                Width = 340,
                Height = 150,
                WindowStyle = WindowStyle.None,
                AllowsTransparency = true,
                Background = Brushes.Transparent,
                ResizeMode = ResizeMode.NoResize,
                ShowInTaskbar = false,
                ShowActivated = false,
                Topmost = true,
                WindowStartupLocation = WindowStartupLocation.Manual,
                FontFamily = new FontFamily("Segoe UI")
            };

            // Krug: siva podloga + plavi luk koji se okreće
            var track = new Ellipse { Width = 46, Height = 46, StrokeThickness = 5, Stroke = Brush("#374151") };
            var arc = new Ellipse
            {
                Width = 46,
                Height = 46,
                StrokeThickness = 5,
                Stroke = Brush("#38BDF8"),
                StrokeDashArray = new DoubleCollection { 9, 18 },
                StrokeDashCap = PenLineCap.Round,
                RenderTransformOrigin = new Point(0.5, 0.5)
            };
            var rotate = new RotateTransform();
            arc.RenderTransform = rotate;
            rotate.BeginAnimation(RotateTransform.AngleProperty,
                new DoubleAnimation(0, 360, TimeSpan.FromSeconds(0.9)) { RepeatBehavior = RepeatBehavior.Forever });

            var spinner = new Grid { Width = 46, Height = 46, Margin = new Thickness(0, 0, 18, 0), VerticalAlignment = VerticalAlignment.Center };
            spinner.Children.Add(track);
            spinner.Children.Add(arc);

            var texts = new StackPanel { VerticalAlignment = VerticalAlignment.Center };
            texts.Children.Add(new TextBlock
            {
                Text = "Učitavanje u toku...",
                Foreground = Brushes.White,
                FontSize = 17,
                FontWeight = FontWeights.Bold
            });
            texts.Children.Add(new TextBlock
            {
                Text = string.IsNullOrWhiteSpace(poruka) ? "Molimo sačekajte." : poruka,
                Foreground = Brush("#9CA3AF"),
                FontSize = 12,
                Margin = new Thickness(0, 4, 0, 0),
                TextWrapping = TextWrapping.Wrap,
                MaxWidth = 220
            });

            var row = new StackPanel { Orientation = Orientation.Horizontal, HorizontalAlignment = HorizontalAlignment.Center, VerticalAlignment = VerticalAlignment.Center };
            row.Children.Add(spinner);
            row.Children.Add(texts);

            w.Content = new Border
            {
                Background = Brush("#252830"),
                BorderBrush = Brush("#0288D1"),
                BorderThickness = new Thickness(2),
                CornerRadius = new CornerRadius(12),
                Margin = new Thickness(8),
                Effect = new System.Windows.Media.Effects.DropShadowEffect { BlurRadius = 14, ShadowDepth = 0, Opacity = 0.6 },
                Child = row
            };
            return w;
        }

        private static SolidColorBrush Brush(string hex)
        {
            var b = new SolidColorBrush((Color)ColorConverter.ConvertFromString(hex));
            b.Freeze();
            return b;
        }

        public void Dispose()
        {
            if (_closed) return;
            _closed = true;
            var d = _dispatcher;
            if (d == null) return;
            try
            {
                d.BeginInvoke(new Action(() =>
                {
                    try { _window?.Close(); } catch { }
                    d.InvokeShutdown();
                }));
            }
            catch { }
        }
    }
}
