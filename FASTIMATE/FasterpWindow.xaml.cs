using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.ComponentModel;
using System.Linq;
using System.Runtime.InteropServices;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Input;
using SolidWorks.Interop.sldworks;
using SolidWorks.Interop.swconst;
using Mouse = System.Windows.Input.Mouse;

namespace FASTIMATE
{
    /// <summary>
    /// FASTERP – poređenje SolidWorks strukture (custom properties) sa ERP strukturom.
    /// Lijevo stablo SolidWorks, desno gore SolidWorks info (odabrani sklop i podkomponente), dolje ERP info.
    /// </summary>
    public partial class FasterpWindow : Window
    {
        private readonly string _korisnik;
        private SldWorks _swApp;
        private FasterpNode _root;
        private FasterpNode _selected;
        private string _erpKey = "";

        private readonly ObservableCollection<FasterpRow> _solidRows = new ObservableCollection<FasterpRow>();
        private readonly ObservableCollection<FasterpRow> _erpRows = new ObservableCollection<FasterpRow>();
        private readonly ObservableCollection<FasterpRow> _odabranoRows = new ObservableCollection<FasterpRow>();

        private readonly ObservableCollection<string> _vezaOpcije = new ObservableCollection<string> { FasterpProps.VezaFantom, FasterpProps.VezaNormalna };
        private readonly ObservableCollection<string> _tipOpcije = new ObservableCollection<string>();
        private readonly ObservableCollection<string> _zaOpcije = new ObservableCollection<string>();

        private readonly List<DataGridColumn> _hiddenColumns = new List<DataGridColumn>();
        private bool _solidDirty;
        private bool _erpDirty;
        private bool _suppressSelection;

        public FasterpWindow(string korisnik = "")
        {
            InitializeComponent();
            _korisnik = korisnik ?? "";
            lblKorisnik.Text = string.IsNullOrWhiteSpace(_korisnik) ? "" : "Korisnik: " + _korisnik;

            BuildColumns(gridSolid, true);
            BuildColumns(gridErp, false);
            BuildColumns(gridOdabrano, true);
            gridOdabrano.ItemsSource = _odabranoRows;
            gridSolid.ItemsSource = _solidRows;
            SourceInitialized += (s, e) => FitToScreen();
            gridErp.ItemsSource = _erpRows;

            // Prvo se iscrta prozor, pa se tek onda (uz prozorčić "Učitavanje u toku...") čita iz SolidWorksa
            ContentRendered += (s, e) => Dispatcher.BeginInvoke(new Action(() => LoadFromSolid(null)),
                System.Windows.Threading.DispatcherPriority.Background);
            Closing += FasterpWindow_Closing;
        }

        /// <summary>Veličina prozora prema radnoj površini monitora (bez taskbara), uz uzimanje u obzir DPI skaliranja.</summary>
        private void FitToScreen()
        {
            Rect wa = SystemParameters.WorkArea;
            try
            {
                var handle = new System.Windows.Interop.WindowInteropHelper(Owner ?? this).Handle;
                var screen = System.Windows.Forms.Screen.FromHandle(handle);
                var src = PresentationSource.FromVisual(this);
                if (src?.CompositionTarget != null)
                {
                    var m = src.CompositionTarget.TransformFromDevice;
                    var tl = m.Transform(new Point(screen.WorkingArea.Left, screen.WorkingArea.Top));
                    var br = m.Transform(new Point(screen.WorkingArea.Right, screen.WorkingArea.Bottom));
                    wa = new Rect(tl, br);
                }
            }
            catch { }

            MinWidth = Math.Min(MinWidth, wa.Width);
            MinHeight = Math.Min(MinHeight, wa.Height);
            Width = Math.Max(MinWidth, Math.Floor(wa.Width * 0.94));
            Height = Math.Max(MinHeight, Math.Floor(wa.Height * 0.94));
            Left = wa.Left + (wa.Width - Width) / 2;
            Top = wa.Top + (wa.Height - Height) / 2;
        }

        // ------------------------------------------------------------------ dugmad prozora

        private void BtnMinimiziraj_Click(object sender, RoutedEventArgs e)
        {
            WindowState = WindowState.Minimized;
        }

        private void BtnMaksimiziraj_Click(object sender, RoutedEventArgs e)
        {
            WindowState = WindowState == WindowState.Maximized ? WindowState.Normal : WindowState.Maximized;
        }

        private void BtnZatvori_Click(object sender, RoutedEventArgs e)
        {
            Close();
        }

        private void Window_StateChanged(object sender, EventArgs e)
        {
            bool max = WindowState == WindowState.Maximized;
            // Maksimiziran prozor bez okvira izlazi preko ivica ekrana – margina to poravna
            rootBorder.Margin = max ? new Thickness(7) : new Thickness(0);
            btnMaksimiziraj.Content = max ? "❐" : "☐";
            btnMaksimiziraj.ToolTip = max ? "Vrati" : "Maksimiziraj";
        }

        // ------------------------------------------------------------------ kolone

        private void BuildColumns(DataGrid grid, bool solid)
        {
            grid.Columns.Clear();
            grid.Columns.Add(TextCol("N", "N", 40, true));
            grid.Columns.Add(TextCol("Br.", "Br", 45, solid));
            grid.Columns.Add(TextCol("Poz", "Poz", 50, false));
            grid.Columns.Add(ComboCol("Veza", "Veza", _vezaOpcije, 115, false));
            grid.Columns.Add(TextCol("Ident", "Ident", 95, false));
            grid.Columns.Add(TextCol("Oznaka", "Oznaka", 180, solid));
            grid.Columns.Add(ComboCol("TIP", "Tip", _tipOpcije, 140, true));
            var tipId = TextCol("TIP ID", "TipId", 60, true);
            grid.Columns.Add(tipId);
            var naziv = TextCol("Naziv", "Naziv", 200, false);
            naziv.Width = new DataGridLength(1, DataGridLengthUnitType.Star);
            naziv.MinWidth = 160;
            grid.Columns.Add(naziv);
            grid.Columns.Add(TextCol("Količina", "Kolicina", 85, solid));
            grid.Columns.Add(TextCol("JM", "JM", 55, false));
            grid.Columns.Add(ComboCol("ZA", "ZA", _zaOpcije, 85, true));
            var zaId = TextCol("ZA ID", "ZaId", 55, true);
            grid.Columns.Add(zaId);
            grid.Columns.Add(TextCol("Napomena", "Napomena", 170, false));

            // Nevidljive kolone (pozadinski brojevi za TIP i ZA)
            tipId.Visibility = Visibility.Collapsed;
            zaId.Visibility = Visibility.Collapsed;
            _hiddenColumns.Add(tipId);
            _hiddenColumns.Add(zaId);
        }

        private static DataGridTextColumn TextCol(string header, string prop, double width, bool readOnly)
        {
            return new DataGridTextColumn
            {
                Header = header,
                Binding = new Binding(prop) { Mode = readOnly ? BindingMode.OneWay : BindingMode.TwoWay },
                Width = width,
                IsReadOnly = readOnly,
                ElementStyle = CellTextStyle // isto poravnanje kao kolone sa padajućom listom (Veza, TIP, ZA)
            };
        }

        /// <summary>Tekst u ćeliji vertikalno centriran, sa istim razmakom u svim kolonama.</summary>
        private static readonly Style CellTextStyle = CreateCellTextStyle();

        private static Style CreateCellTextStyle()
        {
            var s = new Style(typeof(TextBlock));
            s.Setters.Add(new Setter(TextBlock.VerticalAlignmentProperty, VerticalAlignment.Center));
            s.Setters.Add(new Setter(TextBlock.MarginProperty, new Thickness(4, 0, 4, 0)));
            s.Seal();
            return s;
        }

        /// <summary>Kolona sa padajućom listom u koju se može upisati i nova vrijednost.</summary>
        private static DataGridTemplateColumn ComboCol(string header, string prop, IEnumerable<string> items, double width, bool editable)
        {
            var col = new DataGridTemplateColumn { Header = header, Width = width, SortMemberPath = prop };

            var tb = new FrameworkElementFactory(typeof(TextBlock));
            tb.SetBinding(TextBlock.TextProperty, new Binding(prop));
            tb.SetValue(TextBlock.MarginProperty, new Thickness(4, 0, 4, 0));
            tb.SetValue(TextBlock.VerticalAlignmentProperty, VerticalAlignment.Center);
            col.CellTemplate = new DataTemplate { VisualTree = tb };

            var cb = new FrameworkElementFactory(typeof(ComboBox));
            cb.SetValue(ComboBox.IsEditableProperty, editable);
            cb.SetValue(ComboBox.ItemsSourceProperty, items);
            cb.SetValue(ComboBox.ForegroundProperty, System.Windows.Media.Brushes.Black);
            cb.SetBinding(ComboBox.TextProperty, new Binding(prop) { Mode = BindingMode.TwoWay, UpdateSourceTrigger = UpdateSourceTrigger.PropertyChanged });
            cb.AddHandler(FrameworkElement.LoadedEvent, new RoutedEventHandler((s, e) =>
            {
                var c = (ComboBox)s;
                c.Focus();
                if (!c.IsEditable) c.IsDropDownOpen = true;
            }));
            col.CellEditingTemplate = new DataTemplate { VisualTree = cb };
            return col;
        }

        private void ChkSkrivene_Changed(object sender, RoutedEventArgs e)
        {
            var vis = chkSkrivene.IsChecked == true ? Visibility.Visible : Visibility.Collapsed;
            foreach (var c in _hiddenColumns) c.Visibility = vis;
        }

        private void RefreshSifarnikOpcije()
        {
            Fill(_tipOpcije, FasterpSifarnik.Nazivi(FasterpSifarnik.GrupaTip));
            Fill(_zaOpcije, FasterpSifarnik.Nazivi(FasterpSifarnik.GrupaZa));
        }

        private static void Fill(ObservableCollection<string> target, IEnumerable<string> values)
        {
            target.Clear();
            foreach (var v in values) target.Add(v);
        }

        // ------------------------------------------------------------------ SolidWorks

        private bool ConnectSolid()
        {
            try
            {
                _swApp = (SldWorks)Marshal.GetActiveObject("SldWorks.Application");
            }
            catch
            {
                _swApp = null;
            }
            if (_swApp == null)
            {
                MessageBox.Show(this, "SolidWorks aplikacija nije pokrenuta.", "FastERP", MessageBoxButton.OK, MessageBoxImage.Warning);
                return false;
            }
            if (_swApp.ActiveDoc == null)
            {
                MessageBox.Show(this, "Nema otvorenog dokumenta u SolidWorks-u.", "FastERP", MessageBoxButton.OK, MessageBoxImage.Warning);
                return false;
            }
            return true;
        }

        /// <summary>Učitava (ili ponovo učitava) stablo; selectKey = ključ čvora koji treba ponovo odabrati.</summary>
        private void LoadFromSolid(string selectKey)
        {
            if (!ConnectSolid()) { SetStatus("SolidWorks nije dostupan."); return; }

            _solidDirty = false; // pozivalac je već potvrdio odbacivanje izmjena
            _erpDirty = false;
            _selected = null;

            Mouse.OverrideCursor = Cursors.Wait;
            var ucitavanje = UcitavanjeProzor.Prikazi(this, "Čitanje strukture iz SolidWorksa...");
            try
            {
                _root = FasterpReader.BuildTree(_swApp);
                RefreshSifarnikOpcije();

                _suppressSelection = true;
                treeSolid.ItemsSource = _root == null ? null : new[] { _root };
                _suppressSelection = false;

                if (_root == null) { SetStatus("Nema aktivnog dokumenta."); return; }


                FasterpNode toSelect = null;
                if (selectKey != null)
                    toSelect = _root.SelfAndDescendants().FirstOrDefault(n => n.Key == selectKey && n.HasIdent);
                // inače prva komponenta (redom u stablu) koja ima Ident
                if (toSelect == null)
                    toSelect = _root.SelfAndDescendants().FirstOrDefault(n => n.HasIdent);
                if (toSelect != null)
                    SelectNode(toSelect);
                else
                    ClearView();

                int count = _root.SelfAndDescendants().Count() - 1;
                SetStatus("Učitano iz SolidWorksa: " + count + " komponenti (" + DateTime.Now.ToString("HH:mm:ss") + ").");
            }
            catch (Exception ex)
            {
                ucitavanje.Dispose();
                Mouse.OverrideCursor = null;
                MessageBox.Show(this, "Greška pri čitanju SolidWorks strukture:\n" + ex.Message, "FastERP", MessageBoxButton.OK, MessageBoxImage.Error);
            }
            finally
            {
                ucitavanje.Dispose();
                Mouse.OverrideCursor = null;
            }
        }

        private void SelectNode(FasterpNode node)
        {
            if (node == null) return;
            if (_selected != null && _selected != node) _selected.IsSelected = false;
            for (var p = node.Parent; p != null; p = p.Parent) p.IsExpanded = true;
            ShowNode(node);
            node.IsSelected = true; // samo vizuelno – podaci su već prikazani
        }

        private void BtnOsvjezi_Click(object sender, RoutedEventArgs e)
        {
            if (!ConfirmDiscard()) return;
            LoadFromSolid(_selected?.Key);
        }

        private void TreeItem_Selected(object sender, RoutedEventArgs e)
        {
            if (e.OriginalSource is TreeViewItem item) item.BringIntoView();
        }

        /// <summary>Klik na sivu komponentu (bez Identa) ne selektuje je; strelica za otvaranje i dalje radi.</summary>
        private void TreeItem_PreviewMouseLeftButtonDown(object sender, MouseButtonEventArgs e)
        {
            var source = e.OriginalSource as DependencyObject;
            if (FindAncestor<TreeViewItem>(source) != sender) return; // događaj stiže i do nadređenih stavki
            if (FindAncestor<System.Windows.Controls.Primitives.ToggleButton>(source) != null) return; // strelica

            if (sender is TreeViewItem item && item.DataContext is FasterpNode node && !node.HasIdent)
                e.Handled = true;
        }

        private static T FindAncestor<T>(DependencyObject d) where T : DependencyObject
        {
            while (d != null && !(d is T))
                d = d is System.Windows.Media.Visual || d is System.Windows.Media.Media3D.Visual3D
                    ? System.Windows.Media.VisualTreeHelper.GetParent(d)
                    : LogicalTreeHelper.GetParent(d);
            return d as T;
        }

        private void ClearView()
        {
            _selected = null;
            foreach (var r in _solidRows) r.PropertyChanged -= SolidRow_Changed;
            foreach (var r in _odabranoRows) r.PropertyChanged -= SolidRow_Changed;
            foreach (var r in _erpRows) r.PropertyChanged -= ErpRow_Changed;
            _solidRows.Clear();
            _odabranoRows.Clear();
            _erpRows.Clear();
            _erpKey = "";
            _solidDirty = false;
            _erpDirty = false;
        }

        // ------------------------------------------------------------------ tabovi

        private void Tab_Checked(object sender, RoutedEventArgs e)
        {
            if (panelStruktura == null || panelTehnologija == null) return; // tokom InitializeComponent
            bool struktura = tabStruktura.IsChecked == true;
            panelStruktura.Visibility = struktura ? Visibility.Visible : Visibility.Collapsed;
            panelTehnologija.Visibility = struktura ? Visibility.Collapsed : Visibility.Visible;
        }

        private void TreeSolid_SelectedItemChanged(object sender, RoutedPropertyChangedEventArgs<object> e)
        {
            if (_suppressSelection) return;
            var node = e.NewValue as FasterpNode;
            if (node == null || node == _selected) return;

            if (!node.HasIdent || !ConfirmDiscard())
            {
                // vrati prethodni odabir
                _suppressSelection = true;
                node.IsSelected = false;
                if (_selected != null) _selected.IsSelected = true;
                _suppressSelection = false;
                return;
            }
            ShowNode(node);
        }

        /// <summary>Puni gornju tabelu (odabrani sklop i njegove podkomponente) i donju (ERP) za odabrani ident.</summary>
        private void ShowNode(FasterpNode node)
        {
            _selected = node;
            foreach (var r in _solidRows) r.PropertyChanged -= SolidRow_Changed;
            foreach (var r in _odabranoRows) r.PropertyChanged -= SolidRow_Changed;
            _solidRows.Clear();
            _odabranoRows.Clear();

            // Red iznad: sam selektovani sklop/dio
            var self = SolidRow(node.Data);
            self.Kolicina = node.Quantity;
            self.PropertyChanged += SolidRow_Changed;
            _odabranoRows.Add(self);

            // SolidWorks info: ono što je ispod njega u stablu (jedan red po SW-File Name, sa količinom)
            int br = 1;
            foreach (var child in node.Children)
            {
                var row = SolidRow(child.Data);
                row.Br = br++;
                row.Kolicina = child.Quantity; // broj istih fajlova ispod odabranog sklopa
                row.PropertyChanged += SolidRow_Changed;
                _solidRows.Add(row);
            }


            _erpKey = ErpKey(node.Data);
            LoadErp();

            _solidDirty = false;
            _erpDirty = false;
            Uporedi();
        }


        private static string IdentSuffix(FasterpRow d)
        {
            return d == null || string.IsNullOrWhiteSpace(d.Ident) ? "" : "  [" + d.Ident + "]";
        }

        private static string ErpKey(FasterpRow d)
        {
            if (d == null) return "";
            return !string.IsNullOrWhiteSpace(d.Ident) ? d.Ident.Trim() : (d.Oznaka ?? "").Trim();
        }

        private static FasterpRow SolidRow(FasterpRow source)
        {
            var r = source.CloneData();
            r.Model = source.Model;
            r.Configuration = source.Configuration;
            r.IsAssembly = source.IsAssembly;
            return r;
        }

        private void LoadErp()
        {
            foreach (var r in _erpRows) r.PropertyChanged -= ErpRow_Changed;
            _erpRows.Clear();
            foreach (var r in FasterpErpStore.Load(_erpKey))
            {
                r.PropertyChanged += ErpRow_Changed;
                _erpRows.Add(r);
            }
        }

        private void SolidRow_Changed(object sender, PropertyChangedEventArgs e)
        {
            if (e.PropertyName != nameof(FasterpRow.Status) && e.PropertyName != nameof(FasterpRow.N)) _solidDirty = true;
        }

        private void ErpRow_Changed(object sender, PropertyChangedEventArgs e)
        {
            if (e.PropertyName != nameof(FasterpRow.Status) && e.PropertyName != nameof(FasterpRow.N)) _erpDirty = true;
        }

        // ------------------------------------------------------------------ pretraga

        private void TxtPretraga_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.Key == Key.Enter) { Trazi(); e.Handled = true; }
        }

        private void BtnTrazi_Click(object sender, RoutedEventArgs e)
        {
            Trazi();
        }

        /// <summary>Traži sljedeći čvor (poslije trenutno odabranog) po identu, oznaci ili nazivu.</summary>
        private void Trazi()
        {
            string text = txtPretraga.Text.Trim();
            if (_root == null || text.Length == 0) return;

            var all = _root.SelfAndDescendants().ToList();
            int start = _selected == null ? -1 : all.IndexOf(_selected);
            FasterpNode found = null;
            for (int i = 1; i <= all.Count; i++)
            {
                var n = all[(start + i + all.Count) % all.Count];
                if (n.Matches(text)) { found = n; break; }
            }

            if (found == null)
            {
                SetStatus("Pretraga: nije pronađeno \"" + text + "\".");
                return;
            }
            SelectNode(found);
            SetStatus("Pretraga: pronađeno " + found.DisplayName + IdentSuffix(found.Data) + ".");
        }

        // ------------------------------------------------------------------ poređenje i kopiranje

        private void BtnUporedi_Click(object sender, RoutedEventArgs e)
        {
            CommitEdits();
            Uporedi();
            int novi = _solidRows.Count(r => r.Status == FasterpStatus.Novi);
            int prom = _solidRows.Count(r => r.Status == FasterpStatus.Promijenjen);
            int izbr = _erpRows.Count(r => r.Status == FasterpStatus.Izbrisan);
            SetStatus(novi + prom + izbr == 0
                ? "Uporedi: SolidWorks i ERP su isti."
                : "Uporedi: novih " + novi + ", promijenjenih " + prom + ", izbrisanih " + izbr + ".");
        }

        private void Uporedi()
        {
            var erpByKey = new Dictionary<string, FasterpRow>();
            foreach (var r in _erpRows) if (!erpByKey.ContainsKey(r.Key)) erpByKey[r.Key] = r;
            var solidByKey = new Dictionary<string, FasterpRow>();
            foreach (var r in _solidRows) if (!solidByKey.ContainsKey(r.Key)) solidByKey[r.Key] = r;

            foreach (var s in _solidRows)
            {
                if (!erpByKey.TryGetValue(s.Key, out var e)) s.Status = FasterpStatus.Novi;
                else s.Status = s.DiffersFrom(e) ? FasterpStatus.Promijenjen : FasterpStatus.NemaPromjene;
            }
            foreach (var e in _erpRows)
            {
                if (!solidByKey.TryGetValue(e.Key, out var s)) e.Status = FasterpStatus.Izbrisan;
                else e.Status = e.DiffersFrom(s) ? FasterpStatus.Promijenjen : FasterpStatus.NemaPromjene;
            }
        }

        /// <summary>Gornja tabela -> donja (ERP dobija strukturu iz SolidWorksa). Snima se tek na "Izmjena podataka (ERP)".</summary>
        private void BtnKopirajUErp_Click(object sender, RoutedEventArgs e)
        {
            CommitEdits();
            if (_selected == null) return;
            if (string.IsNullOrWhiteSpace(_erpKey))
            {
                MessageBox.Show(this, "Odabrani sklop nema IDENT – nije moguće vezati ERP strukturu.", "FastERP", MessageBoxButton.OK, MessageBoxImage.Warning);
                return;
            }

            var oldByKey = new Dictionary<string, FasterpRow>();
            foreach (var r in _erpRows) if (!oldByKey.ContainsKey(r.Key)) oldByKey[r.Key] = r;

            foreach (var r in _erpRows) r.PropertyChanged -= ErpRow_Changed;
            _erpRows.Clear();
            foreach (var s in _solidRows)
            {
                var r = s.CloneData();
                // napomena iz ERP-a se zadržava ako u SolidWorksu nije upisana
                if (string.IsNullOrWhiteSpace(r.Napomena) && oldByKey.TryGetValue(r.Key, out var old)) r.Napomena = old.Napomena;
                r.PropertyChanged += ErpRow_Changed;
                _erpRows.Add(r);
            }
            _erpDirty = true;
            Uporedi();
            SetStatus("Kopirano iz SolidWorks u ERP. Kliknite \"Izmjena podataka (ERP)\" da sačuvate.");
        }

        /// <summary>Donja tabela -> gornja: podaci iz ERP-a se prepisuju na iste idente u SolidWorks tabeli.</summary>
        private void BtnKopirajUSolid_Click(object sender, RoutedEventArgs e)
        {
            CommitEdits();
            var erpByKey = new Dictionary<string, FasterpRow>();
            foreach (var r in _erpRows) if (!erpByKey.ContainsKey(r.Key)) erpByKey[r.Key] = r;

            int n = 0;
            foreach (var s in _solidRows)
            {
                if (!erpByKey.TryGetValue(s.Key, out var src)) continue;
                s.CopyDataFrom(src, false);
                n++;
            }
            if (n > 0) _solidDirty = true;
            Uporedi();
            int bez = _erpRows.Count(r => r.Status == FasterpStatus.Izbrisan);
            SetStatus("Kopirano iz ERP u SolidWorks: " + n + " stavki." +
                      (bez > 0 ? " " + bez + " ERP stavki ne postoji u SolidWorks sklopu (komponente se ne dodaju automatski)." : "") +
                      (n > 0 ? " Kliknite \"Izmjena podataka (SolidWorks)\" da upišete u SolidWorks." : ""));
        }

        // ------------------------------------------------------------------ snimanje

        /// <summary>Upisuje izmjene iz gornje tabele u custom properties SolidWorks modela.</summary>
        private void BtnIzmjenaSolid_Click(object sender, RoutedEventArgs e)
        {
            CommitEdits();
            if (_solidRows.Count == 0) return;
            if (!ConnectSolid()) return;

            var rows = _odabranoRows.Concat(_solidRows).Where(r => r.Model != null).ToList();
            if (rows.Count == 0)
            {
                MessageBox.Show(this, "Nijedna komponenta nije učitana u SolidWorksu (lightweight) – nema gdje upisati.", "FastERP", MessageBoxButton.OK, MessageBoxImage.Warning);
                return;
            }

            if (MessageBox.Show(this, "Upisati podatke iz tabele \"SolidWorks info\" u custom properties za " + rows.Count + " komponenti?",
                    "FastERP – Izmjena podataka", MessageBoxButton.YesNo, MessageBoxImage.Question) != MessageBoxResult.Yes)
                return;

            Mouse.OverrideCursor = Cursors.Wait;
            var errors = new List<string>();
            var models = new List<ModelDoc2>();
            try
            {
                foreach (var r in rows)
                {
                    try
                    {
                        r.EnsureIds();
                        FasterpReader.WriteRow(r);
                        if (!models.Contains(r.Model)) models.Add(r.Model);
                    }
                    catch (Exception ex) { errors.Add(r.Oznaka + ": " + ex.Message); }
                }
            }
            finally { Mouse.OverrideCursor = null; }

            if (models.Count > 0 && MessageBox.Show(this,
                    "Podaci su upisani u SolidWorks.\nSačuvati izmijenjene dokumente (" + models.Count + ")?",
                    "FastERP", MessageBoxButton.YesNo, MessageBoxImage.Question) == MessageBoxResult.Yes)
            {
                Mouse.OverrideCursor = Cursors.Wait;
                try
                {
                    foreach (var m in models)
                    {
                        int err = 0, warn = 0;
                        try
                        {
                            if (!m.Save3((int)swSaveAsOptions_e.swSaveAsOptions_Silent, ref err, ref warn))
                                errors.Add((m.GetTitle() ?? "") + ": snimanje nije uspjelo (kod " + err + ")");
                        }
                        catch (Exception ex) { errors.Add((m.GetTitle() ?? "") + ": " + ex.Message); }
                    }
                }
                finally { Mouse.OverrideCursor = null; }
            }

            if (errors.Count > 0)
                MessageBox.Show(this, "Neke stavke nisu upisane:\n" + string.Join("\n", errors.Take(15)), "FastERP", MessageBoxButton.OK, MessageBoxImage.Warning);

            _solidDirty = false;
            LoadFromSolid(_selected?.Key); // osvježi i ponovo selektuj isto
            SetStatus("Izmjena podataka (SolidWorks): upisano " + (rows.Count - errors.Count) + " stavki.");
        }

        /// <summary>Snima donju tabelu kao ERP strukturu odabranog identa.</summary>
        private void BtnIzmjenaErp_Click(object sender, RoutedEventArgs e)
        {
            CommitEdits();
            if (string.IsNullOrWhiteSpace(_erpKey))
            {
                MessageBox.Show(this, "Odabrani sklop nema IDENT – nije moguće snimiti ERP strukturu.", "FastERP", MessageBoxButton.OK, MessageBoxImage.Warning);
                return;
            }
            try
            {
                int br = 1;
                foreach (var r in _erpRows)
                {
                    r.EnsureIds();
                    if (r.Br <= 0) r.Br = br;
                    br++;
                }
                FasterpErpStore.Save(_erpKey, _erpRows);
                RefreshSifarnikOpcije();
                _erpDirty = false;
                LoadErp();
                _erpDirty = false;
                Uporedi();
                SetStatus("Izmjena podataka (ERP): sačuvano " + _erpRows.Count + " stavki za " + _erpKey + ".");
            }
            catch (Exception ex)
            {
                MessageBox.Show(this, "Greška pri snimanju ERP podataka:\n" + ex.Message, "FastERP", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        // ------------------------------------------------------------------ pomoćno

        private void CommitEdits()
        {
            gridOdabrano.CommitEdit(DataGridEditingUnit.Row, true);
            gridSolid.CommitEdit(DataGridEditingUnit.Row, true);
            gridErp.CommitEdit(DataGridEditingUnit.Row, true);
        }

        private bool ConfirmDiscard()
        {
            CommitEdits();
            if (!_solidDirty && !_erpDirty) return true;
            return MessageBox.Show(this, "Imate izmjene koje nisu sačuvane. Nastaviti bez snimanja?",
                "FastERP", MessageBoxButton.YesNo, MessageBoxImage.Warning) == MessageBoxResult.Yes;
        }

        private void FasterpWindow_Closing(object sender, CancelEventArgs e)
        {
            if (!ConfirmDiscard()) e.Cancel = true;
        }

        private void SetStatus(string text)
        {
            lblStatus.Text = text;
        }
    }
}
