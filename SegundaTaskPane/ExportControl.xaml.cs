using System;
using System.IO;
using System.Linq;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Media;
using System.Windows.Input;
using System.Windows.Interop;
using Microsoft.Win32;
using SolidWorks.Interop.sldworks;
using SolidWorks.Interop.swconst;

namespace SegundaTaskPane
{
    public partial class ExportControl : UserControl
    {
        private ISldWorks _swApp;
        private bool _isDarkTheme = true;
        private readonly string _macroPath = @"Y:\01 PRJ - PROJETO\Projetos\Solid Defaults\Macro SW";
        private const string REG_KEY = @"Software\Potenza\SolidWorksAddin";
        private bool _isLoadingProps = false;
        private bool _hasPendingChanges = false;
        private string _currentDocKey = null;
        private ModelDoc2 _lastModelRef = null;
        private string _currentTemplate = "";
        private IntPtr _hwndSource = IntPtr.Zero;
        private HwndSource _source;
        private HwndSourceHook _hook;

        public ISldWorks SwApp
        {
            get => _swApp;
            set
            {
                _swApp = value;
                if (_swApp != null)
                {
                    var app = (SldWorks)_swApp;
                    app.ActiveDocChangeNotify += OnDocChange;
                    app.ActiveModelDocChangeNotify += OnDocChange;
                    app.FileCloseNotify += OnFileClose;
                }
            }
        }

        public ExportControl()
        {
            InitializeComponent();
            CarregarConfiguracoes();
            this.Loaded += (s, e) => AtualizarInterface();
            this.Loaded += (s, e) => CacheHwnd();
        }

        private int OnDocChange()
        {
            Dispatcher.BeginInvoke(new Action(() =>
            {
                ModelDoc2 newModel = (ModelDoc2)SwApp?.ActiveDoc;
                string newKey = GetDocKey(newModel);

                if (_hasPendingChanges && _currentDocKey != null && newKey != _currentDocKey && _lastModelRef != null)
                {
                    /*var result = MessageBox.Show("Salvar alterações das propriedades antes de trocar de peça/desenho?",
                                                 "Salvar propriedades",
                                                 MessageBoxButton.YesNo,
                                                 MessageBoxImage.Question);
                    
                    if (result == MessageBoxResult.Yes)
                    {
                        SalvarCamposDigitados(_lastModelRef);
                    }
                    */
                    // Mesmo que escolha não salvar, não perguntar novamente para esta troca
                    _hasPendingChanges = false;
                }

            if (SwApp == null || newModel == null)
            {
                LimparCampos();
                PropriedadesBorder.Visibility = Visibility.Collapsed;
                PanelMacrosPeca.Visibility = Visibility.Collapsed;
                PanelMacrosMontagem.Visibility = Visibility.Collapsed;
                PanelMacrosDesenho.Visibility = Visibility.Collapsed;
                _currentDocKey = null;
                _lastModelRef = null;
                _currentTemplate = "";
                _hasPendingChanges = false;
                return;
            }

                // Atualiza doc de referência antes de repintar para evitar prompts duplicados em eventos consecutivos
                _currentDocKey = newKey;
                _lastModelRef = newModel;

                AtualizarInterface();
            }));
            return 0;
        }
        private int OnFileClose(string name, int reason) { Dispatcher.BeginInvoke(new Action(() => { LimparCampos(); AtualizarInterface(); })); return 0; }

        public void AtualizarInterface()
        {
            if (SwApp == null) return;
            ModelDoc2 swModel = (ModelDoc2)SwApp.ActiveDoc;

            LimparCampos();
            PropriedadesBorder.Visibility = Visibility.Collapsed;
            PanelMacrosPeca.Visibility = Visibility.Collapsed;
            PanelMacrosMontagem.Visibility = Visibility.Collapsed;
            PanelMacrosDesenho.Visibility = Visibility.Collapsed;

            if (swModel == null)
            {
                _currentDocKey = null;
                _lastModelRef = null;
                _hasPendingChanges = false;
                _currentTemplate = "";
                return;
            }

            int type = swModel.GetType();
            bool isModel = type == (int)swDocumentTypes_e.swDocASSEMBLY || type == (int)swDocumentTypes_e.swDocPART;
            bool isSheetMetal = type == (int)swDocumentTypes_e.swDocPART && IsSheetMetal(swModel);

            _currentDocKey = GetDocKey(swModel);
            _lastModelRef = swModel;
            _currentTemplate = swModel.Extension.get_CustomPropertyBuilderTemplate(false);

            if (isModel)
            {
                PropriedadesBorder.Visibility = Visibility.Visible;
                CarregarPropriedades(swModel, isSheetMetal);

                if (type == (int)swDocumentTypes_e.swDocASSEMBLY)
                {
                    PanelMacrosMontagem.Visibility = Visibility.Visible;
                    LblTituloMacros.Text = "MACROS - MONTAGEM";
                    AjustarVisibilidadeRows(false, false);
                }
                else
                {
                    PanelMacrosPeca.Visibility = Visibility.Visible;
                    LblTituloMacros.Text = "MACROS - PEÇA";
                    AjustarVisibilidadeRows(true, isSheetMetal);
                }
            }
            else if (type == (int)swDocumentTypes_e.swDocDRAWING)
            {
                PanelMacrosDesenho.Visibility = Visibility.Visible;
                LblTituloMacros.Text = "MACROS - DESENHO";
            }
        }

        private void CarregarPropriedades(ModelDoc2 swModel, bool isSheetMetal)
        {
            _isLoadingProps = true;
            try
            {
                string template = !string.IsNullOrEmpty(_currentTemplate) ? Path.GetFileName(_currentTemplate).ToLower() : "";

                // Dados Principais
                TxtNomProjeto.Text = GetProp(swModel, "Nom.Projeto");
                TxtCodProjeto.Text = GetProp(swModel, "Conjunto");
                TxtDenominacao.Text = GetProp(swModel, "Denominação");
                TxtRevisao.Text = GetProp(swModel, "Rev.");

                // Lógica de Template para Código e Número Desenho
                TxtCodigo.Text = ObterValorPorTemplate(swModel, "Código", template);
                TxtNumeroDesenho.Text = ObterValorPorTemplate(swModel, "Número Desenho", template);

                // Detalhes Expandidos
                TxtProjetista.Text = GetProp(swModel, "Projetista");
                TxtDataProjeto.Text = GetProp(swModel, "Data do Projeto");
                TxtAprov.Text = GetProp(swModel, "Aprov.");
                TxtDatAprov.Text = GetProp(swModel, "Dat.Aprov.");
                TxtMaterial.Text = GetProp(swModel, "Material");
                TxtObs.Text = GetProp(swModel, "Obs.");

                // Peso Físico (Calculado)
                MassProperty massProps = swModel.Extension.CreateMassProperty();
                TxtPeso.Text = massProps != null ? massProps.Mass.ToString("N3") : "0.000";

                if (swModel.GetType() == (int)swDocumentTypes_e.swDocPART)
                {
                    TxtEspessura.Text = isSheetMetal ? GetProp(swModel, "Espessura") : "---";
                    TxtCatalogo.Text = "---";
                }
                else
                {
                    TxtCatalogo.Text = GetProp(swModel, "Catálogo");
                    TxtEspessura.Text = "---";
                }

                _hasPendingChanges = false;
            }
            finally
            {
                _isLoadingProps = false;
            }
        }

        private void TextBox_TextChanged(object sender, TextChangedEventArgs e)
        {
            if (_isLoadingProps) return;
            _hasPendingChanges = true;
        }

        private void TextBox_GotKeyboardFocus(object sender, KeyboardFocusChangedEventArgs e)
        {
            EnsureKeyboardFocus(sender as TextBox);
        }

        private void TextBox_PreviewMouseDown(object sender, MouseButtonEventArgs e)
        {
            EnsureKeyboardFocus(sender as TextBox);
        }

        // EVENTO PARA SALVAR AUTOMATICAMENTE AO EDITAR O TEXTO
        private void TextBox_LostFocus(object sender, RoutedEventArgs e)
        {
            if (_isLoadingProps) return;
            TextBox tb = sender as TextBox;
            if (tb == null || SwApp == null) return;

            ModelDoc2 swModel = (ModelDoc2)SwApp.ActiveDoc;
            if (swModel == null || tb.Tag == null) return;

            string propName = tb.Tag.ToString();
            string newValue = tb.Text;
            string template = !string.IsNullOrEmpty(_currentTemplate) ? Path.GetFileName(_currentTemplate).ToLower() : "";
            SetProp(swModel, propName, newValue, template);
        }

        private string ObterValorPorTemplate(ModelDoc2 swModel, string prop, string template)
        {
            bool isConfig = false;
            if (prop == "Número Desenho" && template == "propriedade de peca nova.prtprp") isConfig = true;
            else if (prop == "Código")
            {
                string[] targets = { "05 - catalogo ns", "propriedade de montagem.asmprp", "propriedade de peça nova.prtprp", "propriedade de peça.prtprp" };
                if (targets.Any(t => template.Contains(t))) isConfig = true;
            }

            CustomPropertyManager mgr = isConfig ? ((Configuration)swModel.GetActiveConfiguration()).CustomPropertyManager : swModel.Extension.get_CustomPropertyManager("");
            mgr.Get4(prop, false, out _, out string res);
            return string.IsNullOrWhiteSpace(res) ? "---" : res;
        }

        private void AjustarVisibilidadeRows(bool isPart, bool isSheetMetal)
        {
            RowMaterial.Visibility = isPart ? Visibility.Visible : Visibility.Collapsed;
            RowEspessura.Visibility = (isPart && isSheetMetal) ? Visibility.Visible : Visibility.Collapsed;
            RowCatalogo.Visibility = !isPart ? Visibility.Visible : Visibility.Collapsed;
        }

        private string GetProp(ModelDoc2 model, string name)
        {
            if (model == null) return "---";

            string result = GetPropValue(model.Extension.get_CustomPropertyManager(""), name);

            if (string.IsNullOrWhiteSpace(result))
            {
                Configuration cfg = (Configuration)model.GetActiveConfiguration();
                result = GetPropValue(cfg?.CustomPropertyManager, name);
            }

            return string.IsNullOrWhiteSpace(result) ? "---" : result;
        }

        private string GetPropValue(CustomPropertyManager mgr, string name)
        {
            if (mgr == null) return null;
            mgr.Get4(name, false, out _, out string res);
            return res;
        }

        private void ExecutarMacro(string nome)
        {
            if (SwApp == null) return;
            string path = Path.Combine(_macroPath, nome);
            if (File.Exists(path)) SwApp.RunMacro2(path, "", "main", 1, out _);
            else MessageBox.Show("Macro não encontrada: " + nome);
        }

        // Handlers de Macros
        private void BtnMacro01_Click(object sender, RoutedEventArgs e) => ExecutarMacro("01-Cria uma folha de desenho com 3 vistas.swp");
        private void BtnMacro02_Click(object sender, RoutedEventArgs e) => ExecutarMacro("02-Renumerar Folhas de Desenho.swp");
        private void BtnMacro03_Click(object sender, RoutedEventArgs e) => ExecutarMacro("03-Renomear, Criar Cópias e Criar Revisões em Peças e Montagens.swp");
        private void BtnMacro04_Click(object sender, RoutedEventArgs e) => ExecutarMacro("04-Salva PDFs de uma Montagem-ERP.swp");
        private void BtnMacro05_Click(object sender, RoutedEventArgs e) => ExecutarMacro("05-Impressao Automatica de Desenhos.swp");
        private void BtnMacro06_Click(object sender, RoutedEventArgs e) => ExecutarMacro("06-Soldagem 3D.swp");
        private void BtnMacro07_Click(object sender, RoutedEventArgs e) => ExecutarMacro("07-Lista de Chapas DXF QTY.swp");
        private void BtnMacro13_Click(object sender, RoutedEventArgs e) => ExecutarMacro("13-Salva em PDF - ERP - CIGAM.swp");
        private void BtnMacro14_Click(object sender, RoutedEventArgs e) => ExecutarMacro("14-Navegador catalogos.swp");
        private void BtnMacro15_Click(object sender, RoutedEventArgs e) => ExecutarMacro("15-Cria propriedade espessura de chapa.swp");
        private void BtnMacro16_Click(object sender, RoutedEventArgs e) => ExecutarMacro("16-Cria cores chapa metalica.swp");
        private void BtnMacro17_Click(object sender, RoutedEventArgs e) => ExecutarMacro("17-Cria cavidades de alivio.swp");
        private void BtnMacro18_Click(object sender, RoutedEventArgs e) => ExecutarMacro("18-Excel Importação CIGAM.swp");
        private void BtnMacro22_Click(object sender, RoutedEventArgs e) => ExecutarMacro("22-Criar e Renomear Mangueiras.swp");
        private void BtnMacro24_Click(object sender, RoutedEventArgs e) => ExecutarMacro("24-Link das Vistas a Lista de materiais.swp");
        private void BtnMacro25_Click(object sender, RoutedEventArgs e) => ExecutarMacro("25-Verifica Revisões de todas as peças da montagem.swp");
        private void BtnMacro26_Click(object sender, RoutedEventArgs e) {
            if (SwApp == null) return;
            ModelDoc2 swModel = (ModelDoc2)SwApp.ActiveDoc;
            int type = swModel.GetType();
            bool isSheetMetal = type == (int)swDocumentTypes_e.swDocPART && IsSheetMetal(swModel);
            ExecutarMacro("26-Pacote de preenchimento propriedades personalizadas na Peça.swp");
            CarregarPropriedades(swModel, isSheetMetal);
        }

        private void BtnMacro40_Click(object sender, RoutedEventArgs e) => ExecutarMacro("40-Pós importador de Engenharia.swp");
        private void BtnMacro41_Click(object sender, RoutedEventArgs e) => ExecutarMacro("41-Criar Revisão WORD.swp");
        private void BtnMacro50_Click(object sender, RoutedEventArgs e) => ExecutarMacro("50 - Cria Catalogo de peças.swp");
        private void BtnMacro55_Click(object sender, RoutedEventArgs e) => ExecutarMacro("55-Cadastra Todas as pecas Selecionadas.swp");

        private void BtnTheme_Click(object sender, RoutedEventArgs e)
        {
            _isDarkTheme = !_isDarkTheme;
            AplicarTema();
            SalvarConfiguracoes();
        }

        private void BtnAtualizarProps_Click(object sender, RoutedEventArgs e)
        {
            if (SwApp == null) return;
            ModelDoc2 swModel = (ModelDoc2)SwApp.ActiveDoc;
            if (swModel == null) return;
            AtualizarInterface();
        }

        private void BtnSalvarProps_Click(object sender, RoutedEventArgs e)
        {
            if (SwApp == null) return;
            ModelDoc2 swModel = (ModelDoc2)SwApp.ActiveDoc;
            if (swModel == null)
            {
                LimparCampos();
                PropriedadesBorder.Visibility = Visibility.Collapsed;
                return;
            }

            SalvarCamposDigitados(swModel);
            _hasPendingChanges = false;
        }

        private void SalvarCamposDigitados(ModelDoc2 swModel)
        {
            TextBox[] campos =
            {
                TxtNomProjeto, TxtCodProjeto, TxtDenominacao, TxtNumeroDesenho, TxtCodigo, TxtRevisao,
                TxtProjetista, TxtDataProjeto, TxtAprov, TxtDatAprov, TxtMaterial, TxtEspessura, TxtCatalogo, TxtObs
            };

            foreach (var tb in campos)
            {
                if (tb?.Tag == null || tb.IsReadOnly) continue;
                string propName = tb.Tag.ToString();
                string template = !string.IsNullOrEmpty(_currentTemplate) ? Path.GetFileName(_currentTemplate).ToLower() : "";
                SetProp(swModel, propName, tb.Text, template);
            }
        }

        private string GetDocKey(ModelDoc2 model)
        {
            if (model == null) return null;
            string path = model.GetPathName();
            if (!string.IsNullOrWhiteSpace(path)) return path.ToLowerInvariant();
            return $"{model.GetTitle()}_{model.GetType()}".ToLowerInvariant();
        }

        private void SetProp(ModelDoc2 swModel, string propName, string value, string template)
        {
            bool isConfig = false;
            if (propName == "Número Desenho" && template == "propriedade de peca nova.prtprp") isConfig = true;
            else if (propName == "Código")
            {
                string[] targets = { "05 - catalogo ns", "propriedade de montagem.asmprp", "propriedade de peça nova.prtprp", "propriedade de peça.prtprp" };
                if (targets.Any(t => template.Contains(t))) isConfig = true;
            }

            CustomPropertyManager mgr = isConfig
                ? ((Configuration)swModel.GetActiveConfiguration()).CustomPropertyManager
                : swModel.Extension.get_CustomPropertyManager("");

            mgr.Add3(propName, (int)swCustomInfoType_e.swCustomInfoText, value, (int)swCustomPropertyAddOption_e.swCustomPropertyReplaceValue);
        }

        private void CacheHwnd()
        {
            _source = (HwndSource)PresentationSource.FromVisual(this);
            _hwndSource = _source?.Handle ?? IntPtr.Zero;

            if (_source != null && _hook == null)
            {
                _hook = new HwndSourceHook(WndProc);
                _source.AddHook(_hook);
            }
        }

        private void EnsureKeyboardFocus(TextBox tb)
        {
            if (tb == null) return;

            if (_hwndSource == IntPtr.Zero)
                CacheHwnd();

            if (_hwndSource != IntPtr.Zero)
            {
                NativeMethods.SetFocus(_hwndSource);
            }

            if (!tb.IsKeyboardFocused)
            {
                tb.Focus();
                Keyboard.Focus(tb);
            }
        }

        private IntPtr WndProc(IntPtr hwnd, int msg, IntPtr wParam, IntPtr lParam, ref bool handled)
        {
            const int WM_GETDLGCODE = 0x0087;

            if (msg == WM_GETDLGCODE)
            {
                const int DLGC_WANTALLCHARS = 0x0004;
                const int DLGC_WANTARROWS = 0x0001;
                const int DLGC_WANTTAB = 0x0002;
                handled = true;
                return new IntPtr(DLGC_WANTALLCHARS | DLGC_WANTARROWS | DLGC_WANTTAB);
            }

            return IntPtr.Zero;
        }

        private bool IsSheetMetal(ModelDoc2 model)
        {
            if (model == null || model.GetType() != (int)swDocumentTypes_e.swDocPART) return false;

            Feature feat = (Feature)model.FirstFeature();
            while (feat != null)
            {
                string typeName = feat.GetTypeName2();
                if (typeName == "SheetMetal" || typeName == "SheetMetalFolder")
                    return true;
                feat = (Feature)feat.GetNextFeature();
            }

            return false;
        }

        private void AplicarTema()
        {
            SetResource("BgColor", _isDarkTheme ? "#121212" : "#F5F5F5");
            SetResource("CardColor", _isDarkTheme ? "#1E1E1E" : "#FFFFFF");
            SetResource("MainColor", _isDarkTheme ? "#00B4FF" : "#005A9E");
            SetResource("BtnBg", _isDarkTheme ? "#2A2A2A" : "#E1E1E1");
            SetResource("BtnText", _isDarkTheme ? "#FFFFFF" : "#000000");
            SetResource("BorderColor", _isDarkTheme ? "#333333" : "#CCCCCC");
        }

        private void SalvarConfiguracoes()
        {
            try
            {
                RegistryKey key = Registry.CurrentUser.CreateSubKey(REG_KEY);
                key.SetValue("DarkTheme", _isDarkTheme ? 1 : 0);
                key.Close();
            }
            catch { }
        }

        private void CarregarConfiguracoes()
        {
            try
            {
                RegistryKey key = Registry.CurrentUser.OpenSubKey(REG_KEY);
                if (key != null)
                {
                    _isDarkTheme = (int)key.GetValue("DarkTheme", 1) == 1;
                    key.Close();
                }
            }
            catch { _isDarkTheme = true; }
            AplicarTema();
        }

        private void SetResource(string key, string colorHex) => this.Resources[key] = new SolidColorBrush((Color)ColorConverter.ConvertFromString(colorHex));

        private void BtnExpandir_Click(object sender, RoutedEventArgs e)
        {
            GridExpandido.Visibility = GridExpandido.Visibility == Visibility.Collapsed ? Visibility.Visible : Visibility.Collapsed;
            BtnExpandir.Content = GridExpandido.Visibility == Visibility.Visible ? "▲ RECOLHER" : "▼ MAIS DETALHES";
        }

        private void LimparCampos()
        {
            TxtNomProjeto.Text = TxtCodProjeto.Text = TxtDenominacao.Text = TxtNumeroDesenho.Text = TxtCodigo.Text = TxtRevisao.Text = "---";
            TxtProjetista.Text = TxtDataProjeto.Text = TxtPeso.Text = "---";
            TxtAprov.Text = TxtDatAprov.Text = TxtObs.Text = TxtMaterial.Text = TxtEspessura.Text = TxtCatalogo.Text = "---";
        }
    }

    internal static class NativeMethods
    {
        [System.Runtime.InteropServices.DllImport("user32.dll")]
        internal static extern IntPtr SetFocus(IntPtr hWnd);
    }
}
