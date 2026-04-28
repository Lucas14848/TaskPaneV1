using System;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
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
        // DECLARAÇÃO DOS CAMPOS (O que estava faltando e gerando os erros CS0103)
        private ISldWorks _swApp;
        private SldWorks _swAppEvents;      // Suporte para eventos do Aplicativo
        private ModelDoc2 _swDoc;           // Suporte para eventos do Documento
        private string _currentTemplate = ""; // Armazena o template do documento atual

        private bool _isDarkTheme = true;
        private readonly string _macroPath = @"Y:\01 PRJ - PROJETO\Projetos\Solid Defaults\Macro SW";
        private const string REG_KEY = @"Software\Potenza\SolidWorksAddin";
        private bool _isLoadingProps = false;
        private string _currentDocKey = null;
        private ModelDoc2 _lastModelRef = null;
        private IntPtr _hwndSource = IntPtr.Zero;
        private HwndSource _source;
        private HwndSourceHook _hook;

        public ISldWorks SwApp
        {
            get => _swApp;
            set
            {
                if (_swApp != null) DetachAppEvents();
                _swApp = value;
                if (_swApp != null) AttachAppEvents();
            }
        }

        public ExportControl()
        {
            InitializeComponent();
            CarregarConfiguracoes();
            this.Loaded += (s, e) => { AtualizarInterface(); CacheHwnd(); };
        }

        // --- LÓGICA DE EVENTOS DO APLICATIVO ---
        private void AttachAppEvents()
        {
            if (_swApp == null) return;
            _swAppEvents = (SldWorks)_swApp;
            _swAppEvents.ActiveDocChangeNotify += OnDocChange;
            _swAppEvents.ActiveModelDocChangeNotify += OnDocChange;
            _swAppEvents.FileOpenNotify2 += OnFileOpenNotify2;
            _swAppEvents.FileNewNotify2 += OnFileNewNotify2;

            AttachDocumentEvents();
        }

        private void DetachAppEvents()
        {
            if (_swAppEvents == null) return;
            _swAppEvents.ActiveDocChangeNotify -= OnDocChange;
            _swAppEvents.ActiveModelDocChangeNotify -= OnDocChange;
            _swAppEvents.FileOpenNotify2 -= OnFileOpenNotify2;
            _swAppEvents.FileNewNotify2 -= OnFileNewNotify2;
            _swAppEvents = null;
        }

        // --- LÓGICA DE EVENTOS DO DOCUMENTO (O QUE RESOLVE O FECHAMENTO) ---
        private void AttachDocumentEvents()
        {
            if (_swApp == null) return;
            ModelDoc2 doc = (ModelDoc2)_swApp.ActiveDoc;
            if (doc == null) return;

            if (_swDoc != doc)
            {
                DetachDocumentEvents();
                _swDoc = doc;

                if (_swDoc is PartDoc part) { part.DestroyNotify += OnDocDestroy; part.NewSelectionNotify += OnSelectionChanged; part.ClearSelectionsNotify += OnSelectionChanged; }
                else if (_swDoc is AssemblyDoc assy) { assy.DestroyNotify += OnDocDestroy; assy.NewSelectionNotify += OnSelectionChanged; assy.ClearSelectionsNotify += OnSelectionChanged; }
                else if (_swDoc is DrawingDoc draw) { draw.DestroyNotify += OnDocDestroy; draw.NewSelectionNotify += OnSelectionChanged; draw.ClearSelectionsNotify += OnSelectionChanged; }
            }
        }

        private void DetachDocumentEvents()
        {
            if (_swDoc == null) return;
            if (_swDoc is PartDoc part) { part.DestroyNotify -= OnDocDestroy; part.NewSelectionNotify -= OnSelectionChanged; part.ClearSelectionsNotify -= OnSelectionChanged; }
            else if (_swDoc is AssemblyDoc assy) { assy.DestroyNotify -= OnDocDestroy; assy.NewSelectionNotify -= OnSelectionChanged; assy.ClearSelectionsNotify -= OnSelectionChanged; }
            else if (_swDoc is DrawingDoc draw) { draw.DestroyNotify -= OnDocDestroy; draw.NewSelectionNotify -= OnSelectionChanged; draw.ClearSelectionsNotify -= OnSelectionChanged; }
            _swDoc = null;
        }

        private int OnDocDestroy()
        {
            Dispatcher.BeginInvoke(new Action(() =>
            {
                if (_swApp == null || _swApp.ActiveDoc == null)
                {
                    ForcarLimpezaUI();
                }
                else
                {
                    AttachDocumentEvents();
                    AtualizarInterface();
                }
            }));
            return 0;
        }

        private int OnSelectionChanged()
        {
            Dispatcher.BeginInvoke(new Action(() => { AtualizarInterface(); }));
            return 0;
        }

        private int OnDocChange() { AttachDocumentEvents(); AtualizarInterface(); return 0; }
        private int OnFileOpenNotify2(string fileName) { AttachDocumentEvents(); AtualizarInterface(); return 0; }
        private int OnFileNewNotify2(object newDoc, int docType, string templateName) { AttachDocumentEvents(); AtualizarInterface(); return 0; }

        // --- INTERFACE E UI ---
        private void ForcarLimpezaUI()
        {
            _currentDocKey = null;
            _lastModelRef = null;
            _currentTemplate = "";

            LimparCampos();

            PropriedadesBorder.Visibility = Visibility.Collapsed;
            PanelMacrosPeca.Visibility = Visibility.Collapsed;
            PanelMacrosMontagem.Visibility = Visibility.Collapsed;
            PanelMacrosDesenho.Visibility = Visibility.Collapsed;
        }

        public void AtualizarInterface()
        {
            if (SwApp == null) return;
            ModelDoc2 swModel = (ModelDoc2)SwApp.ActiveDoc;

            if (swModel == null)
            {
                ForcarLimpezaUI();
                return;
            }

            ModelDoc2 targetModel = swModel;

            // Lógica de seleção
            // Importante: não chamar APIs de Assembly (ex.: GetSelectedObjectsComponent4) quando o doc ativo é Drawing,
            // pois ao clicar em vistas no ambiente de desenho isso pode causar crash do SolidWorks.
            int activeDocType;
            try { activeDocType = swModel.GetType(); }
            catch { return; }

            ISelectionMgr selMgr = null;
            try { selMgr = (ISelectionMgr)swModel.SelectionManager; }
            catch { selMgr = null; }

            if (selMgr != null && selMgr.GetSelectedObjectCount2(-1) > 0)
            {
                if (activeDocType == (int)swDocumentTypes_e.swDocASSEMBLY)
                {
                    try
                    {
                        Component2 comp = (Component2)selMgr.GetSelectedObjectsComponent4(1, -1);
                        if (comp == null)
                        {
                            object selObj = selMgr.GetSelectedObject6(1, -1);
                            if (selObj is Component2 c) comp = c;
                        }

                        if (comp != null)
                        {
                            ModelDoc2 compModel = (ModelDoc2)comp.GetModelDoc2();
                            if (compModel != null) targetModel = compModel;
                        }
                    }
                    catch (COMException)
                    {
                        System.Diagnostics.Debug.WriteLine("Seleção (Assembly) falhou ao resolver componente selecionado.");
                    }
                }
                // Em Drawing: manter o documento ativo como Drawing.
                // Mesmo se a seleÃ§Ã£o for uma vista/componente, nÃ£o trocar a UI para PeÃ§a/Montagem.
            }

            int type = targetModel.GetType();
            bool isModel = type == (int)swDocumentTypes_e.swDocASSEMBLY || type == (int)swDocumentTypes_e.swDocPART;
            bool isSheetMetal = type == (int)swDocumentTypes_e.swDocPART && IsSheetMetal(targetModel);

            _currentDocKey = GetDocKey(targetModel);
            _lastModelRef = targetModel;
            _currentTemplate = targetModel.Extension.get_CustomPropertyBuilderTemplate(false);

            PanelMacrosPeca.Visibility = Visibility.Collapsed;
            PanelMacrosMontagem.Visibility = Visibility.Collapsed;
            PanelMacrosDesenho.Visibility = Visibility.Collapsed;

            if (isModel)
            {
                PropriedadesBorder.Visibility = Visibility.Visible;
                CarregarPropriedades(targetModel, isSheetMetal);

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
                PropriedadesBorder.Visibility = Visibility.Collapsed;
                PanelMacrosDesenho.Visibility = Visibility.Visible;
                LblTituloMacros.Text = "MACROS - DESENHO";
            }
        }

        private ModelDoc2 TryGetSelectedModelFromDrawing(ISelectionMgr selMgr)
        {
            if (selMgr == null) return null;

            try
            {
                int count = selMgr.GetSelectedObjectCount2(-1);
                for (int i = 1; i <= count; i++)
                {
                    object selObj;
                    try { selObj = selMgr.GetSelectedObject6(i, -1); }
                    catch (COMException) { continue; }

                    if (selObj is Component2 comp)
                    {
                        ModelDoc2 compModel = (ModelDoc2)comp.GetModelDoc2();
                        if (compModel != null) return compModel;
                    }

                    if (selObj is DrawingComponent drawComp)
                    {
                        Component2 drawingComponent = null;
                        try { drawingComponent = (Component2)drawComp.Component; }
                        catch (COMException) { drawingComponent = null; }

                        if (drawingComponent != null)
                        {
                            ModelDoc2 drawingModel = (ModelDoc2)drawingComponent.GetModelDoc2();
                            if (drawingModel != null) return drawingModel;
                        }
                    }

                    if (selObj is View view)
                    {
                        object refDoc;
                        try { refDoc = view.ReferencedDocument; }
                        catch (COMException) { continue; }

                        if (refDoc is ModelDoc2 refModel) return refModel;
                    }
                }
            }
            catch (COMException)
            {
                System.Diagnostics.Debug.WriteLine("Seleção (Drawing) falhou ao resolver documento referenciado.");
            }

            return null;
        }

        private void CarregarPropriedades(ModelDoc2 swModel, bool isSheetMetal)
        {
            _isLoadingProps = true;
            try
            {
                string template = !string.IsNullOrEmpty(_currentTemplate) ? Path.GetFileName(_currentTemplate).ToLower() : "";

                TxtNomProjeto.Text = GetProp(swModel, "Nom.Projeto");
                TxtCodProjeto.Text = GetProp(swModel, "Conjunto");
                TxtDenominacao.Text = GetProp(swModel, "Denominação");
                TxtRevisao.Text = GetProp(swModel, "Rev.");
                TxtCodigo.Text = ObterValorPorTemplate(swModel, "Código", template);
                TxtNumeroDesenho.Text = ObterValorPorTemplate(swModel, "Número Desenho", template);
                TxtProjetista.Text = GetProp(swModel, "Projetista");
                TxtDataProjeto.Text = GetProp(swModel, "Data do Projeto");
                TxtAprov.Text = GetProp(swModel, "Aprov.");
                TxtDatAprov.Text = GetProp(swModel, "Dat.Aprov.");
                TxtMaterial.Text = GetProp(swModel, "Material");
                TxtObs.Text = GetProp(swModel, "Obs.");

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
            }
            finally { _isLoadingProps = false; }
        }

        private void TextBox_TextChanged(object sender, TextChangedEventArgs e) { }
        private void TextBox_GotKeyboardFocus(object sender, KeyboardFocusChangedEventArgs e) => EnsureKeyboardFocus(sender as TextBox);
        private void TextBox_PreviewMouseDown(object sender, MouseButtonEventArgs e) => EnsureKeyboardFocus(sender as TextBox);

        private void TextBox_LostFocus(object sender, RoutedEventArgs e)
        {
            if (_isLoadingProps) return;
            TextBox tb = sender as TextBox;
            if (tb == null || SwApp == null) return;
            ModelDoc2 targetModel = _lastModelRef;
            if (targetModel == null || tb.Tag == null) return;

            string template = !string.IsNullOrEmpty(_currentTemplate) ? Path.GetFileName(_currentTemplate).ToLower() : "";
            SetProp(targetModel, tb.Tag.ToString(), tb.Text, template);
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

        // --- HANDLERS DE BOTÕES E MACROS ---
        private void BtnMacro01_Click(object sender, RoutedEventArgs e) => ExecutarMacro("01-Cria uma folha de desenho com 3 vistas.swp");
        private void BtnMacro02_Click(object sender, RoutedEventArgs e) => ExecutarMacro("02-Renumerar Folhas de Desenho.swp");
        private void BtnMacro03_Click(object sender, RoutedEventArgs e) => ExecutarMacro("03-Renomear, Criar Cópias e Criar Revisões em Peças e Montagens.swp");
        private void BtnMacro04_Click(object sender, RoutedEventArgs e) => ExecutarMacro("04-Salva PDFs de uma Montagem-ERP.swp");
        private void BtnMacro05_Click(object sender, RoutedEventArgs e) => ExecutarMacro("05-Impressao Automatica de Desenhos.swp");
        private void BtnMacro06_Click(object sender, RoutedEventArgs e) => ExecutarMacro("06-Solda por Varredura.swp");
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
        private void BtnMacro26_Click(object sender, RoutedEventArgs e)
        {
            if (SwApp == null) return;
            ExecutarMacro("26-Pacote de preenchimento propriedades personalizadas na Peça.swp");
            AtualizarInterface();
        }
        private void BtnMacro40_Click(object sender, RoutedEventArgs e) => ExecutarMacro("40-Pós importador de Engenharia.swp");
        private void BtnMacro41_Click(object sender, RoutedEventArgs e) => ExecutarMacro("41-Criar Revisão WORD.swp");
        private void BtnMacro50_Click(object sender, RoutedEventArgs e) => ExecutarMacro("50 - Cria Catalogo de peças.swp");
        private void BtnMacro55_Click(object sender, RoutedEventArgs e) => ExecutarMacro("55-Cadastra Todas as pecas Selecionadas.swp");

        private void BtnTheme_Click(object sender, RoutedEventArgs e) { _isDarkTheme = !_isDarkTheme; AplicarTema(); SalvarConfiguracoes(); }
        private void BtnAtualizarProps_Click(object sender, RoutedEventArgs e) => AtualizarInterface();
        private void BtnSalvarProps_Click(object sender, RoutedEventArgs e)
        {
            if (_lastModelRef != null) { SalvarCamposDigitados(_lastModelRef); }
        }

        private void SalvarCamposDigitados(ModelDoc2 swModel)
        {
            TextBox[] campos = { TxtNomProjeto, TxtCodProjeto, TxtDenominacao, TxtNumeroDesenho, TxtCodigo, TxtRevisao, TxtProjetista, TxtDataProjeto, TxtAprov, TxtDatAprov, TxtMaterial, TxtEspessura, TxtCatalogo, TxtObs };
            string template = !string.IsNullOrEmpty(_currentTemplate) ? Path.GetFileName(_currentTemplate).ToLower() : "";
            foreach (var tb in campos) if (tb?.Tag != null && !tb.IsReadOnly) SetProp(swModel, tb.Tag.ToString(), tb.Text, template);
        }

        private string GetDocKey(ModelDoc2 model)
        {
            if (model == null) return null;
            string path = model.GetPathName();
            return !string.IsNullOrWhiteSpace(path) ? path.ToLowerInvariant() : $"{model.GetTitle()}_{model.GetType()}".ToLowerInvariant();
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
            CustomPropertyManager mgr = isConfig ? ((Configuration)swModel.GetActiveConfiguration()).CustomPropertyManager : swModel.Extension.get_CustomPropertyManager("");
            mgr.Add3(propName, (int)swCustomInfoType_e.swCustomInfoText, value, (int)swCustomPropertyAddOption_e.swCustomPropertyReplaceValue);
        }

        // --- MANIPULAÇÃO DE HWND E FOCO ---
        private void CacheHwnd()
        {
            _source = (HwndSource)PresentationSource.FromVisual(this);
            _hwndSource = _source?.Handle ?? IntPtr.Zero;
            if (_source != null && _hook == null) { _hook = new HwndSourceHook(WndProc); _source.AddHook(_hook); }
        }

        private void EnsureKeyboardFocus(TextBox tb)
        {
            if (tb == null) return;
            if (_hwndSource == IntPtr.Zero) CacheHwnd();
            if (_hwndSource != IntPtr.Zero) NativeMethods.SetFocus(_hwndSource);
            if (!tb.IsKeyboardFocused) { tb.Focus(); Keyboard.Focus(tb); }
        }

        private IntPtr WndProc(IntPtr hwnd, int msg, IntPtr wParam, IntPtr lParam, ref bool handled)
        {
            if (msg == 0x0087) { handled = true; return new IntPtr(0x0004 | 0x0001 | 0x0002); }
            return IntPtr.Zero;
        }

        private bool IsSheetMetal(ModelDoc2 model)
        {
            if (model == null || model.GetType() != (int)swDocumentTypes_e.swDocPART) return false;
            Feature feat = (Feature)model.FirstFeature();
            while (feat != null)
            {
                string typeName = feat.GetTypeName2();
                if (typeName == "SheetMetal" || typeName == "SheetMetalFolder") return true;
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

        private void SalvarConfiguracoes() { try { RegistryKey key = Registry.CurrentUser.CreateSubKey(REG_KEY); key.SetValue("DarkTheme", _isDarkTheme ? 1 : 0); key.Close(); } catch { } }
        private void CarregarConfiguracoes() { try { RegistryKey key = Registry.CurrentUser.OpenSubKey(REG_KEY); if (key != null) { _isDarkTheme = (int)key.GetValue("DarkTheme", 1) == 1; key.Close(); } } catch { _isDarkTheme = true; } AplicarTema(); }
        private void SetResource(string key, string colorHex) => this.Resources[key] = new SolidColorBrush((Color)ColorConverter.ConvertFromString(colorHex));
        private void BtnExpandir_Click(object sender, RoutedEventArgs e) { GridExpandido.Visibility = GridExpandido.Visibility == Visibility.Collapsed ? Visibility.Visible : Visibility.Collapsed; BtnExpandir.Content = GridExpandido.Visibility == Visibility.Visible ? "▲ RECOLHER" : "▼ MAIS DETALHES"; }
        private void LimparCampos() { TxtNomProjeto.Text = TxtCodProjeto.Text = TxtDenominacao.Text = TxtNumeroDesenho.Text = TxtCodigo.Text = TxtRevisao.Text = TxtProjetista.Text = TxtDataProjeto.Text = TxtPeso.Text = TxtAprov.Text = TxtDatAprov.Text = TxtObs.Text = TxtMaterial.Text = TxtEspessura.Text = TxtCatalogo.Text = "---"; }
    }

    internal static class NativeMethods { [System.Runtime.InteropServices.DllImport("user32.dll")] internal static extern IntPtr SetFocus(IntPtr hWnd); }
}
