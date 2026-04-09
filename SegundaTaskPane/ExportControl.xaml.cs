using System;
using System.IO;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Media;
using System.Windows.Input;
using SolidWorks.Interop.sldworks;
using SolidWorks.Interop.swconst;

namespace SegundaTaskPane
{
    public partial class ExportControl : UserControl
    {
        private ISldWorks _swApp;
        private bool _isDarkTheme = true;
        private string _macroPath = @"Y:\01 PRJ - PROJETO\Projetos\Solid Defaults\Macro SW";

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
            this.Loaded += (s, e) => AtualizarInterface();
        }

        private int OnDocChange() { Dispatcher.BeginInvoke(new Action(() => AtualizarInterface())); return 0; }
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

            if (swModel == null) return;

            int type = swModel.GetType();

            if (type == (int)swDocumentTypes_e.swDocASSEMBLY || type == (int)swDocumentTypes_e.swDocPART)
            {
                PropriedadesBorder.Visibility = Visibility.Visible;
                CarregarPropriedades(swModel);

                if (type == (int)swDocumentTypes_e.swDocASSEMBLY)
                {
                    PanelMacrosMontagem.Visibility = Visibility.Visible;
                    LblTituloMacros.Text = "MACROS - MONTAGEM";
                }
                else
                {
                    PanelMacrosPeca.Visibility = Visibility.Visible;
                    LblTituloMacros.Text = "MACROS - PEÇA";
                }
            }
            else if (type == (int)swDocumentTypes_e.swDocDRAWING)
            {
                PanelMacrosDesenho.Visibility = Visibility.Visible;
                LblTituloMacros.Text = "MACROS - DESENHO";
            }
        }

        private void CarregarPropriedades(ModelDoc2 swModel)
        {
            CustomPropertyManager propMgr = swModel.Extension.get_CustomPropertyManager("");
            TxtDenominacao.Text = GetProp(propMgr, "Denominação");
            TxtNumeroDesenho.Text = GetProp(propMgr, "Número Desenho");
            TxtCodigo.Text = GetProp(propMgr, "Código");
            TxtMaterial.Text = GetProp(propMgr, "Material");
            TxtRevisao.Text = GetProp(propMgr, "Rev.");

            TxtProjetista.Text = GetProp(propMgr, "Projetista");
            TxtDataProjeto.Text = GetProp(propMgr, "Data do Projeto");
            TxtEspessura.Text = GetProp(propMgr, "Espessura");
            TxtPeso.Text = GetProp(propMgr, "Peso");
            TxtNomProjeto.Text = GetProp(propMgr, "Nom.Projeto");
            TxtConjunto.Text = GetProp(propMgr, "Conjunto");
            TxtAprov.Text = GetProp(propMgr, "Aprov.");
            TxtDatAprov.Text = GetProp(propMgr, "Dat.Aprov.");
            TxtObs.Text = GetProp(propMgr, "Obs.");
        }

        private void CopyText_Click(object sender, MouseButtonEventArgs e)
        {
            if (sender is TextBlock tb && tb.Text != "---")
            {
                Clipboard.SetText(tb.Text);
                // Opcional: Feedback visual rápido mudando a cor
                var originalBrush = tb.Foreground;
                tb.Foreground = Brushes.White;
                System.Threading.Tasks.Task.Delay(200).ContinueWith(_ => Dispatcher.Invoke(() => tb.Foreground = originalBrush));
            }
        }

        private void LimparCampos()
        {
            TxtDenominacao.Text = TxtNumeroDesenho.Text = TxtCodigo.Text = TxtMaterial.Text = TxtRevisao.Text = "---";
            TxtProjetista.Text = TxtDataProjeto.Text = TxtEspessura.Text = TxtPeso.Text = "---";
            TxtNomProjeto.Text = TxtConjunto.Text = TxtAprov.Text = TxtDatAprov.Text = TxtObs.Text = "---";
        }

        private string GetProp(CustomPropertyManager propMgr, string name)
        {
            string val = "", res = "";
            propMgr.Get4(name, false, out val, out res);
            return string.IsNullOrWhiteSpace(res) ? "---" : res;
        }

        private void ExecutarMacro(string nome)
        {
            if (SwApp == null) return;
            string path = Path.Combine(_macroPath, nome);
            if (File.Exists(path)) SwApp.RunMacro2(path, "", "main", 1, out _);
            else MessageBox.Show("Macro não encontrada: " + nome);
        }

        // Handlers de Macros (Resumidos para brevidade)
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
        private void BtnMacro26_Click(object sender, RoutedEventArgs e) => ExecutarMacro("26-Pacote de preenchimento propriedades personalizadas na Peça.swp");
        private void BtnMacro40_Click(object sender, RoutedEventArgs e) => ExecutarMacro("40-Pós importador de Engenharia.swp");
        private void BtnMacro41_Click(object sender, RoutedEventArgs e) => ExecutarMacro("41-Criar Revisão WORD.swp");
        private void BtnMacro50_Click(object sender, RoutedEventArgs e) => ExecutarMacro("50 - Cria Catalogo de peças.swp");
        private void BtnMacro55_Click(object sender, RoutedEventArgs e) => ExecutarMacro("55-Cadastra Todas as pecas Selecionadas.swp");

        private void BtnTheme_Click(object sender, RoutedEventArgs e)
        {
            _isDarkTheme = !_isDarkTheme;
            SetResource("BgColor", _isDarkTheme ? "#121212" : "#F5F5F5");
            SetResource("CardColor", _isDarkTheme ? "#1E1E1E" : "#FFFFFF");
            SetResource("MainColor", _isDarkTheme ? "#00B4FF" : "#005A9E");
            SetResource("BtnBg", _isDarkTheme ? "#2A2A2A" : "#E1E1E1");
            SetResource("BtnText", _isDarkTheme ? "#FFFFFF" : "#000000");
            SetResource("BorderColor", _isDarkTheme ? "#333333" : "#CCCCCC");
        }

        private void SetResource(string key, string colorHex) => this.Resources[key] = new SolidColorBrush((Color)ColorConverter.ConvertFromString(colorHex));

        private void BtnExpandir_Click(object sender, RoutedEventArgs e)
        {
            GridExpandido.Visibility = GridExpandido.Visibility == Visibility.Collapsed ? Visibility.Visible : Visibility.Collapsed;
            BtnExpandir.Content = GridExpandido.Visibility == Visibility.Visible ? "▲ RECOLHER" : "▼ MAIS DETALHES";
        }
    }
}