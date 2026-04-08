using System;
using System.IO;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Media;
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

        private int OnFileClose(string name, int reason)
        {
            Dispatcher.BeginInvoke(new Action(() => AtualizarInterface()));
            return 0;
        }

        public void AtualizarInterface()
        {
            if (SwApp == null) return;
            ModelDoc2 swModel = (ModelDoc2)SwApp.ActiveDoc;

            LimparCampos();
            PropriedadesBorder.Visibility = Visibility.Collapsed;
            PanelMacrosPeca.Visibility = Visibility.Collapsed;
            PanelMacrosMontagem.Visibility = Visibility.Collapsed;

            if (swModel == null) return;

            int type = swModel.GetType();

            // Lógica para MONTAGEM
            if (type == (int)swDocumentTypes_e.swDocASSEMBLY)
            {
                PropriedadesBorder.Visibility = Visibility.Visible;
                PanelMacrosMontagem.Visibility = Visibility.Visible;
                LblTituloMacros.Text = "LISTA DE MACROS - MONTAGEM";
                CarregarPropriedades(swModel);
            }
            // Lógica para PEÇA
            else if (type == (int)swDocumentTypes_e.swDocPART)
            {
                PropriedadesBorder.Visibility = Visibility.Visible;
                PanelMacrosPeca.Visibility = Visibility.Visible;
                LblTituloMacros.Text = "LISTA DE MACROS - PEÇA";
                CarregarPropriedades(swModel);
            }
        }

        private void CarregarPropriedades(ModelDoc2 swModel)
        {
            CustomPropertyManager propMgr = swModel.Extension.get_CustomPropertyManager("");
            TxtDenominacao.Text = GetProp(propMgr, "Denominação");
            TxtNumeroDesenho.Text = GetProp(propMgr, "Número Desenho");
            TxtMaterial.Text = GetProp(propMgr, "Material");
            TxtRevisao.Text = GetProp(propMgr, "Rev.");
        }

        private void LimparCampos()
        {
            TxtDenominacao.Text = TxtNumeroDesenho.Text = TxtMaterial.Text = TxtRevisao.Text = "---";
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
            AtualizarInterface();
        }

        // Handlers de Clique
        private void BtnMacro01_Click(object sender, RoutedEventArgs e) => ExecutarMacro("01-Cria uma folha de desenho com 3 vistas.swp");
        private void BtnMacro03_Click(object sender, RoutedEventArgs e) => ExecutarMacro("03-Renomear, Criar Cópias e Criar Revisões em Peças e Montagens.swp");
        private void BtnMacro04_Click(object sender, RoutedEventArgs e) => ExecutarMacro("04-Salva PDFs de uma Montagem-ERP.swp");
        private void BtnMacro06_Click(object sender, RoutedEventArgs e) => ExecutarMacro("06-Soldagem 3D.swp");
        private void BtnMacro07_Click(object sender, RoutedEventArgs e) => ExecutarMacro("07-Lista de Chapas DXF QTY.swp");
        private void BtnMacro14_Click(object sender, RoutedEventArgs e) => ExecutarMacro("14-Navegador catalogos.swp");
        private void BtnMacro15_Click(object sender, RoutedEventArgs e) => ExecutarMacro("15-Cria propriedade espessura de chapa.swp");
        private void BtnMacro16_Click(object sender, RoutedEventArgs e) => ExecutarMacro("16-Cria cores chapa metalica.swp");
        private void BtnMacro17_Click(object sender, RoutedEventArgs e) => ExecutarMacro("17-Cria cavidades de alivio.swp");
        private void BtnMacro18_Click(object sender, RoutedEventArgs e) => ExecutarMacro("18-Excel Importação CIGAM.swp");
        private void BtnMacro22_Click(object sender, RoutedEventArgs e) => ExecutarMacro("22-Criar e Renomear Mangueiras.swp");
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
    }
}