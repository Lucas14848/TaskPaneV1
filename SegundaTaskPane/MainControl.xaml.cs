using System.Windows;
using System.Windows.Controls;
using SolidWorks.Interop.sldworks;

namespace SegundaTaskPane
{
    public partial class MainControl : UserControl
    {
        // Use 'readonly' para o serviço conforme a sugestão do Visual Studio (Clean Code)
        private SolidWorksService _service;
        private ISldWorks _swApp;

        public ISldWorks SwApp
        {
            get => _swApp;
            set
            {
                _swApp = value;
                // Garante que o serviço seja criado assim que o Addin injetar o SwApp
                if (_swApp != null)
                {
                    _service = new SolidWorksService(_swApp);
                }
            }
        }

        public MainControl()
        {
            InitializeComponent();
        }

        // Os métodos abaixo só funcionarão se os nomes no SolidWorksService.cs 
        // forem EXATAMENTE iguais e estiverem como PUBLIC.
        private void BtnNovo_Click(object sender, RoutedEventArgs e) => _service?.CriarNovoArquivoPeça();

        private void BtnEsboco_Click(object sender, RoutedEventArgs e) => _service?.IniciarEsbocoFrontal();

        private void BtnCriarCubo_Click(object sender, RoutedEventArgs e) => _service?.CriarCubo(0.1);
    }
}