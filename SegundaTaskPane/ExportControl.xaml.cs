using System.Windows;
using System.Windows.Controls;
using SolidWorks.Interop.sldworks;

namespace SegundaTaskPane
{
    public partial class ExportControl : UserControl
    {
        public ISldWorks SwApp { get; set; }

        public ExportControl()
        {
            InitializeComponent();
        }

        private void BtnExportPdf_Click(object sender, RoutedEventArgs e)
        {
            if (SwApp == null) return;
            MessageBox.Show("Lógica de PDF será inserida aqui!");
        }
    }
}