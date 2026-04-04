using System;
using SolidWorks.Interop.sldworks;
using SolidWorks.Interop.swconst;

namespace SegundaTaskPane
{
    public class ImagemService
    {
        private readonly ISldWorks _swApp;

        public ImagemService(ISldWorks swApp) => _swApp = swApp;

        public void CriarNovoArquivoPeça()
        {
            string template = _swApp.GetUserPreferenceStringValue((int)swUserPreferenceStringValue_e.swDefaultTemplatePart);
            if (!string.IsNullOrEmpty(template))
            {
                _swApp.NewDocument(template, 0, 0, 0);
            }
        }

        public void IniciarEsbocoFrontal()
        {
            ModelDoc2 swModel = (ModelDoc2)_swApp.ActiveDoc;
            if (swModel == null) return;

            bool status = swModel.Extension.SelectByID2("Plano Frontal", "PLANE", 0, 0, 0, false, 0, null, 0);
            if (!status) status = swModel.Extension.SelectByID2("Front Plane", "PLANE", 0, 0, 0, false, 0, null, 0);

            if (status) swModel.SketchManager.InsertSketch(true);
        }

        public void CriarCubo(double tamanho)
        {
            ModelDoc2 swModel = (ModelDoc2)_swApp.ActiveDoc;
            if (swModel == null) return;

            IniciarEsbocoFrontal();
            swModel.SketchManager.CreateCornerRectangle(0, 0, 0, tamanho, tamanho, 0);
            swModel.SketchManager.InsertSketch(true);

            swModel.FeatureManager.FeatureExtrusion3(true, false, false,
                (int)swEndConditions_e.swEndCondBlind, 0, tamanho, 0, false, false,
                false, false, 0, 0, false, false, false, false, true, true, true,
                (int)swStartConditions_e.swStartSketchPlane, 0, false);
        }
    }
}