using SolidWorks.Interop.sldworks;
using System;
using System.Runtime.InteropServices;

public class FuncoesBasicas
{
    // Estas são as suas "Variáveis Globais"
    public static SldWorks swApp;
    public static ModelDoc2 swModel;

    public static void Conectar()
    {
        try
        {
            // Conecta ao SW 2026
            swApp = (SldWorks)Marshal.GetActiveObject("SldWorks.Application.34");

            // Define o documento ativo
            swModel = (ModelDoc2)swApp.ActiveDoc;

            if (swModel != null)
            {
                Console.WriteLine("Conectado com sucesso ao: " + swModel.GetTitle());
            }
            else
            {
                Console.WriteLine("SolidWorks conectado, mas não há documento ativo.");
            }
        }
        catch (Exception ex)
        {
            Console.WriteLine("Erro ao conectar: " + ex.Message);
        }
    }

    public static void CriarEsboçoQuadrado(double lado)
    {
        // Agora você não precisa mais passar swModel como parâmetro, 
        // pois ele já está "global" nesta classe.
        if (swModel == null) return;

        swModel.Extension.SelectByID2("Plano Frontal", "PLANE", 0, 0, 0, false, 0, null, 0);
        swModel.SketchManager.InsertSketch(true);

        swModel.SketchManager.CreateLine(0, 0, 0, lado, 0, 0);
        swModel.SketchManager.CreateLine(lado, 0, 0, lado, lado, 0);
        swModel.SketchManager.CreateLine(lado, lado, 0, 0, lado, 0);
        swModel.SketchManager.CreateLine(0, lado, 0, 0, 0, 0);

        swModel.SketchManager.InsertSketch(true);
    }
}