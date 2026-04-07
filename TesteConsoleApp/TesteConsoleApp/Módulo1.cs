using SolidWorks.Interop.swconst;
using System;

class Módulo1
{
    [STAThread]
    static void Main()
    {
        FuncoesBasicas.Conectar();
        FuncoesBasicas.CriarEsboçoQuadrado(0.1);

        string valor = FuncoesBasicas.swModel.GetTitle();

        Console.WriteLine("\nFim da operação. Pressione qualquer tecla...");
        Console.ReadKey();
    }
}