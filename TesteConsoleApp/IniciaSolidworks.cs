using System;
using System.Runtime.InteropServices;
using SolidWorks.Interop.sldworks;
using SolidWorks.Interop.swconst;

namespace Lab
{
    class Program
    {
        [STAThread]
        static void Main()
        {
            try
            {
                // 1. Conexão
                SldWorks swApp = (SldWorks)Marshal.GetActiveObject("SldWorks.Application.34");
                ModelDoc2 swModel = (ModelDoc2)swApp.ActiveDoc;

                if (swModel != null)
                {
                    Console.WriteLine("Conectado! Iniciando sequência...");

                    // --- SEQUÊNCIA DE FUNÇÕES ---

                    // Passo 1: Criar o Esboço
                    bool sketchSuccess = CreateBaseSketch(swModel, 0.1);

                    if (sketchSuccess)
                    {
                        Console.WriteLine("Lógica aqui");
                    }

                    // ----------------------------
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine("Erro na sequência: " + ex.Message);
            }

            Console.WriteLine("\nFim da operação. Pressione qualquer tecla...");
            Console.ReadKey();
        }
    }
}