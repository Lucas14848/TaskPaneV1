using System;
using System.Runtime.InteropServices;
using System.Windows.Forms.Integration;
using Microsoft.Win32;
using SolidWorks.Interop.sldworks;
using SolidWorks.Interop.swpublished;
using SolidWorks.Interop.swconst;

namespace SegundaTaskPane
{
    [Guid("C4E6A446-2943-45C8-9692-78D9902ED097"), ComVisible(true)]
    public class SolidWorksAddin : ISwAddin
    {
        private ISldWorks _swApp;
        private int _cookie;
        private TaskpaneView _taskPane1;
        public bool ConnectToSW(object ThisSW, int Cookie)
        {
            _swApp = (ISldWorks)ThisSW;
            _cookie = Cookie;

            try
            {
                _taskPane1 = _swApp.CreateTaskpaneView2(string.Empty, "Comandos SW");
                var ui1 = new ExportControl { SwApp = _swApp };
                var host1 = new ElementHost { Child = ui1 };
                _taskPane1.DisplayWindowFromHandlex64(host1.Handle.ToInt64());
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine("Erro ao criar TaskPane: " + ex.Message);
                return false;
            }

            return true;
        }

        public bool ConnectToSW(object ThisSW, int Cookie, out bool Result)
        {
            Result = ConnectToSW(ThisSW, Cookie);
            return true;
        }

        public bool DisconnectFromSW()
        {
            if (_taskPane1 != null)
            {
                _taskPane1.DeleteView();
                Marshal.ReleaseComObject(_taskPane1);
            }
            return true;
        }

        #region Registro Automático (O que faz aparecer nos Suplementos)

        [ComRegisterFunction]
        public static void RegisterFunction(Type t)
        {
            try
            {
                string keyname = "SOFTWARE\\SolidWorks\\Addins\\{" + t.GUID.ToString() + "}";
                RegistryKey addinKey = Registry.LocalMachine.CreateSubKey(keyname);
                addinKey.SetValue(null, 1);
                addinKey.SetValue("Description", "Add-in do Lucas");
                addinKey.SetValue("Title", "Comandos SW");

                keyname = "Software\\SolidWorks\\AddInsStartup\\{" + t.GUID.ToString() + "}";
                RegistryKey startupKey = Registry.CurrentUser.CreateSubKey(keyname);
                startupKey.SetValue(null, 1);
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine("Erro no registro: " + ex.Message);
            }
        }

        [ComUnregisterFunction]
        public static void UnregisterFunction(Type t)
        {
            try
            {
                Registry.LocalMachine.DeleteSubKey("SOFTWARE\\SolidWorks\\Addins\\{" + t.GUID.ToString() + "}", false);
                Registry.CurrentUser.DeleteSubKey("Software\\SolidWorks\\AddInsStartup\\{" + t.GUID.ToString() + "}", false);
            }
            catch { }
        }

        #endregion
    }
}
