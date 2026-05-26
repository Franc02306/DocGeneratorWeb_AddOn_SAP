using AddOn_DocGenerator.SL;
using AddOn_DocGenerator.UL.Resources.FormClasses;
using SAPbouiCOM;
using System;

namespace AddOn_DocGenerator.UL.Events
{
    public class MenuEventHandler
    {
        private readonly Application _app;

        public MenuEventHandler()
        {
            _app = SapConnection.GetApplication();
        }

        public void Register()
        {
            _app.MenuEvent += OnMenuEvent;
        }

        private void OnMenuEvent(ref MenuEvent pVal, out bool bubbleEvent)
        {
            bubbleEvent = true;

            try
            {
                if (pVal.BeforeAction)
                    return;

                if (pVal.MenuUID == "DOCGEN_GENERATOR")
                {
                    var form = new UL_FrmMain();
                    form.Show();
                }
            }
            catch (Exception ex)
            {
                _app.StatusBar.SetText($"Error abriendo formulario: {ex.Message}", BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
            }
        }
    }
}
