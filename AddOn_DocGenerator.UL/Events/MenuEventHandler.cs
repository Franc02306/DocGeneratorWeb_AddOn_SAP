using AddOn_DocGenerator.SL;
using AddOn_DocGenerator.UL.Resources.FormClasses;
using SAPbouiCOM;
using System;

namespace AddOn_DocGenerator.UL.Events
{
    public class MenuEventHandler
    {
        private readonly Application _app;

        private const string MENU_ADD_ROW = "DOCGEN_ADD_ROW";
        private const string MENU_DEL_ROW = "DOCGEN_DEL_ROW";

        /// <summary>
        /// Inicializa el manejador de eventos de menú obteniendo la instancia activa de la aplicación SAP Business One.
        /// </summary>
        public MenuEventHandler()
        {
            _app = SapConnection.GetApplication();
        }

        /// <summary>
        /// Registra el evento de menú para capturar las acciones ejecutadas desde SAP Business One.
        /// </summary>
        public void Register()
        {
            _app.MenuEvent += OnMenuEvent;
        }

        /// <summary>
        /// Procesa los eventos de menú, abre el formulario principal y ejecuta acciones sobre la matrix según la opción seleccionada.
        /// </summary>
        private void OnMenuEvent(ref MenuEvent pVal, out bool bubbleEvent)
        {
            bubbleEvent = true;

            try
            {
                if (pVal.BeforeAction)
                    return;

                if (pVal.MenuUID == "DOCGEN_GENERATOR")
                {
                    new UL_FrmMain().Show();
                    return;
                }

                Form activeForm = _app.Forms.ActiveForm;

                if (activeForm.UniqueID != UL_FrmMain.FORM_UID)
                    return;

                switch (pVal.MenuUID)
                {
                    case MENU_ADD_ROW:
                        UL_FrmMain.AddMatrixRow(activeForm);
                        break;

                    case MENU_DEL_ROW:
                        UL_FrmMain.DeleteSelectedMatrixRow(activeForm);
                        break;
                }
            }
            catch (Exception ex)
            {
                _app.StatusBar.SetText($"Error en menú: {ex.Message}", BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
            }
        }
    }
}