using AddOn_DocGenerator.SL;
using AddOn_DocGenerator.UL.Resources.FormClasses;
using SAPbouiCOM;
using System;

namespace AddOn_DocGenerator.UL.Events
{
    public class ItemEventHandler
    {
        private bool isRegistered;

        /// <summary>
        /// Registra el manejador global de eventos de items de SAP Business One.
        /// </summary>
        public void Register()
        {
            if (isRegistered)
                return;

            Application app = SapConnection.GetApplication();

            app.ItemEvent += OnItemEvent;

            isRegistered = true;
        }

        /// <summary>
        /// Desregistra el manejador global de eventos de items de SAP Business One.
        /// </summary>
        public void Unregister()
        {
            if (!isRegistered)
                return;

            Application app = SapConnection.GetApplication();

            app.ItemEvent -= OnItemEvent;

            isRegistered = false;
        }

        /// <summary>
        /// Recibe los eventos globales de SAP Business One y los deriva al formulario correspondiente.
        /// </summary>
        private void OnItemEvent(string formUid, ref ItemEvent eventInfo, out bool bubbleEvent)
        {
            bubbleEvent = true;

            try
            {
                RouteItemEvent(formUid, eventInfo);
            }
            catch (Exception ex)
            {
                ShowError($"Error en evento de formulario: {ex.Message}");
            }
        }

        /// <summary>
        /// Deriva el evento recibido hacia la clase encargada del formulario correspondiente.
        /// </summary>
        private static void RouteItemEvent(string formUid, ItemEvent eventInfo)
        {
            if (formUid == UL_FrmMain.FORM_UID)
            {
                UL_FrmMain.HandleItemEvent(formUid, eventInfo);
            }
        }

        /// <summary>
        /// Muestra un mensaje de error en la barra de estado de SAP Business One.
        /// </summary>
        private static void ShowError(string message)
        {
            SapConnection.GetApplication().StatusBar.SetText(
                message,
                BoMessageTime.bmt_Short,
                BoStatusBarMessageType.smt_Error
            );
        }
    }
}
