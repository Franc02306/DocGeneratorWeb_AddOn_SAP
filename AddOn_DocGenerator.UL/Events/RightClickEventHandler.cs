using AddOn_DocGenerator.SL;
using AddOn_DocGenerator.UL.Resources.FormClasses;
using SAPbouiCOM;
using System;

namespace AddOn_DocGenerator.UL.Events
{
    public class RightClickEventHandler
    {
        private readonly Application _app;

        private const string MENU_ADD_ROW = "DOCGEN_ADD_ROW";
        private const string MENU_DEL_ROW = "DOCGEN_DEL_ROW";

        /// <summary>
        /// Inicializa el manejador de clic derecho obteniendo la instancia activa de la aplicación SAP Business One.
        /// </summary>
        public RightClickEventHandler()
        {
            _app = SapConnection.GetApplication();
        }

        /// <summary>
        /// Registra el evento de clic derecho para capturar la apertura del menú contextual en SAP Business One.
        /// </summary>
        public void Register()
        {
            _app.RightClickEvent += OnRightClickEvent;
        }

        /// <summary>
        /// Procesa el evento de clic derecho y agrega opciones personalizadas al menú contextual cuando se ejecuta sobre la matrix del formulario principal.
        /// </summary>
        private void OnRightClickEvent(ref ContextMenuInfo eventInfo, out bool bubbleEvent)
        {
            bubbleEvent = true;

            try
            {
                if (eventInfo.BeforeAction == false)
                    return;

                if (eventInfo.FormUID != UL_FrmMain.FORM_UID)
                    return;

                if (eventInfo.ItemUID != UL_FrmMain.MATRIX_UID)
                    return;

                AddContextMenu();
            }
            catch (Exception ex)
            {
                _app.StatusBar.SetText($"Error menú contextual: {ex.Message}", BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
            }
        }

        /// <summary>
        /// Agrega las opciones personalizadas de agregar y eliminar línea al menú contextual de SAP Business One si aún no existen.
        /// </summary>
        private void AddContextMenu()
        {
            if (!_app.Menus.Exists(MENU_ADD_ROW))
            {
                MenuCreationParams addMenu = (MenuCreationParams)_app.CreateObject(BoCreatableObjectType.cot_MenuCreationParams);

                addMenu.Type = BoMenuType.mt_STRING;
                addMenu.UniqueID = MENU_ADD_ROW;
                addMenu.String = "Agregar línea";
                addMenu.Enabled = true;

                _app.Menus.Item("1280").SubMenus.AddEx(addMenu);
            }

            if (!_app.Menus.Exists(MENU_DEL_ROW))
            {
                MenuCreationParams delMenu = (MenuCreationParams)_app.CreateObject(BoCreatableObjectType.cot_MenuCreationParams);

                delMenu.Type = BoMenuType.mt_STRING;
                delMenu.UniqueID = MENU_DEL_ROW;
                delMenu.String = "Eliminar línea";
                delMenu.Enabled = true;

                _app.Menus.Item("1280").SubMenus.AddEx(delMenu);
            }
        }
    }
}