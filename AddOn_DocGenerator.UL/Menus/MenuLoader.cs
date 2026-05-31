using AddOn_DocGenerator.SL;
using SAPbouiCOM;
using System;
using System.IO;
using System.Xml;

namespace AddOn_DocGenerator.UL.Menus
{
    public class MenuLoader
    {
        private const string MAIN_MENU_UID = "DOCGEN_MAIN";
        private const string GENERATOR_MENU_UID = "DOCGEN_GENERATOR";

        /// <summary>
        /// Carga el menú personalizado del AddOn desde el archivo XML y lo registra en SAP Business One.
        /// </summary>
        public void Load()
        {
            try
            {
                Application app = SapConnection.GetApplication();

                string menuPath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "Resources", "Menus", "Menu.xml");

                string iconPath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "Resources", "Images", "doc-gen-menu.bmp");

                if (!File.Exists(menuPath))
                    throw new FileNotFoundException("No se encontró el archivo Menu.xml", menuPath);

                if (!File.Exists(iconPath))
                    throw new FileNotFoundException("No se encontró el ícono doc-gen-menu.bmp", iconPath);

                RemoveMenuIfExists(app, GENERATOR_MENU_UID);
                RemoveMenuIfExists(app, MAIN_MENU_UID);

                XmlDocument xml = new XmlDocument();
                xml.Load(menuPath);

                XmlNode mainMenuNode = xml.SelectSingleNode($"//Menu[@UniqueID='{MAIN_MENU_UID}']");

                if (mainMenuNode?.Attributes != null)
                    mainMenuNode.Attributes["Image"].Value = iconPath;

                app.LoadBatchActions(xml.InnerXml);

                app.StatusBar.SetText("Menú Doc Generator cargado correctamente.", BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Success);
            }
            catch (Exception ex)
            {
                SapConnection.GetApplication().StatusBar.SetText($"Error cargando menú: {ex.Message}", BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
            }
        }

        /// <summary>
        /// Elimina un menú de SAP Business One si ya existe, evitando errores por duplicidad al volver a cargarlo.
        /// </summary>
        private static void RemoveMenuIfExists(Application app, string menuUid)
        {
            try
            {
                app.Menus.Item(menuUid);
                app.Menus.RemoveEx(menuUid);
            }
            catch
            {
            }
        }
    }
}