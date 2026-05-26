using AddOn_DocGenerator.SL;
using SAPbouiCOM;
using System;
using System.IO;
using System.Xml;

namespace AddOn_DocGenerator.UL.Menus
{
    public class MenuLoader
    {
        public void Load()
        {
            try
            {
                var app = SapConnection.GetApplication();

                string path = Path.Combine(
                    AppDomain.CurrentDomain.BaseDirectory,
                    "Resources",
                    "Menus",
                    "Menu.xml"
                );

                if (!File.Exists(path))
                    throw new FileNotFoundException("No se encontró el archivo Menu.xml", path);

                XmlDocument xml = new XmlDocument();
                xml.Load(path);

                app.LoadBatchActions(xml.InnerXml);

                app.StatusBar.SetText(
                    "Menú Doc Generator cargado correctamente.",
                    BoMessageTime.bmt_Short,
                    BoStatusBarMessageType.smt_Success
                );
            }
            catch (Exception ex)
            {
                SapConnection.GetApplication().StatusBar.SetText(
                    $"Error cargando menú: {ex.Message}",
                    BoMessageTime.bmt_Short,
                    BoStatusBarMessageType.smt_Error
                );
            }
        }
    }
}