using System;
using System.Windows.Forms;
using AddOn_DocGenerator.SL;
using AddOn_DocGenerator.UL.Events;
using AddOn_DocGenerator.UL.Menus;

namespace AddOn_DocGenerator
{
    internal static class Program
    {
        [STAThread]
        static void Main(string[] args)
        {
            try
            {
                SapConnection.Connect();

                new MenuLoader().Load();

                new MenuEventHandler().Register();
                new RightClickEventHandler().Register();

                Application.Run();
            }
            catch (Exception ex)
            {
                System.IO.File.WriteAllText("AddOn_DocGenerator_Error.txt", ex.ToString());
            }
        }
    }
}