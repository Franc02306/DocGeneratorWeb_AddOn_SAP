using AddOn_DocGenerator.SL;
using SAPbouiCOM;
using System;
using System.IO;

namespace AddOn_DocGenerator.UL.Resources.FormClasses
{
    public class UL_FrmMain
    {
        private const string FORM_UID = "UDO_F_DG_DOC_GEN";

        public void Show()
        {
            var app = SapConnection.GetApplication();

            try
            {
                Form form = app.Forms.Item(FORM_UID);

                form.Select();
                form.Visible = true;
            }
            catch
            {
                LoadFormFromSrf();
            }
        }

        private void LoadFormFromSrf()
        {
            var app = SapConnection.GetApplication();

            string path = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "Resources", "Forms", "FrmMain.srf");

            if (!File.Exists(path)) 
                throw new FileNotFoundException("No se encontró FrmMain.srf", path);

            string xml = File.ReadAllText(path);

            app.LoadBatchActions(xml);

            Form form = app.Forms.Item(FORM_UID);

            form.Visible = true;
        }
    }
}
