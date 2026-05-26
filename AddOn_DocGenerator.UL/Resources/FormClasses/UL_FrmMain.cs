using AddOn_DocGenerator.SL;
using SAPbouiCOM;
using System;
using System.IO;

namespace AddOn_DocGenerator.UL.Resources.FormClasses
{
    public class UL_FrmMain
    {
        public const string FORM_UID = "UDO_F_DG_DOC_GEN";
        public const string MATRIX_UID = "0_U_G";
        private const string COL_NUM = "#";

        public void Show()
        {
            var app = SapConnection.GetApplication();

            try
            {
                Form form = app.Forms.Item(FORM_UID);
                form.Select();

                InitMatrix(form);
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

            InitMatrix(form);

            form.Visible = true;
        }

        public static void InitMatrix(Form form)
        {
            Matrix matrix = (Matrix)form.Items.Item(MATRIX_UID).Specific;

            if (matrix.RowCount == 0)
            {
                matrix.AddRow(1);
            }

            RenumberMatrix(matrix);

            matrix.FlushToDataSource();
        }

        private static void RenumberMatrix(Matrix matrix)
        {
            for (int i = 1; i <= matrix.RowCount; i++)
            {
                EditText cell = (EditText)matrix.Columns.Item(COL_NUM).Cells.Item(i).Specific;
                cell.Value = i.ToString();
            }
        }

        public static void AddMatrixRow(Form form)
        {
            Matrix matrix = (Matrix)form.Items.Item(MATRIX_UID).Specific;

            matrix.AddRow(1);

            RenumberMatrix(matrix);

            matrix.FlushToDataSource();
        }

        public static void DeleteSelectedMatrixRow(Form form)
        {
            Matrix matrix = (Matrix)form.Items.Item(MATRIX_UID).Specific;

            int selectedRow = matrix.GetNextSelectedRow(0, BoOrderType.ot_RowOrder);

            if (selectedRow <= 0)
                throw new Exception("Selecciona una línea de la matrix.");

            matrix.DeleteRow(selectedRow);

            if (matrix.RowCount == 0)
                matrix.AddRow(1);

            RenumberMatrix(matrix);

            matrix.FlushToDataSource();
        }
    }
}