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

        // CFL
        private const string SUPPLIER_CFL_UID = "CFL_OCRD";

        /// <summary>
        /// Muestra el formulario principal del UDO si ya está abierto; de lo contrario, lo carga desde el archivo SRF.
        /// </summary>
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

        /// <summary>
        /// Carga el formulario principal desde el archivo SRF, inicializa la matrix y muestra el formulario.
        /// </summary>
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

        /// <summary>
        /// Inicializa la matrix del formulario, asegurando que tenga al menos una fila y que sus líneas estén numeradas.
        /// </summary>
        public static void InitMatrix(Form form)
        {
            ConfigureSupplierChooseFromList(form);

            Matrix matrix = (Matrix)form.Items.Item(MATRIX_UID).Specific;

            if (matrix.RowCount == 0)
            {
                matrix.AddRow(1);
            }

            RenumberMatrix(matrix);

            matrix.FlushToDataSource();
        }

        /// <summary>
        /// Renumera las filas de la matrix asignando el número de línea correspondiente a la columna de numeración.
        /// </summary>
        private static void RenumberMatrix(Matrix matrix)
        {
            for (int i = 1; i <= matrix.RowCount; i++)
            {
                EditText cell = (EditText)matrix.Columns.Item(COL_NUM).Cells.Item(i).Specific;
                cell.Value = i.ToString();
            }
        }

        /// <summary>
        /// Agrega una nueva fila a la matrix, renumera sus líneas y sincroniza los datos con el datasource.
        /// </summary>
        public static void AddMatrixRow(Form form)
        {
            Matrix matrix = (Matrix)form.Items.Item(MATRIX_UID).Specific;

            matrix.AddRow(1);

            RenumberMatrix(matrix);

            matrix.FlushToDataSource();
        }

        /// <summary>
        /// Elimina la fila seleccionada de la matrix, asegura que quede al menos una fila, renumera las líneas y sincroniza los datos con el datasource.
        /// </summary>
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

        private static void ConfigureSupplierChooseFromList(Form form)
        {
            ChooseFromList supplierCfl = form.ChooseFromLists.Item(SUPPLIER_CFL_UID);

            Conditions conditions = new Conditions();

            Condition condition = conditions.Add();
            condition.Alias = "CardType";
            condition.Operation = BoConditionOperation.co_EQUAL;
            condition.CondVal = "S";
            condition.Relationship = BoConditionRelationship.cr_AND;

            condition = conditions.Add();
            condition.Alias = "CardCode";
            condition.Operation = BoConditionOperation.co_START;
            condition.CondVal = "P";

            supplierCfl.SetConditions(conditions);
        }
    }
}