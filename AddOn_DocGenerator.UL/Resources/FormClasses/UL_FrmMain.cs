using AddOn_DocGenerator.SL;
using SAPbouiCOM;
using System;
using System.IO;

namespace AddOn_DocGenerator.UL.Resources.FormClasses
{
    public class UL_FrmMain
    {
        // MAIN
        public const string FORM_UID = "UDO_F_DG_DOC_GEN";
        private const string DS_DETAIL = "@DG_DOC_DET";

        // CFL
        private const string SUPPLIER_CFL_UID = "CFL_OCRD";

        // UID
        private const string COL_NUM = "#";

        public const string MATRIX_UID = "MtxDocs";
        private const string COL_CARD_CODE = "ClmCodPrv";
        private const string COL_RUC = "ClmRucPrv";
        private const string COL_CARD_NAME = "ClmRznPrv";

        // CLM
        private const string FIELD_CARD_CODE = "U_CARDCODE";
        private const string FIELD_RUC = "U_RUC";
        private const string FIELD_CARD_NAME = "U_CARDNAME";

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

        /// <summary>
        /// Configura el Choose From List de proveedores para mostrar únicamente socios de negocio tipo proveedor cuyo código empiece con "P".
        /// </summary>
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

        /// <summary>
        /// Maneja los eventos del formulario principal derivados desde el manejador global de eventos.
        /// </summary>
        public static void HandleItemEvent(string formUid, ItemEvent eventInfo)
        {
            HandleSupplierChooseFromList(formUid, eventInfo);
        }

        /// <summary>
        /// Maneja la selección del proveedor desde el Choose From List y completa RUC y razón social en la matrix.
        /// </summary>
        private static void HandleSupplierChooseFromList(string formUid, ItemEvent eventInfo)
        {
            if (formUid != FORM_UID)
                return;

            if (eventInfo.EventType != BoEventTypes.et_CHOOSE_FROM_LIST)
                return;

            if (eventInfo.BeforeAction)
                return;

            if (eventInfo.ItemUID != MATRIX_UID)
                return;

            if (eventInfo.ColUID != COL_CARD_CODE)
                return;

            if (eventInfo.Row <= 0)
                return;

            IChooseFromListEvent chooseFromListEvent = (IChooseFromListEvent)eventInfo;

            if (chooseFromListEvent.ChooseFromListUID != SUPPLIER_CFL_UID)
                return;

            DataTable selectedObjects = chooseFromListEvent.SelectedObjects;

            if (selectedObjects == null || selectedObjects.Rows.Count == 0)
                return;

            string cardCode = GetDataTableValue(selectedObjects, "CardCode");
            string ruc = GetDataTableValue(selectedObjects, "LicTradNum");
            string cardName = GetDataTableValue(selectedObjects, "CardName");

            Application app = SapConnection.GetApplication();
            Form form = app.Forms.Item(formUid);

            try
            {
                form.Freeze(true);

                Matrix matrix = (Matrix)form.Items.Item(MATRIX_UID).Specific;
                DBDataSource detailDataSource = form.DataSources.DBDataSources.Item(DS_DETAIL);

                matrix.FlushToDataSource();

                int dataSourceRow = eventInfo.Row - 1;

                detailDataSource.SetValue(FIELD_CARD_CODE, dataSourceRow, cardCode);
                detailDataSource.SetValue(FIELD_RUC, dataSourceRow, ruc);
                detailDataSource.SetValue(FIELD_CARD_NAME, dataSourceRow, cardName);

                matrix.LoadFromDataSource();

                RenumberMatrix(matrix);

                matrix.FlushToDataSource();
            }
            finally
            {
                form.Freeze(false);
            }
        }

        /// <summary>
        /// Obtiene un valor seguro desde el DataTable retornado por el Choose From List.
        /// </summary>
        private static string GetDataTableValue(DataTable dataTable, string columnName)
        {
            object value = dataTable.GetValue(columnName, 0);

            return value == null ? string.Empty : value.ToString();
        }
    }
}