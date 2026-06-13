using AddOn_DocGenerator.SL;
using SAPbouiCOM;
using System;
using System.IO;

namespace AddOn_DocGenerator.UL.Resources.FormClasses
{
    public class UL_FrmMain
    {
        #region DEFINICIÓN DE VARIABLES DE FORMULARIO

        // MAIN
        public const string FORM_UID = "UDO_F_DG_DOC_GEN";
        private const string DS_DETAIL = "@DG_DOC_DET";

        // CFL
        private const string CRD_CFL_UID = "CFL_OCRD";
        private const string PRJ_CFL_UID = "CFL_OPRJ";
        private const string WHS_CFL_UID = "CFL_OWHS";
        private const string ITM_CFL_UID = "CFL_OITM";
        private const string PRC_CFL_UID_1 = "CFL_OPRC1"; // Dimensión 1
        private const string PRC_CFL_UID_2 = "CFL_OPRC2"; // Dimensión 2
        private const string PRC_CFL_UID_3 = "CFL_OPRC3"; // Dimensión 3
        private const string PRC_CFL_UID_4 = "CFL_OPRC4"; // Dimensión 4
        private const string PRC_CFL_UID_5 = "CFL_OPRC5"; // Dimensión 5

        // COL
        public const string MATRIX_UID = "MtxDocs";
        private const string COL_NUM = "#";
        private const string COL_CARD_CODE = "ClmCodPrv";
        private const string COL_RUC = "ClmRucPrv";
        private const string COL_CARD_NAME = "ClmRznPrv";
        private const string COL_WHS_CODE = "ClmCodWhs";
        private const string COL_PRJ_CODE = "ClmCodPrj";
        private const string COL_ITM_CODE = "ClmItmCde";
        private const string COL_DIM1 = "ClmDim1";
        private const string COL_DIM2 = "ClmDim2";
        private const string COL_DIM3 = "ClmDim3";
        private const string COL_DIM4 = "ClmDim4";
        private const string COL_DIM5 = "ClmDim5";

        // FIELD
        private const string FIELD_CARD_CODE = "U_CARDCODE";
        private const string FIELD_RUC = "U_RUC";
        private const string FIELD_CARD_NAME = "U_CARDNAME";
        private const string FIELD_TAX_DATE = "U_TAXDATE";
        private const string FIELD_DOC_DATE = "U_DOCDATE";
        private const string FIELD_DUE_DATE = "U_DUEDATE";
        private const string FIELD_CURRENCY = "U_CURRENCY";
        private const string FIELD_WHS_CODE = "U_WHS_CDE";
        private const string FIELD_PRJ_CODE = "U_PRJ_CDE";

        #endregion

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
            ConfigureProjectChooseFromList(form);
            ConfigureWarehouseChooseFromList(form);

            Matrix matrix = (Matrix)form.Items.Item(MATRIX_UID).Specific;

            if (matrix.RowCount == 0)
            {
                matrix.AddRow(1);
            }

            SetDefaultValuesForCreateMode(form, matrix);

            RenumberMatrix(matrix);

            matrix.FlushToDataSource();
        }

        /// <summary>
        /// Setear data por defecto en la matrix, sólo en MODO CREAR
        /// </summary>
        private static void SetDefaultValuesForCreateMode(Form form, Matrix matrix)
        {
            if (form.Mode != BoFormMode.fm_ADD_MODE)
                return;

            DBDataSource detailDataSource = form.DataSources.DBDataSources.Item(DS_DETAIL);

            matrix.FlushToDataSource();

            string today = DateTime.Now.ToString("yyyyMMdd");

            for (int i = 0; i < detailDataSource.Size; i++)
            {
                if (string.IsNullOrWhiteSpace(detailDataSource.GetValue(FIELD_TAX_DATE, i)))
                    detailDataSource.SetValue(FIELD_TAX_DATE, i, today);

                if (string.IsNullOrWhiteSpace(detailDataSource.GetValue(FIELD_DOC_DATE, i)))
                    detailDataSource.SetValue(FIELD_DOC_DATE, i, today);

                if (string.IsNullOrWhiteSpace(detailDataSource.GetValue(FIELD_DUE_DATE, i)))
                    detailDataSource.SetValue(FIELD_DUE_DATE, i, today);

                if (string.IsNullOrWhiteSpace(detailDataSource.GetValue(FIELD_CURRENCY, i)))
                    detailDataSource.SetValue(FIELD_CURRENCY, i, "SOL");
            }

            matrix.LoadFromDataSource();
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
        /// Configura el CFL de proveedores para mostrar únicamente socios de negocio tipo proveedor cuyo código empiece con "P".
        /// </summary>
        private static void ConfigureSupplierChooseFromList(Form form)
        {
            ChooseFromList supplierCfl = form.ChooseFromLists.Item(CRD_CFL_UID);

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
            condition.Relationship = BoConditionRelationship.cr_AND;

            condition = conditions.Add();
            condition.Alias = "validFor";
            condition.Operation = BoConditionOperation.co_EQUAL;
            condition.CondVal = "Y";
            condition.Relationship = BoConditionRelationship.cr_AND;

            condition = conditions.Add();
            condition.Alias = "frozenFor";
            condition.Operation = BoConditionOperation.co_EQUAL;
            condition.CondVal = "N";

            supplierCfl.SetConditions(conditions);
        }

        /// <summary>
        /// Configura el CFL de proyectos para mostrar únicamente los proyectos activos.
        /// </summary>
        private static void ConfigureProjectChooseFromList(Form form)
        {
            ChooseFromList projectCfl = form.ChooseFromLists.Item(PRJ_CFL_UID);

            Conditions conditions = new Conditions();

            Condition condition = conditions.Add();
            condition.Alias = "Active";
            condition.Operation = BoConditionOperation.co_EQUAL;
            condition.CondVal = "Y";

            projectCfl.SetConditions(conditions);
        }

        /// <summary>
        /// Configura el CFL de almacénes para mostrar únicamente los almacénes activos.
        /// </summary>
        private static void ConfigureWarehouseChooseFromList(Form form)
        {
            ChooseFromList warehouseCfl = form.ChooseFromLists.Item(WHS_CFL_UID);

            Conditions conditions = new Conditions();

            Condition condition = conditions.Add();
            condition.Alias = "Inactive";
            condition.Operation = BoConditionOperation.co_EQUAL;
            condition.CondVal = "N";

            warehouseCfl.SetConditions(conditions);
        }

        /// <summary>
        /// Configura el CFL de artículos para mostrar únicamente artículos activos.
        /// </summary>
        private static void ConfigureItemChooseFromList(Form form)
        {
            ChooseFromList itemCfl = form.ChooseFromLists.Item(ITM_CFL_UID);

            Conditions conditions = new Conditions();

            Condition condition = conditions.Add();
            condition.Alias = "frozenFor";
            condition.Operation = BoConditionOperation.co_EQUAL;
            condition.CondVal = "N";

            itemCfl.SetConditions(conditions);
        }

        /// <summary>
        /// Configura el CFL de centros de costo para mostrar únicamente los centros de costo activos.
        /// </summary>
        private static void ConfigureCostCenterChooseFromList(Form form, string cflUid, string dimensionCode)
        {
            ChooseFromList costCenterCfl = form.ChooseFromLists.Item(cflUid);

            Conditions conditions = new Conditions();

            Condition condition = conditions.Add();
            condition.Alias = "Active";
            condition.Operation = BoConditionOperation.co_EQUAL;
            condition.CondVal = "Y";
            condition.Relationship = BoConditionRelationship.cr_AND;

            condition = conditions.Add();
            condition.Alias = "DimCode";
            condition.Operation = BoConditionOperation.co_EQUAL;
            condition.CondVal = dimensionCode;

            costCenterCfl.SetConditions(conditions);
        }

        /// <summary>
        /// Maneja los eventos del formulario principal derivados desde el manejador global de eventos.
        /// </summary>
        public static void HandleItemEvent(string formUid, ItemEvent eventInfo)
        {
            ConfigureCostCenterChooseFromListBeforeOpen(formUid, eventInfo);
            HandleSupplierChooseFromList(formUid, eventInfo);
            HandleCostCenterChooseFromList(formUid, eventInfo);
            HandleWarehouseChooseFromList(formUid, eventInfo);
            HandleProjectChooseFromList(formUid, eventInfo);
            HandleItemChooseFromList(formUid, eventInfo);
        }

        /// <summary>
        /// Gestión de CFL para consultar centros de costo según dimensión enviada
        /// </summary>
        private static void ConfigureCostCenterChooseFromListBeforeOpen(string formUid, ItemEvent eventInfo)
        {
            if (formUid != FORM_UID)
                return;

            if (eventInfo.EventType != BoEventTypes.et_CHOOSE_FROM_LIST)
                return;

            if (!eventInfo.BeforeAction)
                return;

            if (eventInfo.ItemUID != MATRIX_UID)
                return;

            if (eventInfo.Row <= 0)
                return;

            string cflUid;
            string dimensionCode;

            switch (eventInfo.ColUID)
            {
                case COL_DIM1:
                    cflUid = PRC_CFL_UID_1;
                    dimensionCode = "1";
                    break;

                case COL_DIM2:
                    cflUid = PRC_CFL_UID_2;
                    dimensionCode = "2";
                    break;

                case COL_DIM3:
                    cflUid = PRC_CFL_UID_3;
                    dimensionCode = "3";
                    break;

                case COL_DIM4:
                    cflUid = PRC_CFL_UID_4;
                    dimensionCode = "4";
                    break;

                case COL_DIM5:
                    cflUid = PRC_CFL_UID_5;
                    dimensionCode = "5";
                    break;

                default:
                    return;
            }

            Application app = SapConnection.GetApplication();
            Form form = app.Forms.Item(formUid);

            ConfigureCostCenterChooseFromList(form, cflUid, dimensionCode);
        }

        /// <summary>
        /// Maneja la selección del proveedor desde el CFL y completa RUC y razón social en la matrix.
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

            if (chooseFromListEvent.ChooseFromListUID != CRD_CFL_UID)
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
        /// Maneja la selección del centro de costo desde el CFL y lo coloca en la matrix.
        /// </summary>
        private static void HandleCostCenterChooseFromList(string formUid, ItemEvent eventInfo)
        {
            if (formUid != FORM_UID)
                return;

            if (eventInfo.EventType != BoEventTypes.et_CHOOSE_FROM_LIST)
                return;

            if (eventInfo.BeforeAction)
                return;

            if (eventInfo.ItemUID != MATRIX_UID)
                return;

            if (eventInfo.Row <= 0)
                return;

            string fieldName;

            switch (eventInfo.ColUID)
            {
                case COL_DIM1:
                    fieldName = "U_DIM1";
                    break;

                case COL_DIM2:
                    fieldName = "U_DIM2";
                    break;

                case COL_DIM3:
                    fieldName = "U_DIM3";
                    break;

                case COL_DIM4:
                    fieldName = "U_DIM4";
                    break;

                case COL_DIM5:
                    fieldName = "U_DIM5";
                    break;

                default:
                    return;
            }

            IChooseFromListEvent chooseFromListEvent = (IChooseFromListEvent)eventInfo;

            DataTable selectedObjects = chooseFromListEvent.SelectedObjects;

            if (selectedObjects == null || selectedObjects.Rows.Count == 0)
                return;

            string prcCode = GetDataTableValue(selectedObjects, "PrcCode");

            Application app = SapConnection.GetApplication();
            Form form = app.Forms.Item(formUid);

            Matrix matrix = (Matrix)form.Items.Item(MATRIX_UID).Specific;
            DBDataSource detailDataSource = form.DataSources.DBDataSources.Item(DS_DETAIL);

            matrix.FlushToDataSource();

            int dataSourceRow = eventInfo.Row - 1;

            detailDataSource.SetValue(fieldName, dataSourceRow, prcCode);

            matrix.LoadFromDataSource();

            RenumberMatrix(matrix);

            matrix.FlushToDataSource();
        }

        /// <summary>
        /// Maneja la selección del almacén desde el CFL y lo coloca en la matrix.
        /// </summary>
        private static void HandleWarehouseChooseFromList(string formUid, ItemEvent eventInfo)
        {
            if (formUid != FORM_UID)
                return;

            if (eventInfo.EventType != BoEventTypes.et_CHOOSE_FROM_LIST)
                return;

            if (eventInfo.BeforeAction)
                return;

            if (eventInfo.ItemUID != MATRIX_UID)
                return;

            if (eventInfo.ColUID != COL_WHS_CODE)
                return;

            if (eventInfo.Row <= 0)
                return;

            IChooseFromListEvent chooseFromListEvent = (IChooseFromListEvent)eventInfo;

            if (chooseFromListEvent.ChooseFromListUID != WHS_CFL_UID)
                return;

            DataTable selectedObjects = chooseFromListEvent.SelectedObjects;

            if (selectedObjects == null || selectedObjects.Rows.Count == 0)
                return;

            string whsCode = GetDataTableValue(selectedObjects, "WhsCode");

            Application app = SapConnection.GetApplication();
            Form form = app.Forms.Item(formUid);

            Matrix matrix = (Matrix)form.Items.Item(MATRIX_UID).Specific;
            DBDataSource detailDataSource = form.DataSources.DBDataSources.Item(DS_DETAIL);

            matrix.FlushToDataSource();

            int dataSourceRow = eventInfo.Row - 1;

            detailDataSource.SetValue(FIELD_WHS_CODE, dataSourceRow, whsCode);

            matrix.LoadFromDataSource();

            RenumberMatrix(matrix);

            matrix.FlushToDataSource();
        }

        /// <summary>
        /// Maneja la selección del proyecto desde el CFL y lo coloca en la matrix.
        /// </summary>
        private static void HandleProjectChooseFromList(string formUid, ItemEvent eventInfo)
        {
            if (formUid != FORM_UID)
                return;

            if (eventInfo.EventType != BoEventTypes.et_CHOOSE_FROM_LIST)
                return;

            if (eventInfo.BeforeAction)
                return;

            if (eventInfo.ItemUID != MATRIX_UID)
                return;

            if (eventInfo.ColUID != COL_PRJ_CODE)
                return;

            if (eventInfo.Row <= 0)
                return;

            IChooseFromListEvent chooseFromListEvent = (IChooseFromListEvent)eventInfo;

            if (chooseFromListEvent.ChooseFromListUID != PRJ_CFL_UID)
                return;

            DataTable selectedObjects = chooseFromListEvent.SelectedObjects;

            if (selectedObjects == null || selectedObjects.Rows.Count == 0)
                return;

            string prjCode = GetDataTableValue(selectedObjects, "PrjCode");

            Application app = SapConnection.GetApplication();
            Form form = app.Forms.Item(formUid);

            Matrix matrix = (Matrix)form.Items.Item(MATRIX_UID).Specific;
            DBDataSource detailDataSource = form.DataSources.DBDataSources.Item(DS_DETAIL);

            matrix.FlushToDataSource();

            int dataSourceRow = eventInfo.Row - 1;

            detailDataSource.SetValue(FIELD_PRJ_CODE, dataSourceRow, prjCode);

            matrix.LoadFromDataSource();

            RenumberMatrix(matrix);

            matrix.FlushToDataSource();
        }

        /// <summary>
        /// Maneja la selección del artículo desde el CFL y lo coloca en la matrix.
        /// </summary>
        private static void HandleItemChooseFromList(string formUid, ItemEvent eventInfo)
        {
            if (formUid != FORM_UID)
                return;

            if (eventInfo.EventType != BoEventTypes.et_CHOOSE_FROM_LIST)
                return;

            if (eventInfo.BeforeAction)
                return;

            if (eventInfo.ItemUID != MATRIX_UID)
                return;

            if (eventInfo.ColUID != COL_ITM_CODE)
                return;

            if (eventInfo.Row <= 0)
                return;

            IChooseFromListEvent chooseFromListEvent = (IChooseFromListEvent)eventInfo;

            if (chooseFromListEvent.ChooseFromListUID != ITM_CFL_UID)
                return;

            DataTable selectedObjects = chooseFromListEvent.SelectedObjects;

            if (selectedObjects == null || selectedObjects.Rows.Count == 0)
                return;

            string itemCode = GetDataTableValue(selectedObjects, "ItemCode");

            Application app = SapConnection.GetApplication();
            Form form = app.Forms.Item(formUid);

            Matrix matrix = (Matrix)form.Items.Item(MATRIX_UID).Specific;
            DBDataSource detailDataSource = form.DataSources.DBDataSources.Item(DS_DETAIL);

            matrix.FlushToDataSource();

            int dataSourceRow = eventInfo.Row - 1;

            string fieldName = matrix.Columns.Item(COL_ITM_CODE).DataBind.Alias;

            detailDataSource.SetValue(fieldName, dataSourceRow, itemCode);

            matrix.LoadFromDataSource();

            RenumberMatrix(matrix);

            matrix.FlushToDataSource();
        }

        /// <summary>
        /// Obtiene un valor seguro desde el DataTable retornado por el CFL.
        /// </summary>
        private static string GetDataTableValue(DataTable dataTable, string columnName)
        {
            object value = dataTable.GetValue(columnName, 0);

            return value == null ? string.Empty : value.ToString();
        }
    }
}