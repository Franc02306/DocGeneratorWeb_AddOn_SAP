using SAPbobsCOM;
using SAPbouiCOM;
using System;

namespace AddOn_DocGenerator.SL
{
    public sealed class SapConnection
    {
        private static SAPbouiCOM.Application SBOApplication { get; set; }
        private static SAPbobsCOM.Company SBOCompany { get; set; }

        private SapConnection() { }

        public static SAPbouiCOM.Application GetApplication()
        {
            if (SBOApplication == null)
            {
                try
                {
                    SAPbouiCOM.SboGuiApi sboGuiApi = new SAPbouiCOM.SboGuiApi();

                    string connectionString;

                    #if DEBUG

                    // Token Manual (Dev)
                    connectionString = "0030002C0030002C00530041005000420044005F00440061007400650076002C0050004C006F006D0056004900490056";

                    #else

                    // Token Manual (Prod)
                    var args = Environment.GetCommandLineArgs();

                    if (args.Length < 2)
                        throw new Exception("No se recibió el connection string de SAP Business One.");

                    connectionString = args[1];

                    #endif

                    sboGuiApi.Connect(connectionString);

                    SBOApplication = sboGuiApi.GetApplication(-1);

                    if (SBOApplication == null)
                        throw new Exception("No se pudo obtener la aplicación SAP Business One.");

                    SBOApplication.StatusBar.SetText(
                        "AddOn DocGenerator conectado.",
                        BoMessageTime.bmt_Short,
                        BoStatusBarMessageType.smt_Success
                    );
                }
                catch (Exception ex)
                {
                    throw new Exception($"UI API (GetApplication): {ex.Message}", ex);
                }
            }

            return SBOApplication;
        }

        public static SAPbobsCOM.Company GetCompany()
        {
            if (SBOCompany == null)
            {
                try
                {
                    if (SBOApplication == null)
                        GetApplication();

                    SBOCompany = (SAPbobsCOM.Company)SBOApplication.Company.GetDICompany();

                    if (SBOCompany == null || !SBOCompany.Connected)
                        throw new Exception("No se pudo obtener la conexión DI API.");

                    SBOApplication.StatusBar.SetText(
                        $"Conectado a la sociedad: {SBOCompany.CompanyDB}",
                        BoMessageTime.bmt_Short,
                        BoStatusBarMessageType.smt_Success
                    );
                }
                catch (Exception ex)
                {
                    throw new Exception($"DI API (GetCompany): {ex.Message}", ex);
                }
            }

            return SBOCompany;
        }

        public static int GetServerType()
        {
            SAPbobsCOM.Company company = GetCompany();

            switch (company.DbServerType)
            {
                case BoDataServerTypes.dst_MSSQL:
                case BoDataServerTypes.dst_MSSQL2005:
                case BoDataServerTypes.dst_MSSQL2008:
                case BoDataServerTypes.dst_MSSQL2012:
                case BoDataServerTypes.dst_MSSQL2014:
                case BoDataServerTypes.dst_MSSQL2016:
                    return 0;

                case BoDataServerTypes.dst_HANADB:
                    return 1;

                default:
                    return -1;
            }
        }

        public static void Connect()
        {
            GetApplication();
            GetCompany();
        }
    }
}