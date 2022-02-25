using SAPbouiCOM;
using System;
using System.Globalization;
using System.Runtime.InteropServices;

namespace SBOAddonExplorers
{
    class SAP
    {
        public Application app;
        public SAPbobsCOM.Company company;
        public bool usarHana = true;

        public SAP()
        {
            ConnectUIDI();
        }
        private int ConnectUI(string connectionString = "")
        {
            int returnValue = 0;
            //#if DEBUG
            if (string.IsNullOrEmpty(connectionString))
            {
                connectionString = "0030002C0030002C00530041005000420044005F00440061007400650076002C0050004C006F006D0056004900490056";
            }
            //#endif
            var sboGuiApi = new SboGuiApi();
            sboGuiApi.Connect(connectionString);
            try
            {
                app = sboGuiApi.GetApplication();
            }
            catch (Exception exception)
            {
                var message = string.Format(CultureInfo.InvariantCulture, "{0} Initialization - Error accessing SBO: {1}", "DB_TestConnection", exception.Message);
                Console.WriteLine(message);
                returnValue = -1;
            }
            Marshal.FinalReleaseComObject(sboGuiApi);
            return returnValue;
        }

        private int ConnectUIDI(string connectionString = "")
        {
            int returnValue = ConnectUI(connectionString);
            if (0 == returnValue)
            {
                try
                {
                    company = (SAPbobsCOM.Company)app.Company.GetDICompany();
                }
                catch (Exception exception)
                {
                    var message = string.Format(CultureInfo.InvariantCulture, "{0} Initialization - Error accessing SBO: {1}", "db_TestConnection", exception.Message);
                    Console.WriteLine(message);
                    returnValue = -1;
                }
            }
            return returnValue;
        }
    }
    class DocumentTable
    {
        public string table { get; set; }
        public string detail { get; set; }

        public DocumentTable(string table)
        {
            this.table = table;
            detail = table.Substring(1, table.Length - 1) + 1;
        }

        public int ObjType()
        {
            switch (table)
            {
                case "OQUT":
                    return 23;
                case "ORDR":
                    return 17;
                case "ODLN":
                    return 15;
                case "OINV":
                    return 13;
                case "OPQT":
                    return 540000006;
                case "OPOR":
                    return 22;
                case "OPDN":
                    return 20;
                case "OPCH":
                    return 18;
                case "OIGN":
                    return 59;
                case "OIGE":
                    return 60;
                case "OWTQ":
                    return 1250000001;
                case "OWTR":
                    return 67;
                case "OWOR":
                    return 202;
                default:
                    return 0;
            }
        }
    }
}
