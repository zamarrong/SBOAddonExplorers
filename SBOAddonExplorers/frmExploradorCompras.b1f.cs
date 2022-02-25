using SAPbouiCOM.Framework;
using System;

namespace SBOAddonExplorers
{
    [FormAttribute("SBOAddonExplorers.frmExploradorCompras", "frmExploradorCompras.b1f")]
    class frmExploradorCompras : UserFormBase
    {
        private SAPbouiCOM.Grid Grid0;
        private SAPbouiCOM.Button Button2;
        private SAPbouiCOM.ComboBox ComboBox0;
        private SAPbouiCOM.StaticText StaticText1;
        private SAPbouiCOM.EditText EditText0;
        public frmExploradorCompras()
        {

        }

        /// <summary>
        /// Initialize components. Called by framework after form created.
        /// </summary>
        public override void OnInitializeComponent()
        {
            this.Grid0 = ((SAPbouiCOM.Grid)(this.GetItem("Item_0").Specific));
            this.ComboBox0 = ((SAPbouiCOM.ComboBox)(this.GetItem("Item_5").Specific));
            this.ComboBox0.ComboSelectAfter += new SAPbouiCOM._IComboBoxEvents_ComboSelectAfterEventHandler(this.ComboBox0_ComboSelectAfter);
            this.Button2 = ((SAPbouiCOM.Button)(this.GetItem("Item_6").Specific));
            this.Button2.ClickAfter += new SAPbouiCOM._IButtonEvents_ClickAfterEventHandler(this.Button2_ClickAfter);
            this.StaticText1 = ((SAPbouiCOM.StaticText)(this.GetItem("Item_2").Specific));
            this.EditText0 = ((SAPbouiCOM.EditText)(this.GetItem("Item_3").Specific));
            this.EditText0.KeyDownAfter += new SAPbouiCOM._IEditTextEvents_KeyDownAfterEventHandler(this.EditText0_KeyDownAfter);
            this.Button0 = ((SAPbouiCOM.Button)(this.GetItem("Item_1").Specific));
            this.Button0.ClickAfter += new SAPbouiCOM._IButtonEvents_ClickAfterEventHandler(this.Button0_ClickAfter);
            this.OnCustomInitialize();

        }

        /// <summary>
        /// Initialize form event. Called by framework before form creation.
        /// </summary>
        public override void OnInitializeFormEvents()
        {
            
        }

        private void OnCustomInitialize()
        {

        }

        private void Button2_ClickAfter(object sboObject, SAPbouiCOM.SBOItemEventArg pVal)
        {
            Program.sap.app.Forms.ActiveForm.Close();
        }

        private void EditText0_KeyDownAfter(object sboObject, SAPbouiCOM.SBOItemEventArg pVal)
        {
            if (pVal.CharPressed == 13)
                SetDT(true);
        }

        private void ComboBox0_ComboSelectAfter(object sboObject, SAPbouiCOM.SBOItemEventArg pVal)
        {
            SetDT();
        }

        private void SetDT(bool search = false)
        {
            try
            {
                if (ComboBox0.Selected != null)
                {
                    DocumentTable table = new DocumentTable(ComboBox0.Selected.Value);
                    string query = string.Format("SELECT T1.\"DocEntry\", T1.\"DocNum\", (SELECT \"Descr\" FROM UFD1 WHERE \"FldValue\" = T1.\"U_Type\" AND \"FieldID\" = 6 AND \"TableID\" = 'ADOC') \"Type\", T1.\"CardCode\", T1.\"CardName\", T0.\"U_FollowUp\", T0.\"ItemCode\", T0.\"Dscription\", T0.\"Quantity\", T0.\"OpenQty\", T0.\"WhsCode\", T2.\"OnHand\", T1.\"DocDate\", T0.\"ShipDate\", DAYS_BETWEEN (T1.\"DocDate\", NOW()) \"Days\",  T0.\"LineTotal\", (T0.\"OpenQty\" * T0.\"Price\" * T1.\"DocRate\") \"OpenSum\", ((T0.\"OpenQty\" / T0.\"Quantity\") * 100) \"OpenPercent\", T0.\"FreeTxt\" FROM {1} T0 INNER JOIN {0} T1 ON T0.\"DocEntry\" = T1.\"DocEntry\" INNER JOIN OITW T2 ON T0.\"ItemCode\" = T2.\"ItemCode\" AND T0.\"WhsCode\" = T2.\"WhsCode\" WHERE T0.\"OpenQty\" > 0", table.table, table.detail);

                    if (search)
                        query = string.Format("SELECT T1.\"DocEntry\", T1.\"DocNum\", (SELECT \"Descr\" FROM UFD1 WHERE \"FldValue\" = T1.\"U_Type\" AND \"FieldID\" = 6 AND \"TableID\" = 'ADOC') \"Type\", T1.\"CardCode\", T1.\"CardName\", T0.\"U_FollowUp\", T0.\"ItemCode\", T0.\"Dscription\", T0.\"Quantity\", T0.\"OpenQty\", T0.\"WhsCode\", T2.\"OnHand\", T1.\"DocDate\", T0.\"ShipDate\", DAYS_BETWEEN (T1.\"DocDate\", NOW()) \"Days\", T0.\"LineTotal\", (T0.\"OpenQty\" * T0.\"Price\" * T1.\"DocRate\") \"OpenSum\", ((T0.\"OpenQty\" / T0.\"Quantity\") * 100) \"OpenPercent\", T0.\"FreeTxt\" FROM {1} T0 INNER JOIN {0} T1 ON T0.\"DocEntry\" = T1.\"DocEntry\" INNER JOIN OITW T2 ON T0.\"ItemCode\" = T2.\"ItemCode\" AND T0.\"WhsCode\" = T2.\"WhsCode\"  WHERE T0.\"OpenQty\" > 0 AND (T1.\"DocNum\" LIKE '{2}' OR T1.\"CardCode\" LIKE '{2}' OR T1.\"CardName\" LIKE '{2}' OR T0.\"ItemCode\" LIKE '{2}' OR T0.\"WhsCode\" LIKE '{2}')", table.table, table.detail, EditText0.Value.ToString().Replace("*", "%"));

                    Grid0.DataTable.ExecuteQuery(query);

                    foreach (SAPbouiCOM.GridColumn column in Grid0.Columns)
                    {
                        column.Editable = false;
                        column.TitleObject.Sortable = true;
                    }
                    ((SAPbouiCOM.EditTextColumn)Grid0.Columns.Item("OpenQty")).ColumnSetting.SumType = SAPbouiCOM.BoColumnSumType.bst_Auto;
                    ((SAPbouiCOM.EditTextColumn)Grid0.Columns.Item("OpenSum")).ColumnSetting.SumType = SAPbouiCOM.BoColumnSumType.bst_Auto;
                    Grid0.CommonSetting.EnableArrowKey = true;
                    ((SAPbouiCOM.EditTextColumn)Grid0.Columns.Item("DocEntry")).LinkedObjectType = table.ObjType().ToString();
                    if (!search)
                        ((SAPbouiCOM.EditTextColumn)Grid0.Columns.Item("DocNum")).TitleObject.Sort(SAPbouiCOM.BoGridSortType.gst_Ascending);
                    ((SAPbouiCOM.EditTextColumn)Grid0.Columns.Item("CardCode")).LinkedObjectType = "2";
                    ((SAPbouiCOM.EditTextColumn)Grid0.Columns.Item("ItemCode")).LinkedObjectType = "4";
                    ((SAPbouiCOM.EditTextColumn)Grid0.Columns.Item("WhsCode")).LinkedObjectType = "64";

                    if (table.table.Equals("OPOR"))
                        setOPOR(table);
                }
            }
            catch (Exception ex)
            {
                Program.sap.app.MessageBox(ex.Message);
            }
        }

        private void setOPOR(DocumentTable table)
        {
            Grid0.Columns.Item("U_FollowUp").Editable = true;
            Grid0.Columns.Item("U_FollowUp").Type = SAPbouiCOM.BoGridColumnType.gct_ComboBox;

            if (((SAPbouiCOM.ComboBoxColumn)Grid0.Columns.Item("U_FollowUp")).ValidValues.Count == 0)
            {
                SAPbobsCOM.Recordset rs = (SAPbobsCOM.Recordset)Program.sap.company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
                rs.DoQuery("SELECT \"FldValue\", \"Descr\" FROM UFD1 WHERE \"TableID\" = 'ADO1'");

                for (int i = 0; i < rs.RecordCount; i++)
                {
                    ((SAPbouiCOM.ComboBoxColumn)Grid0.Columns.Item("U_FollowUp")).ValidValues.Add(rs.Fields.Item("FldValue").Value.ToString(), rs.Fields.Item("Descr").Value.ToString());
                    rs.MoveNext();
                }

                //CheckBox0.Checked = true;
                Grid0.Columns.Item("U_FollowUp").ComboSelectAfter += FrmExploradorCompras_ComboSelectAfter;
            }
        }

        private void FrmExploradorCompras_ComboSelectAfter(object sboObject, SAPbouiCOM.SBOItemEventArg pVal)
        {
            SAPbobsCOM.Recordset rs = (SAPbobsCOM.Recordset)Program.sap.company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
            SAPbobsCOM.Documents docto = (SAPbobsCOM.Documents)Program.sap.company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oPurchaseOrders);

            int DocEntry = (int)Grid0.DataTable.GetValue("DocEntry", Grid0.GetDataTableRowIndex(pVal.Row));
            string ItemCode = (string)Grid0.DataTable.GetValue("ItemCode", Grid0.GetDataTableRowIndex(pVal.Row));
            double OpenQty = (double)Grid0.DataTable.GetValue("OpenQty", Grid0.GetDataTableRowIndex(pVal.Row));
            string FollowUp = (string)Grid0.DataTable.GetValue("U_FollowUp", Grid0.GetDataTableRowIndex(pVal.Row));

            if (OpenQty > 0)
            {
                rs.DoQuery(string.Format("SELECT TOP 1 \"LineNum\" FROM {0} WHERE \"DocEntry\" = {1} AND \"ItemCode\" = '{2}'", "RDR1", DocEntry, ItemCode));

                docto.GetByKey(DocEntry);

                docto.Lines.SetCurrentLine((int)rs.Fields.Item("LineNum").Value);
                docto.Lines.UserFields.Fields.Item("U_FollowUp").Value = FollowUp;

                if (docto.Update() == 0)
                {
                    Program.sap.app.StatusBar.SetText("Se actualizó el follow up correctamente.", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Success);
                }
                else
                {
                    Program.sap.app.StatusBar.SetText(Program.sap.company.GetLastErrorDescription(), SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
                }
            }
            else
            {
                Program.sap.app.MessageBox("No es posible modificar esa línea ya que la cantidad abierta es menor o igual a cero.");
            }
        }

        private SAPbouiCOM.Button Button0;

        private void Button0_ClickAfter(object sboObject, SAPbouiCOM.SBOItemEventArg pVal)
        {
            SetDT();
        }
    }
}