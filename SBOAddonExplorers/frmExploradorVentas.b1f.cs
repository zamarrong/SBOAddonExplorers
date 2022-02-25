using SAPbouiCOM.Framework;
using System;
using System.Windows.Forms;

namespace SBOAddonExplorers
{
    [FormAttribute("SBOAddonExplorers.frmExploradorVentas", "frmExploradorVentas.b1f")]
    class frmExploradorVentas : UserFormBase
    {
        private SAPbouiCOM.Grid Grid0;
        private SAPbouiCOM.Button Button2;
        private SAPbouiCOM.ComboBox ComboBox0;
        private SAPbouiCOM.StaticText StaticText1;
        private SAPbouiCOM.EditText EditText0;
        private Timer timer = new Timer();
        public frmExploradorVentas()
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
            this.Button1 = ((SAPbouiCOM.Button)(this.GetItem("Item_4").Specific));
            this.Button1.ClickAfter += new SAPbouiCOM._IButtonEvents_ClickAfterEventHandler(this.Button1_ClickAfter);
            this.CheckBox0 = ((SAPbouiCOM.CheckBox)(this.GetItem("Item_7").Specific));

            CheckBox0.DataBind.SetBound(true, "", "UD_0");
            CheckBox0.ValOff = "N";
            CheckBox0.ValOn = "Y";

            CheckBox0.Checked = false;

            CheckBox0.PressedAfter += CheckBox0_ClickAfter;

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
            timer.Interval = 300000;
            timer.Tick += Timer_Tick;
            timer.Start();
        }

        private void CheckBox0_ClickAfter(object sboObject, SAPbouiCOM.SBOItemEventArg pVal)
        {
            try
            {
                setORDR();
            }
            catch (Exception ex)
            {
                Program.sap.app.MessageBox(ex.Message);
            }
        }

        private void Timer_Tick(object sender, EventArgs e)
        {
            SetDT(true);
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
                    string query = string.Format("SELECT T1.\"DocEntry\", T1.\"DocNum\",  T1.\"CardCode\", T1.\"CardName\", T0.\"LineNum\", T0.\"ItemCode\", T0.\"Dscription\", T0.\"Quantity\", T0.\"OpenQty\", T2.\"OnHand\", T0.\"WhsCode\", T0.\"U_DirectShipping\", T4.\"OnHand\" \"DSOnHand\", T1.\"DocDate\", T0.\"ShipDate\", DAYS_BETWEEN (T1.\"DocDate\", NOW()) \"Days\",  T0.\"LineTotal\", (T0.\"OpenQty\" * T0.\"Price\" * T1.\"DocRate\") \"OpenSum\", ((T0.\"OpenQty\" / T0.\"Quantity\") * 100) \"OpenPercent\", T0.\"FreeTxt\", T3.\"SlpName\" FROM {1} T0 INNER JOIN {0} T1 ON T0.\"DocEntry\" = T1.\"DocEntry\" INNER JOIN OITW T2 ON T0.\"ItemCode\" = T2.\"ItemCode\" AND T0.\"WhsCode\" = T2.\"WhsCode\" INNER JOIN OSLP T3 ON T3.\"SlpCode\" = T1.\"SlpCode\" LEFT JOIN OITW T4 ON T0.\"ItemCode\" = T4.\"ItemCode\" AND T0.\"U_DirectShipping\" = T4.\"WhsCode\" WHERE T0.\"OpenQty\" > 0", table.table, table.detail);
                    
                    if (search || EditText0.Value.Length > 0)
                        query = string.Format("SELECT T1.\"DocEntry\", T1.\"DocNum\", T1.\"CardCode\", T1.\"CardName\", T0.\"LineNum\", T0.\"ItemCode\", T0.\"Dscription\", T0.\"Quantity\", T0.\"OpenQty\", T2.\"OnHand\", T0.\"WhsCode\", T0.\"U_DirectShipping\", T4.\"OnHand\" \"DSOnHand\", T1.\"DocDate\", T0.\"ShipDate\", DAYS_BETWEEN (T1.\"DocDate\", NOW()) \"Days\", T0.\"LineTotal\", (T0.\"OpenQty\" * T0.\"Price\" * T1.\"DocRate\") \"OpenSum\", ((T0.\"OpenQty\" / T0.\"Quantity\") * 100) \"OpenPercent\", T0.\"FreeTxt\", T3.\"SlpName\" FROM {1} T0 INNER JOIN {0} T1 ON T0.\"DocEntry\" = T1.\"DocEntry\" INNER JOIN OITW T2 ON T0.\"ItemCode\" = T2.\"ItemCode\" AND T0.\"WhsCode\" = T2.\"WhsCode\" INNER JOIN OSLP T3 ON T3.\"SlpCode\" = T1.\"SlpCode\" LEFT JOIN OITW T4 ON T0.\"ItemCode\" = T4.\"ItemCode\" AND T0.\"U_DirectShipping\" = T4.\"WhsCode\" WHERE T0.\"OpenQty\" > 0 AND (T1.\"DocNum\" LIKE '{2}' OR T1.\"CardCode\" LIKE '{2}' OR T1.\"CardName\" LIKE '{2}' OR T0.\"ItemCode\" LIKE '{2}' OR T0.\"WhsCode\" LIKE '{2}')", table.table, table.detail, EditText0.Value.ToString().Replace("*", "%"));

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
                    ((SAPbouiCOM.EditTextColumn)Grid0.Columns.Item("OnHand")).DoubleClickAfter += FrmExploradorVentas_DoubleClickAfter;
                    ((SAPbouiCOM.EditTextColumn)Grid0.Columns.Item("U_DirectShipping")).LinkedObjectType = "64";

                    if (table.table.Equals("ORDR"))
                    {
                        CheckBox0.Item.Visible = true;
                        setORDR();
                    }
                    else
                    {
                        CheckBox0.Item.Visible = false;
                        CheckBox0.Checked = false;
                        Grid0.Columns.Item("U_DirectShipping").Editable = false;
                        Grid0.Columns.Item("U_DirectShipping").Type = SAPbouiCOM.BoGridColumnType.gct_EditText;
                    }
                }
            }
            catch (Exception ex)
            {
                Program.sap.app.MessageBox(ex.Message);
            }
        }

        private void FrmExploradorVentas_DoubleClickAfter(object sboObject, SAPbouiCOM.SBOItemEventArg pVal)
        {
            string ItemCode = (string)Grid0.DataTable.GetValue("ItemCode", Grid0.GetDataTableRowIndex(pVal.Row));
            frmExploradorInventario activeForm = new frmExploradorInventario(ItemCode);
            activeForm.Show();
        }

        private void setORDR()
        {
            try
            {
                if (CheckBox0.Checked)
                {
                    Grid0.Columns.Item("U_DirectShipping").Editable = true;
                    Grid0.Columns.Item("U_DirectShipping").Type = SAPbouiCOM.BoGridColumnType.gct_ComboBox;

                    if (((SAPbouiCOM.ComboBoxColumn)Grid0.Columns.Item("U_DirectShipping")).ValidValues.Count == 0)
                    {
                        SAPbobsCOM.Recordset rs = (SAPbobsCOM.Recordset)Program.sap.company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
                        rs.DoQuery("SELECT \"WhsCode\", \"WhsName\" FROM OWHS");

                        ((SAPbouiCOM.ComboBoxColumn)Grid0.Columns.Item("U_DirectShipping")).ValidValues.Add("", "- Ningún almacén -");

                        for (int i = 0; i < rs.RecordCount; i++)
                        {
                            ((SAPbouiCOM.ComboBoxColumn)Grid0.Columns.Item("U_DirectShipping")).ValidValues.Add(rs.Fields.Item("WhsCode").Value.ToString(), rs.Fields.Item("WhsName").Value.ToString());
                            rs.MoveNext();
                        }

                        Grid0.Columns.Item("U_DirectShipping").ComboSelectAfter += FrmExploradorVentas_ComboSelectAfter;
                    }
                }
                else
                {
                    foreach (SAPbouiCOM.GridColumn column in Grid0.Columns)
                    {
                        column.Editable = false;
                        column.TitleObject.Sortable = true;
                    }

                    Grid0.Columns.Item("U_DirectShipping").Type = SAPbouiCOM.BoGridColumnType.gct_EditText;
                }
            }
            catch
            {
                SetDT();
            }
        }

        private void FrmExploradorVentas_ComboSelectAfter(object sboObject, SAPbouiCOM.SBOItemEventArg pVal)
        {
            SAPbobsCOM.Recordset rs = (SAPbobsCOM.Recordset)Program.sap.company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
            SAPbobsCOM.Documents docto = (SAPbobsCOM.Documents)Program.sap.company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oOrders);

            int DocEntry = (int)Grid0.DataTable.GetValue("DocEntry", Grid0.GetDataTableRowIndex(pVal.Row));
            int LineNum = (int)Grid0.DataTable.GetValue("LineNum", Grid0.GetDataTableRowIndex(pVal.Row));
            string ItemCode = (string)Grid0.DataTable.GetValue("ItemCode", Grid0.GetDataTableRowIndex(pVal.Row));
            double OpenQty = (double)Grid0.DataTable.GetValue("OpenQty", Grid0.GetDataTableRowIndex(pVal.Row));
            string DirectShipping = (string)Grid0.DataTable.GetValue("U_DirectShipping", Grid0.GetDataTableRowIndex(pVal.Row));

            if (OpenQty > 0)
            {
                docto.GetByKey(DocEntry);

                docto.Lines.SetCurrentLine(LineNum);

                if (docto.Lines.LineStatus == SAPbobsCOM.BoStatus.bost_Open)
                {
                    if (docto.Lines.ItemCode == ItemCode)
                    {
                        docto.Lines.UserFields.Fields.Item("U_DirectShipping").Value = DirectShipping;

                        if (docto.Update() == 0)
                        {
                            Program.sap.app.StatusBar.SetText("Se actualizó el almacén correctamente.", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Success);
                        }
                        else
                        {
                            Program.sap.app.StatusBar.SetText(Program.sap.company.GetLastErrorDescription(), SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
                        }
                    }
                }
                else
                {
                    Program.sap.app.StatusBar.SetText("No es posible modificar la linea solicitada.", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning);
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

        private SAPbouiCOM.Button Button1;

        private void Button1_ClickAfter(object sboObject, SAPbouiCOM.SBOItemEventArg pVal)
        {
            if (Grid0.Rows.SelectedRows.Count > 0)
            {
                string ItemCode = (string)Grid0.DataTable.GetValue("ItemCode", Grid0.Rows.SelectedRows.Item(0, SAPbouiCOM.BoOrderType.ot_SelectionOrder));
                frmExploradorInventario activeForm = new frmExploradorInventario(ItemCode);
                activeForm.Show();
            }
        }

        private SAPbouiCOM.CheckBox CheckBox0;
    }
}