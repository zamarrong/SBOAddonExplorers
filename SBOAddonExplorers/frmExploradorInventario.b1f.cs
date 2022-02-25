using SAPbouiCOM.Framework;
using System;
using System.Windows.Forms;

namespace SBOAddonExplorers
{
    [FormAttribute("SBOAddonExplorers.frmExploradorInventario", "frmExploradorInventario.b1f")]
    class frmExploradorInventario : UserFormBase
    {
        private SAPbouiCOM.Grid Grid0;
        private SAPbouiCOM.Button Button2;
        private SAPbouiCOM.ComboBox ComboBox0;
        private SAPbouiCOM.StaticText StaticText1;
        private SAPbouiCOM.EditText EditText0;
        private Timer timer = new Timer();
        public frmExploradorInventario(string ItemCode = null)
        {
            if (ItemCode != null)
            {
                EditText0.Value = ItemCode;
                ComboBox0.Select("OITW");
            }
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
            timer.Interval = 300000;
            timer.Tick += Timer_Tick;
            timer.Start();
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
                    string query = "";
                    if (table.table.Equals("OITW"))
                    {
                        if (search || EditText0.Value.Length > 0)
                            query = string.Format("SELECT OITW.\"ItemCode\", OITM.\"ItemName\", OITW.\"WhsCode\", OITW.\"OnHand\", OITW.\"IsCommited\" FROM OITW INNER JOIN OITM ON OITW.\"ItemCode\" = OITM.\"ItemCode\" WHERE OITW.\"ItemCode\" LIKE '{0}' OR OITM.\"ItemName\" LIKE '{0}'", EditText0.Value.ToString().Replace("*", "%"));
                        else
                            query = "SELECT OITW.\"ItemCode\", OITM.\"ItemName\", OITW.\"WhsCode\", OITW.\"OnHand\", OITW.\"IsCommited\" FROM OITW INNER JOIN OITM ON OITW.\"ItemCode\" = OITM.\"ItemCode\"";

                        Grid0.DataTable.ExecuteQuery(query);


                        Grid0.CommonSetting.EnableArrowKey = true;

                        foreach (SAPbouiCOM.GridColumn column in Grid0.Columns)
                        {
                            column.Editable = false;
                            column.TitleObject.Sortable = true;
                        }
                    }
                    else
                    {
                        query = string.Format("SELECT T1.\"DocEntry\", T1.\"DocNum\", (SELECT \"Descr\" FROM UFD1 WHERE \"FldValue\" = T1.\"U_Type\" AND \"FieldID\" = 6 AND \"TableID\" = 'ADOC') \"Type\", T1.\"CardCode\", T1.\"CardName\", T0.\"ItemCode\", T0.\"Dscription\", T0.\"Quantity\", T0.\"OpenQty\", T2.\"OnHand\", T0.\"FromWhsCod\", T0.\"WhsCode\", T1.\"U_FinalDestWhs\", T1.\"DocDate\", T0.\"ShipDate\", DAYS_BETWEEN (T1.\"DocDate\", NOW()) \"Days\", T1.\"Comments\" FROM {1} T0 INNER JOIN {0} T1 ON T0.\"DocEntry\" = T1.\"DocEntry\" INNER JOIN OITW T2 ON T0.\"ItemCode\" = T2.\"ItemCode\" AND T0.\"FromWhsCod\" = T2.\"WhsCode\" WHERE T0.\"OpenQty\" > 0", table.table, table.detail);

                        if (search)
                            query = string.Format("SELECT T1.\"DocEntry\", T1.\"DocNum\", (SELECT \"Descr\" FROM UFD1 WHERE \"FldValue\" = T1.\"U_Type\" AND \"FieldID\" = 6 AND \"TableID\" = 'ADOC') \"Type\", T1.\"CardCode\", T1.\"CardName\", T0.\"ItemCode\", T0.\"Dscription\", T0.\"Quantity\", T0.\"OpenQty\",T2.\"OnHand\", T0.\"FromWhsCod\", T0.\"WhsCode\", T1.\"U_FinalDestWhs\", T1.\"DocDate\", T0.\"ShipDate\", DAYS_BETWEEN (T1.\"DocDate\", NOW()) \"Days\", T1.\"Comments\" FROM {1} T0 INNER JOIN {0} T1 ON T0.\"DocEntry\" = T1.\"DocEntry\" INNER JOIN OITW T2 ON T0.\"ItemCode\" = T2.\"ItemCode\" AND T0.\"FromWhsCod\" = T2.\"WhsCode\" WHERE T0.\"OpenQty\" > 0 AND (T1.\"DocNum\" LIKE '{2}' OR T1.\"CardCode\" LIKE '{2}' OR T1.\"CardName\" LIKE '{2}' OR T0.\"ItemCode\" LIKE '{2}' OR T0.\"WhsCode\" LIKE '{2}' OR T0.\"FromWhsCod\" LIKE '{2}')", table.table, table.detail, EditText0.Value.ToString().Replace("*", "%"));

                        Grid0.DataTable.ExecuteQuery(query);

                        Grid0.CommonSetting.EnableArrowKey = true;

                        foreach (SAPbouiCOM.GridColumn column in Grid0.Columns)
                        {
                            column.Editable = false;
                            column.TitleObject.Sortable = true;
                        }

                        ((SAPbouiCOM.EditTextColumn)Grid0.Columns.Item("OpenQty")).ColumnSetting.SumType = SAPbouiCOM.BoColumnSumType.bst_Auto;
                       
                        ((SAPbouiCOM.EditTextColumn)Grid0.Columns.Item("DocEntry")).LinkedObjectType = table.ObjType().ToString();
                        if (!search)
                            ((SAPbouiCOM.EditTextColumn)Grid0.Columns.Item("DocNum")).TitleObject.Sort(SAPbouiCOM.BoGridSortType.gst_Ascending);
                        ((SAPbouiCOM.EditTextColumn)Grid0.Columns.Item("CardCode")).LinkedObjectType = "2";
                    }

                    ((SAPbouiCOM.EditTextColumn)Grid0.Columns.Item("ItemCode")).LinkedObjectType = "4";
                    ((SAPbouiCOM.EditTextColumn)Grid0.Columns.Item("WhsCode")).LinkedObjectType = "64";

                    if (table.table != "OITW")
                        ((SAPbouiCOM.EditTextColumn)Grid0.Columns.Item("U_FinalDestWhs")).LinkedObjectType = ((SAPbouiCOM.EditTextColumn)Grid0.Columns.Item("FromWhsCod")).LinkedObjectType = "64";
                }
            }
            catch (Exception ex)
            {
                Program.sap.app.MessageBox(ex.Message);
            }
        }

        private SAPbouiCOM.Button Button0;

        private void Button0_ClickAfter(object sboObject, SAPbouiCOM.SBOItemEventArg pVal)
        {
            SetDT();
        }
    }
}