using SAPbouiCOM.Framework;
using System;

namespace SBOAddonExplorers
{
    [FormAttribute("SBOAddonExplorers.frmExploradorProduccion", "frmExploradorProduccion.b1f")]
    class frmExploradorProduccion : UserFormBase
    {
        private SAPbouiCOM.Grid Grid0;
        private SAPbouiCOM.Button Button2;
        private SAPbouiCOM.ComboBox ComboBox0;
        private SAPbouiCOM.StaticText StaticText1;
        private SAPbouiCOM.EditText EditText0;
        public frmExploradorProduccion()
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
                    string query = "";

                    if (table.table.Equals("OWOR"))
                    {
                        query = string.Format("SELECT T1.\"DocEntry\", T1.\"DocNum\", (SELECT \"Descr\" FROM UFD1 WHERE \"FldValue\" = T1.\"U_Type\" AND \"FieldID\" = 0 AND \"TableID\" = 'OWOR') \"Type\", T1.\"ItemCode\", T1.\"Warehouse\", T0.\"ItemCode\", T0.\"wareHouse\", T1.\"Status\", T1.\"Type\", T1.\"PlannedQty\", T1.\"CmpltQty\", T1.\"PostDate\", T1.\"DueDate\" FROM {0} T1 JOIN {1} T0 ON T1.\"DocEntry\" = T0.\"DocEntry\"", table.table, table.detail);

                        if (search)
                            query = string.Format("SELECT T1.\"DocEntry\", T1.\"DocNum\", (SELECT \"Descr\" FROM UFD1 WHERE \"FldValue\" = T1.\"U_Type\" AND \"FieldID\" = 0 AND \"TableID\" = 'OWOR') \"Type\", T1.\"ItemCode\", T1.\"Warehouse\", T0.\"ItemCode\", T0.\"wareHouse\", T1.\"Status\", T1.\"Type\", T1.\"PlannedQty\", T1.\"CmpltQty\", T1.\"PostDate\", T1.\"DueDate\" FROM {0} T1 JOIN {1} T0 ON T1.\"DocEntry\" = T0.\"DocEntry\" WHERE T1.\"DocNum\" LIKE '{2}' OR T1.\"ItemCode\" LIKE '{2}' OR T0.\"ItemCode\" LIKE '{2}'", table.table, table.detail, EditText0.Value.ToString().Replace("*", "%"));

                        Grid0.DataTable.ExecuteQuery(query);

                        foreach (SAPbouiCOM.GridColumn column in Grid0.Columns)
                        {
                            column.Editable = false;
                            column.TitleObject.Sortable = true;
                        }

                        ((SAPbouiCOM.EditTextColumn)Grid0.Columns.Item("PlannedQty")).ColumnSetting.SumType = SAPbouiCOM.BoColumnSumType.bst_Auto;
                        ((SAPbouiCOM.EditTextColumn)Grid0.Columns.Item("CmpltQty")).ColumnSetting.SumType = SAPbouiCOM.BoColumnSumType.bst_Auto;
                        Grid0.CommonSetting.EnableArrowKey = true;
                        ((SAPbouiCOM.EditTextColumn)Grid0.Columns.Item("DocEntry")).LinkedObjectType = table.ObjType().ToString();
                        if (!search)
                            ((SAPbouiCOM.EditTextColumn)Grid0.Columns.Item("DocNum")).TitleObject.Sort(SAPbouiCOM.BoGridSortType.gst_Ascending);
                        ((SAPbouiCOM.EditTextColumn)Grid0.Columns.Item("Warehouse")).LinkedObjectType = ((SAPbouiCOM.EditTextColumn)Grid0.Columns.Item("wareHouse")).LinkedObjectType = "64";
                        ((SAPbouiCOM.EditTextColumn)Grid0.Columns.Item("ItemCode")).LinkedObjectType = "4";
                    } 
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