using SAPbouiCOM.Framework;
using System;

namespace SBOAddonExplorers
{
    class Menu
    {
        public void AddMenuItems()
        {
            SAPbouiCOM.Menus oMenus = null;
            SAPbouiCOM.MenuItem oMenuItem = null;

            oMenus = Application.SBO_Application.Menus;

            SAPbouiCOM.MenuCreationParams oCreationPackage = null;
            oCreationPackage = ((SAPbouiCOM.MenuCreationParams)(Application.SBO_Application.CreateObject(SAPbouiCOM.BoCreatableObjectType.cot_MenuCreationParams)));

            try
            {
                //Agrega producción
                oMenuItem = Application.SBO_Application.Menus.Item("43542");
                oMenus = oMenuItem.SubMenus;

                oCreationPackage.Type = SAPbouiCOM.BoMenuType.mt_STRING;
                oCreationPackage.UniqueID = "SBOAddonExplorers.frmExploradorProduccion";
                oCreationPackage.String = "Explorador de producción";
                oMenus.AddEx(oCreationPackage);

                //Agrega ventas
                oMenuItem = Application.SBO_Application.Menus.Item("12800");
                oMenus = oMenuItem.SubMenus;

                oCreationPackage.Type = SAPbouiCOM.BoMenuType.mt_STRING;
                oCreationPackage.UniqueID = "SBOAddonExplorers.frmExploradorVentas";
                oCreationPackage.String = "Explorador de ventas";
                oMenus.AddEx(oCreationPackage);

                //Agrega compras
                oMenuItem = Application.SBO_Application.Menus.Item("43534");
                oMenus = oMenuItem.SubMenus;

                oCreationPackage.Type = SAPbouiCOM.BoMenuType.mt_STRING;
                oCreationPackage.UniqueID = "SBOAddonExplorers.frmExploradorCompras";
                oCreationPackage.String = "Explorador de compras";
                oMenus.AddEx(oCreationPackage);

                //Agrega inventario
                oMenuItem = Application.SBO_Application.Menus.Item("1760");
                oMenus = oMenuItem.SubMenus;

                oCreationPackage.Type = SAPbouiCOM.BoMenuType.mt_STRING;
                oCreationPackage.UniqueID = "SBOAddonExplorers.frmExploradorInventario";
                oCreationPackage.String = "Explorador de inventario";
                oMenus.AddEx(oCreationPackage);

                //Agrega socios
                oMenuItem = Application.SBO_Application.Menus.Item("43536");
                oMenus = oMenuItem.SubMenus;

                oCreationPackage.Type = SAPbouiCOM.BoMenuType.mt_STRING;
                oCreationPackage.UniqueID = "SBOAddonExplorers.frmExploradorSocios";
                oCreationPackage.String = "Explorador de socios";
                oMenus.AddEx(oCreationPackage);
            }
            catch (Exception er)
            { //  Menu already exists
                Application.SBO_Application.SetStatusBarMessage("Menu Already Exists", SAPbouiCOM.BoMessageTime.bmt_Short, true);
            }
        }

        public void SBO_Application_MenuEvent(ref SAPbouiCOM.MenuEvent pVal, out bool BubbleEvent)
        {
            BubbleEvent = true;

            try
            {
                if (pVal.BeforeAction && pVal.MenuUID == "SBOAddonExplorers.frmExploradorVentas")
                {
                    frmExploradorVentas activeForm = new frmExploradorVentas();
                    activeForm.Show();
                }
                else if (pVal.BeforeAction && pVal.MenuUID == "SBOAddonExplorers.frmExploradorCompras")
                {
                    frmExploradorCompras activeForm = new frmExploradorCompras();
                    activeForm.Show();
                }
                else if (pVal.BeforeAction && pVal.MenuUID == "SBOAddonExplorers.frmExploradorInventario")
                {
                    frmExploradorInventario activeForm = new frmExploradorInventario();
                    activeForm.Show();
                }
                else if (pVal.BeforeAction && pVal.MenuUID == "SBOAddonExplorers.frmExploradorSocios")
                {
                    frmExploradorSocios activeForm = new frmExploradorSocios();
                    activeForm.Show();
                }
                else if (pVal.BeforeAction && pVal.MenuUID == "SBOAddonExplorers.frmExploradorProduccion")
                {
                    frmExploradorProduccion activeForm = new frmExploradorProduccion();
                    activeForm.Show();
                }
            }
            catch (Exception ex)
            {
                Application.SBO_Application.MessageBox(ex.ToString(), 1, "Ok", "", "");
            }
        }

    }
}
