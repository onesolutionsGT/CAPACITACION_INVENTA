using SAPbouiCOM;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace EjemploCapacitacion
{
    class SystemForm
    {
        public SystemForm()
        {
            try
            {
                ApplicationContext.SetApplication();

                ApplicationContext
                    .SBOApplication
                    .StatusBar
                    .SetText("Inicializando Add-On capacitacion", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Success);
                //ApplicationContext.getTotalUltimaFactura();
                //ApplicationContext.AddFactura();
                SetMenuItems();
                SetEvents();
            }
            catch (Exception ex)
            {
                ApplicationContext
                   .SBOApplication
                   .StatusBar
                   .SetText(ApplicationContext.SBOError + " " + ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
                System.Environment.Exit(0);
            }
        }


        private void SetEvents()
        {
            ApplicationContext
                .SBOApplication
                .MenuEvent += new SAPbouiCOM._IApplicationEvents_MenuEventEventHandler(SBOApp_MenuEvent);

            ApplicationContext
                .SBOApplication
                .AppEvent += new SAPbouiCOM._IApplicationEvents_AppEventEventHandler(SBOApp_AppEvent);
        }

        private void SBOApp_AppEvent(BoAppEventTypes EventType)
        {
            if(EventType == SAPbouiCOM.BoAppEventTypes.aet_ShutDown || EventType == SAPbouiCOM.BoAppEventTypes.aet_ServerTerminition)
            {

                ApplicationContext
                    .SBOApplication
                    .StatusBar
                    .SetText("Cerrando Add-On capacitacion", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning);
                System.Environment.Exit(0);
            }
        }

        private void SBOApp_MenuEvent(ref MenuEvent pVal, out bool BubbleEvent)
        {
            BubbleEvent = true;
            try
            {
                if (!pVal.BeforeAction)
                {
                    ApplicationContext
                        .SBOApplication
                        .StatusBar
                        .SetText(pVal.MenuUID, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Success);
                }
            }
            catch (Exception ex)
            {
                ApplicationContext
                    .SBOApplication
                    .StatusBar
                    .SetText(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
            }
            finally
            {

            }
        }

        private void SetMenuItems()
        {
            //Instanciando arreglo de menus en SAP a objeto
            SAPbouiCOM.Menus menus = ApplicationContext.SBOApplication.Menus;
            SAPbouiCOM.MenuCreationParams creationParams = ApplicationContext.SBOApplication.CreateObject(SAPbouiCOM.BoCreatableObjectType.cot_MenuCreationParams);
            SAPbouiCOM.MenuItem menuItem = ApplicationContext.SBOApplication.Menus.Item("43520");//Arreglo de menus en pestaña "modulos"
            menus = menuItem.SubMenus;
            if (ApplicationContext.SBOApplication.Menus.Exists("PRUEBA_CAPACITACION"))
            {
                ApplicationContext.SBOApplication.Menus.RemoveEx("PRUEBA_CAPACITACION");
            }
            creationParams.Type = SAPbouiCOM.BoMenuType.mt_POPUP;
            creationParams.UniqueID = "PRUEBA_CAPACITACION";
            creationParams.String = ("prueba capactiacion");
            creationParams.Enabled = true;
            creationParams.Position = 2;
            //creationParmas.Image = "";
            menus.AddEx(creationParams);
            menuItem = ApplicationContext.SBOApplication.Menus.Item("PRUEBA_CAPACITACION");
            menus = menuItem.SubMenus;
            creationParams.Type = SAPbouiCOM.BoMenuType.mt_STRING;
            creationParams.UniqueID = "MENU_HIJO";
            creationParams.String = "menu hijo";
            menus.AddEx(creationParams);
        }
    }
}
