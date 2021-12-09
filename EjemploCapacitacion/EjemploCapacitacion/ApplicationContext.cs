using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using SAPbobsCOM;
using SAPbouiCOM;
namespace EjemploCapacitacion
{
    class ApplicationContext
    {
        //String de conexión a SAP
        public static string ConnectionString { get; set; }

        //Instancia de la sociedad en SAP (DI API)
        public static SAPbobsCOM.Company SBOCompany { get; set; }

        //Instancia de la aplicacion SAP (UI API)
        public static SAPbouiCOM.Application SBOApplication { get; set; }

        //Instancia de el último error de SAP

        public static string SBOError { get { return "ERROR: (" + SBOCompany.GetLastErrorCode() + ") ->" + SBOCompany.GetLastErrorDescription(); } }
        //ERROR: (codigo de error) -> Descripcion

        public static void SetApplication()
        {
            //CONEXION UI API
            //String de conexión
            ConnectionString = "0030002C0030002C00530041005000420044005F00440061007400650076002C0050004C006F006D0056004900490056";
            //Instanciando un cliente para conectar a SAP
            SAPbouiCOM.SboGuiApi client = new SAPbouiCOM.SboGuiApi();
            client.Connect(ConnectionString);
            //Asignacion de aplicacion
            SBOApplication = client.GetApplication(-1);
            //Conexión DI API
            SBOCompany = new SAPbobsCOM.Company();//Instancia
            string coockies = SBOCompany.GetContextCookie();//Coockies de sesión
            string connectionContext = SBOApplication.Company.GetConnectionContext(coockies);//Contexto de inicio de sesion
            SBOCompany.SetSboLoginContext(connectionContext);//LogIn DI API
            if (SBOCompany.Connect() != 0) { throw new Exception(SBOError); }
        }

        public static void AddFactura()
        {
            //Llamado a la DI API para que inice una transacción en la DB
            SBOCompany.StartTransaction();
            //Instanciamos el objeto (documento)
            Documents nuevaFactura = SBOCompany.GetBusinessObject(BoObjectTypes.oInvoices);
            nuevaFactura.CardCode = "POS001";//Socio de negocios
            nuevaFactura.DocDate = DateTime.Now;//Fecha del documento
            nuevaFactura.TaxDate = DateTime.Now;//Fecha de contabilizacion
            //Para consultar cuantos dias tiene el cliente para pagar la factura, se debe consultar a la base de datos.
            //select "ExtraDays", "ExtraMonth" from OCRD inner join OCTG on OCRD."GroupNum" = OCTG."GroupNum" where "CardCode" = 'POS001';
            Recordset rs = SBOCompany.GetBusinessObject(BoObjectTypes.BoRecordset);
            string query = "select \"ExtraDays\", \"ExtraMonth\" from OCRD inner join OCTG on OCRD.\"GroupNum\" = OCTG.\"GroupNum\" where \"CardCode\" = 'POS001'";
            rs.DoQuery(query);
            Calendar calendar = CultureInfo.InvariantCulture.Calendar;
            int extra_days = rs.Fields.Item("ExtraDays").Value;
            int extra_months = rs.Fields.Item("ExtraMonth").Value;
            DateTime fecha_pago = DateTime.Now;
            fecha_pago = calendar.AddDays(fecha_pago, extra_days);
            fecha_pago = calendar.AddMonths(fecha_pago, extra_months);
            nuevaFactura.DocDueDate = fecha_pago;//Fecha de pago
            nuevaFactura.Series = 72;//Serie
            nuevaFactura.DocCurrency = "QTZ";//Moneda
            nuevaFactura.DocType = SAPbobsCOM.BoDocumentTypes.dDocument_Items;//Tipo de factura

            //Líneas de la factura (detalle)
            nuevaFactura.Lines.ItemCode = "100002";//Item
            nuevaFactura.Lines.AccountCode = "410101001";//Cuenta contable
            nuevaFactura.Lines.TaxCode = "IVA";//Impuesto
            nuevaFactura.Lines.WTLiable = SAPbobsCOM.BoYesNoEnum.tYES;//¿Aplicar retenciond e impuestos?
            nuevaFactura.Lines.Add();//Agregar linea

            nuevaFactura.Lines.ItemCode = "100003";
            nuevaFactura.Lines.AccountCode = "410101001";
            nuevaFactura.Lines.TaxCode = "IVA";
            nuevaFactura.Lines.WTLiable = SAPbobsCOM.BoYesNoEnum.tYES;
            nuevaFactura.Lines.Add();

            int resultado = nuevaFactura.Add();
            if (resultado != 0)
            {
                SBOApplication
                .StatusBar
                .SetText(SBOError, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
                SBOCompany.EndTransaction(BoWfTransOpt.wf_RollBack);//Si hay error, se reporta y se eliminan los cambios realizados.
            }
            else
            {
                SBOApplication
                    .StatusBar
                    .SetText("Factura ingresada correctamente", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Success);
                SBOCompany.EndTransaction(BoWfTransOpt.wf_Commit);//Transaccion realizada con exito, se confirman cambios

                SBOCompany.GetNewObjectCode(out string DocEntry);

                Documents buscarFactura = SBOCompany.GetBusinessObject(BoObjectTypes.oInvoices);


                if (buscarFactura.GetByKey(Convert.ToInt32(DocEntry)) == true)
                {
                    string mensaje = "FACTURA ENCONTRADA: " + DocEntry;
                    SBOApplication
                     .StatusBar
                     .SetText(mensaje, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Success);


                    SBOCompany.StartTransaction();

                    buscarFactura.Comments = "COMENTARIO DE MODIFICACION";

                    int res = buscarFactura.Update();
                    if (res != 0)
                    {
                        SBOApplication
                            .StatusBar
                            .SetText(SBOError, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
                        SBOCompany.EndTransaction(BoWfTransOpt.wf_RollBack);//Si hay error, se reporta y se eliminan los cambios realizados.
                    }
                    else
                    {
                        SBOApplication
                            .StatusBar
                            .SetText("Factura actualizada correctamente DocEntry: " + buscarFactura.DocEntry.ToString(), SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Success);
                        SBOCompany.EndTransaction(BoWfTransOpt.wf_Commit);//Transaccion realizada con exito, se confirman cambios
                    }
                }
                else
                {
                    SBOApplication
                       .StatusBar
                       .SetText(SBOError, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
                }
            }
        }

        public static void getTotalUltimaFactura()
        {
            //Buscar top 10 de facturas (total) y desplegar la informacion del 10mo lugar
            //select "DocEntry", "Series", "CardCode", "DocTotal" from OINV order by "DocTotal" desc limit 10;
            Recordset rs = SBOCompany.GetBusinessObject(BoObjectTypes.BoRecordset);
            string query = "SELECT \"DocEntry\", \"Series\", \"CardCode\", \"DocTotal\" from OINV order by \"DocTotal\" desc limit 10";
            rs.DoQuery(query);
            //EoF -> End of File
            //BoF -> Begin of File

            while (rs.EoF == false)
            {
                //Ir al siguiente registro
                rs.MoveNext();
            }
            //Regresar un espacio
            if (rs.EoF == true)
            {
           //     rs.MovePrevious();
            }


            //rs.MoveLast();
            //rs.MoveFirst();
            string mensaje = "FACTURA ENCONTRADA: DocEntry: " + rs.Fields.Item("DocEntry").Value +
            ", Serie: " + rs.Fields.Item("Series").Value + ", CardCode: " + rs.Fields.Item("CardCode").Value +
            ", Total: " + rs.Fields.Item("DocTotal").Value;
            SBOApplication
             .StatusBar
             .SetText(mensaje, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Success);

        }
    }
}
