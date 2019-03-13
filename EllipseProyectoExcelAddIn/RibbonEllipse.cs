using System;
using System.Data;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using Microsoft.Office.Tools.Ribbon;
using EllipseCommonsClassLibrary;
using EllipseCommonsClassLibrary.Classes;
using EllipseCommonsClassLibrary.Connections;
using Excel = Microsoft.Office.Interop.Excel;
using System.Web.Services.Ellipse;
using EllipseCommonsClassLibrary.Constants;
using EllipseCommonsClassLibrary.Utilities;
using EllipseProyectoExcelAddIn.WorkOrderService;
using Authenticator = EllipseCommonsClassLibrary.AuthenticatorService;
using Screen = EllipseCommonsClassLibrary.ScreenService;
using Oracle.ManagedDataAccess.Client;

namespace EllipseProyectoExcelAddIn
{
    public partial class RibbonEllipse
    {
        ExcelStyleCells _cells;
        EllipseFunctions _eFunctions = new EllipseFunctions();
        FormAuthenticate _frmAuth = new FormAuthenticate();

        private Excel.Application _excelApp;

        private void RibbonEllipse_Load(object sender, RibbonUIEventArgs e)
        {
            _excelApp = Globals.ThisAddIn.Application;

            var environments = Environments.GetEnvironmentList();
            foreach (var env in environments)
            {
                var item = Factory.CreateRibbonDropDownItem();
                item.Label = env;
                drpEnviroment.Items.Add(item);
            }
        }
        
        private void ExecuteQuery()
        {
            OracleConnection sqlOracleConn = null;
            try
            {
                _excelApp.Cursor = Excel.XlMousePointer.xlWait;
                var excelSheet = (Excel.Worksheet)_excelApp.ActiveWorkbook.ActiveSheet;
        
                var titleRow = 1;
                var sqlQuery = @"SELECT WORK_ORDER FROM ELLIPSE.MSF620 WO WHERE" + 
                               @" WO.RAISED_DATE = '20190124' AND WO.WORK_GROUP = 'MTOLOC'";
        
                var dbName = "EL8PROD";
                var dbUser = "SIGCON";
                var dbPass = "ventyx";
        
                var connectionTimeOut = 30;//default ODP 15
                var poolingDataBase = true;//default ODP true
        
                var connectionString = "Data Source=" + dbName + ";User ID=" + dbUser +
                                       ";Password=" + dbPass + "; Connection Timeout=" + 
                                       connectionTimeOut + "; Pooling=" + poolingDataBase.ToString().ToLower();
        
                sqlOracleConn = new OracleConnection(connectionString);
                var sqlOracleComm = new OracleCommand();
        
                if (sqlOracleConn.State != ConnectionState.Open)
                    sqlOracleConn.Open();
                sqlOracleComm.Connection = sqlOracleConn;
                sqlOracleComm.CommandText = sqlQuery;
        
                var dataReader = sqlOracleComm.ExecuteReader();
        
                if (dataReader == null)
                    return;
        
                //Cargo el encabezado de la tabla y doy formato
                for (var i = 0; i < dataReader.FieldCount; i++)
                {
                    var cell = (Excel.Range)excelSheet.Cells[titleRow, i + 1];
                    cell.Value2 = "'" + dataReader.GetName(i);
                }
        
                //cargo los datos 
                if (dataReader.IsClosed || !dataReader.HasRows) return;
        
        
                var currentRow = titleRow + 1;
                while (dataReader.Read())
                {
                    for (var i = 0; i < dataReader.FieldCount; i++)
                    {
                        var cell = (Excel.Range)excelSheet.Cells[currentRow, i + 1];
                        cell.Value2 = "'" + dataReader[i].ToString().Trim();
                    }
                        
                    currentRow++;
                }
        
            }
            catch (Exception ex)
            {
                MessageBox.Show(@"Se ha producido un error. " + ex.Message);
            }
            finally
            {
                if (sqlOracleConn != null && sqlOracleConn.State != ConnectionState.Closed)
                    sqlOracleConn.Close();
                _excelApp.Cursor = Excel.XlMousePointer.xlDefault;
            }
        }
    
        private void ExecuteQueryCommons()
        {
            try
            {
                if (_cells == null)
                    _cells = new ExcelStyleCells(_excelApp);
                _cells.SetCursorWait();
        
                var titleRow = 1;
                var sqlQuery = @"SELECT WORK_ORDER FROM ELLIPSE.MSF620 WO" +
                               @" WHERE WO.RAISED_DATE = '20190124' AND WO.WORK_GROUP = 'MTOLOC'";
                var tableName = "table";
                _eFunctions.SetDBSettings(drpEnviroment.SelectedItem.Label);
                var dataReader = _eFunctions.GetQueryResult(sqlQuery);
        
                if (dataReader == null)
                    return;
        
                //Cargo el encabezado de la tabla y doy formato
                for (var i = 0; i < dataReader.FieldCount; i++)
                    _cells.GetCell(i + 1, titleRow).Value2 = "'" + dataReader.GetName(i);
        
                _cells.FormatAsTable(_cells.GetRange(1, titleRow, dataReader.FieldCount, titleRow + 1), tableName);
        
                //cargo los datos 
                if (dataReader.IsClosed || !dataReader.HasRows) return;
        
        
                var currentRow = titleRow + 1;
                while (dataReader.Read())
                {
                    for (var i = 0; i < dataReader.FieldCount; i++)
                        _cells.GetCell(i + 1, currentRow).Value2 = "'" + dataReader[i].ToString().Trim();
                    currentRow++;
                }
        
            }
            catch (Exception ex)
            {
                Debugger.LogError("RibbonEllipse:GetQueryResult()", "\n\rMessage:" + ex.Message +
                                                                    "\n\rSource:" + ex.Source + "\n\rStackTrace:" + ex.StackTrace);
                MessageBox.Show(@"Se ha producido un error. " + ex.Message);
            }
            finally
            {
                if (_cells != null) _cells.SetCursorDefault();
                _eFunctions.CloseConnection();
            }
        }
        
        
        private void Authenticate()
        {
            //creación del servicio
            var service = new Authenticator.AuthenticatorService();
            service.Url = @"http://ews-el8prod.lmnerp01.cerrejon.com/ews/services/AuthenticatorService";
            //creación del contexto de operación
            var opContext = new Authenticator.OperationContext
            {
                district = "ICOR",
                position = ""
            };
            try
            {
                var excelSheet = (Excel.Worksheet)_excelApp.ActiveWorkbook.ActiveSheet;

                //Encabezado de consumo
                var EllipseUser = "HHERNAND";
                var EllipsePswd = "ene2014";
                var EllipseDsct = "ICOR";
                var EllipsePost = "COMC0";


                ClientConversation.authenticate(EllipseUser, EllipsePswd, EllipseDsct, EllipsePost);
                //Recuerde que el encabezado SOAP es enviado con todas las solicitudes
                service.authenticate(opContext);
            }
            catch (Exception ex)
            {
                MessageBox.Show(@"Se ha producido un error al intentar realizar la autenticación. Asegúrese que los datos ingresados sean correctos e intente nuevamente." + "\n\n" + ex.Message);
            }
        }

        private void AuthenticateCommons()
        {
            try
            {
                _frmAuth.StartPosition = FormStartPosition.CenterScreen;
                _frmAuth.SelectedEnvironment = drpEnviroment.SelectedItem.Label;
                if (_frmAuth.ShowDialog() != DialogResult.OK) return;
            }
            catch (Exception ex)
            {
                MessageBox.Show(@"Se ha producido un error al intentar realizar la autenticación. Asegúrese que los datos ingresados sean correctos e intente nuevamente." + "\n\n" + ex.Message);
            }
        }

        private void ScreenService()
        {
            //Proceso del Servicio Screen
            var service = new Screen.ScreenService();
            var urlService = _eFunctions.GetServicesUrl(drpEnviroment.SelectedItem.Label);
            service.Url = urlService + "/ScreenService";

            //Instanciar el Contexto de Operación
            var opContext = new Screen.OperationContext
            {
                district = _frmAuth.EllipseDsct,
                position = _frmAuth.EllipsePost
            };

            //Instanciar el SOAP
            ClientConversation.authenticate(_frmAuth.EllipseUser, _frmAuth.EllipsePswd);

            //Solicitud 1 y Respuesta 1
            var reply = service.executeScreen(opContext, "MSO435");

            //validamos el ingreso al programa
            if (reply.mapName != "MSM435A")
                throw new Exception ("ERROR:" + "No se pudo establecer comunicación con el servicio");

            //arreglo para los campos del screen
            var arrayFields = new ArrayScreenNameValue();

            //se adicionan los campos que se vayan a enviar
            arrayFields.Add("OPTION1I", "1");
            arrayFields.Add("MODEL_CODE1I", "CÓDIGO MODELO");
            arrayFields.Add("STAT_DATE1I", "FECHA MODELO");
            arrayFields.Add("SHIFT1I", "TURNO MODELO");

            //Solicitud 2
            var request = new Screen.ScreenSubmitRequestDTO();
            request.screenFields = arrayFields.ToArray();
            request.screenKey = "1";

            //Respuesta 2
            reply = service.submit(opContext, request);

            //Existencia y nombre de pantalla de respuesta
            if (reply == null || reply.mapName == "MSM435B")
                throw new Exception("ERROR:" + "Se ha producido un error al intentar enviar la solicitud");

            //La respuesta tiene un error o una advertencia
            if (_eFunctions.CheckReplyError(reply) || _eFunctions.CheckReplyWarning(reply))
                throw new Exception("ERROR:" + "Se ha producido un error al intentar enviar la solicitud");

            //La respuesta pide confirmación
            if (reply.functionKeys == "XMIT-Confirm")
                reply = service.submit(opContext, request);

            //si necesitas obtener los campso del reply y trabajar con ellos
            var replyFields = new ArrayScreenNameValue(reply.screenFields);
            var woProject = replyFields.GetField("WO_PROJ1I1").value.Equals("");
        }

        private void GeneralService()
        {
            try
            {
                if (_cells == null)
                    _cells = new ExcelStyleCells(_excelApp);

                //Creación del Servicio
                var service = new WorkOrderService.WorkOrderService();
                var urlService = _eFunctions.GetServicesUrl(drpEnviroment.SelectedItem.Label);
                service.Url = urlService + "/WorkOrderService";
                
                //Instanciar el Contexto de Operación
                var opContext = new WorkOrderService.OperationContext
                {
                    district = _frmAuth.EllipseDsct,
                    position = _frmAuth.EllipsePost
                };

                //Instanciar el SOAP
                ClientConversation.authenticate(_frmAuth.EllipseUser, _frmAuth.EllipsePswd);

                //Se cargan los parámetros de  la solicitud
                var request = new WorkOrderServiceCreateRequestDTO();
                request.districtCode = "ICOR";
                request.workGroup = "MTOLOC";
                request.workOrderDesc = "ORDEN DE PRUEBA";
                request.workOrderType = "CO";
                request.maintenanceType = "CO";
                request.equipmentNo = "1000016";

                //se envía la acción
                var reply = service.create(opContext, request);

                //se analiza la respuesta y se hacen las acciones pertinentes
                _cells.GetCell(1, 1).Value2 = reply.workOrder.prefix + reply.workOrder.no;
                
                
            }
            catch (Exception ex)
            {
                MessageBox.Show(@"Se ha producido un error al intentar crear la orden de trabajo." + "\n\n" + ex.Message);
            }
        }
        private void btnExecute_Click(object sender, RibbonControlEventArgs e)
        {

            _frmAuth.StartPosition = FormStartPosition.CenterScreen;
            _frmAuth.SelectedEnvironment = drpEnviroment.SelectedItem.Label;
            if (_frmAuth.ShowDialog() != DialogResult.OK) return;

            //ExecuteQuery();
            //ExecuteQueryCommons();
            //Authenticate();
            //AuthenticateCommons();
            GeneralService();
        }

        private void btnAbout_Click(object sender, RibbonControlEventArgs e)
        {
            new AboutBoxExcelAddIn("Desarrollador 1", "Desarrollador 2").ShowDialog();
        }
    }
}
