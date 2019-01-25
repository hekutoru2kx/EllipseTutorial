using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using Microsoft.Office.Tools.Ribbon;
using EllipseCommonsClassLibrary;
using EllipseCommonsClassLibrary.Classes;
using EllipseCommonsClassLibrary.Connections;
using Excel = Microsoft.Office.Interop.Excel;

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

            var enviroments = Environments.GetEnviromentList();
            foreach (var env in enviroments)
            {
                var item = Factory.CreateRibbonDropDownItem();
                item.Label = env;
                drpEnviroment.Items.Add(item);
            }
        }

        private void ExecuteQuery()
        {
            try
            {
                if (_cells == null)
                    _cells = new ExcelStyleCells(_excelApp);
                _cells.SetCursorWait();

                var titleRow = 1;
                var sqlQuery = @"SELECT WORK_ORDER FROM ELLIPSE.MSF620 WO WHERE WO.RAISED_DATE = '20190124' AND WO.WORK_GROUP = 'MTOLOC'";
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
                Debugger.LogError("RibbonEllipse:GetQueryResult()", "\n\rMessage:" + ex.Message + "\n\rSource:" + ex.Source + "\n\rStackTrace:" + ex.StackTrace);
                MessageBox.Show(@"Se ha producido un error. " + ex.Message);
            }
            finally
            {
                if (_cells != null) _cells.SetCursorDefault();
                _eFunctions.CloseConnection();
            }
        }

        private void btnExecute_Click(object sender, RibbonControlEventArgs e)
        {
            _frmAuth.StartPosition = FormStartPosition.CenterScreen;
            _frmAuth.SelectedEnviroment = drpEnviroment.SelectedItem.Label;
            if (_frmAuth.ShowDialog() != DialogResult.OK) return;

            ExecuteQuery();
        }
    }
}
