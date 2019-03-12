﻿using System;
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

            var environments = Environments.GetEnvironmentList();
            foreach (var env in environments)
            {
                var item = Factory.CreateRibbonDropDownItem();
                item.Label = env;
                drpEnviroment.Items.Add(item);
            }
        }
        //private void ExecuteQuery()
        //{
        //    OracleConnection sqlOracleConn = null;
        //    try
        //    {
        //        _excelApp.Cursor = Excel.XlMousePointer.xlWait;
        //        var excelSheet = (Excel.Worksheet)_excelApp.ActiveWorkbook.ActiveSheet;
        //
        //        var titleRow = 1;
        //        var sqlQuery = @"SELECT WORK_ORDER FROM ELLIPSE.MSF620 WO WHERE WO.RAISED_DATE = '20190124' AND WO.WORK_GROUP = 'MTOLOC'";
        //        var tableName = "table";
        //
        //        var dbName = "EL8PROD";
        //        var dbUser = "SIGCON";
        //        var dbPass = "ventyx";
        //
        //        var connectionTimeOut = 30;//default ODP 15
        //        var poolingDataBase = true;//default ODP true
        //
        //        var connectionString = "Data Source=" + dbName + ";User ID=" + dbUser + ";Password=" + dbPass + "; Connection Timeout=" + connectionTimeOut + "; Pooling=" + poolingDataBase.ToString().ToLower();
        //
        //        sqlOracleConn = new OracleConnection(connectionString);
        //        var sqlOracleComm = new OracleCommand();
        //
        //        if (sqlOracleConn.State != ConnectionState.Open)
        //            sqlOracleConn.Open();
        //        sqlOracleComm.Connection = sqlOracleConn;
        //        sqlOracleComm.CommandText = sqlQuery;
        //
        //        var dataReader = sqlOracleComm.ExecuteReader();
        //
        //        if (dataReader == null)
        //            return;
        //
        //        //Cargo el encabezado de la tabla y doy formato
        //        for (var i = 0; i < dataReader.FieldCount; i++)
        //        {
        //            var cell = (Excel.Range)excelSheet.Cells[titleRow, i + 1];
        //            cell.Value2 = "'" + dataReader.GetName(i);
        //        }
        //
        //        //cargo los datos 
        //        if (dataReader.IsClosed || !dataReader.HasRows) return;
        //
        //
        //        var currentRow = titleRow + 1;
        //        while (dataReader.Read())
        //        {
        //            for (var i = 0; i < dataReader.FieldCount; i++)
        //            {
        //                var cell = (Excel.Range)excelSheet.Cells[currentRow, i + 1];
        //                cell.Value2 = "'" + dataReader[i].ToString().Trim();
        //            }
        //                
        //            currentRow++;
        //        }
        //
        //    }
        //    catch (Exception ex)
        //    {
        //        MessageBox.Show(@"Se ha producido un error. " + ex.Message);
        //    }
        //    finally
        //    {
        //        if (sqlOracleConn != null && sqlOracleConn.State != ConnectionState.Closed)
        //            sqlOracleConn.Close();
        //        _excelApp.Cursor = Excel.XlMousePointer.xlDefault;
        //    }
        //}
        private void ExecuteQueryCommons()
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
            _frmAuth.SelectedEnvironment = drpEnviroment.SelectedItem.Label;
            if (_frmAuth.ShowDialog() != DialogResult.OK) return;

            ExecuteQueryCommons();
        }
    }
}
