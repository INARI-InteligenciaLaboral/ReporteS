using Microsoft.Office.Interop.Excel;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;
using Reporte.Models;
using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Linq;
using System.Threading;
using System.Web;
using System.Web.Mvc;

namespace Reporte.Controllers
{
    public class SopController : Controller
    {
        public bool MImporte = false;
        public bool MMensual = false;
        List<SelectListItem> lEmpresas = new List<SelectListItem>();

        // GET: Sop
        public ActionResult Index()
        {
            return View();
        }
        // GET: Home
        public ActionResult SopradeMensual()
        {
            System.Data.DataTable mEmpresas = SqlClass.Sqlsop.ObtenerEmp();

            foreach (DataRow Row in mEmpresas.Rows)
            {
                lEmpresas.Add(new SelectListItem()
                {
                    Text = Row[1].ToString(),
                    Value = Row[0].ToString()
                });
            }
            var lAnos = new List<SelectListItem>();
            lAnos = SqlClass.ListasStaticas.ObtenerAnos();
            var lMeses = new List<SelectListItem>();
            lMeses = SqlClass.ListasStaticas.ObtenerMesesSop();
            var model = new CascadingDropdownsModel
            {
                Empresas = lEmpresas,
                Anos = lAnos,
                Meses = lMeses
            };
            return View(model);
        }
        // GET: Home
        public ActionResult SopradeProcesos()
        {
            System.Data.DataTable mEmpresas = SqlClass.Sqlsop.ObtenerEmp();
            var lEmpresas = new List<SelectListItem>();
            foreach (DataRow Row in mEmpresas.Rows)
            {
                lEmpresas.Add(new SelectListItem()
                {
                    Text = Row[1].ToString(),
                    Value = Row[0].ToString()
                });
            }
            var model = new CascadingDropdownsModel
            {
                Empresas = lEmpresas
            };
            return View(model);
        }
        public ActionResult GetProcesos(string empresas)
        {
            System.Data.DataTable mProcesos = SqlClass.Sqlsop.ObtenerPro(empresas);
            var lProcesos = new List<SelectListItem>();
            string sopProceso = "";
            foreach (DataRow Row in mProcesos.Rows)
            {
                if (!sopProceso.Equals(value: ""))
                {
                    sopProceso = sopProceso + ",";
                }
                sopProceso = sopProceso + Row[0].ToString();
                lProcesos.Add(new SelectListItem()
                {
                    Text = Row[1].ToString(),
                    Value = Row[0].ToString(),
                });
            }
            lProcesos.Add(new SelectListItem()
            {
                Text = "Todos los procesos",
                Value = sopProceso
            });
            var mList = new MultiSelectList(lProcesos, dataValueField: "Value", dataTextField: "Text");
            List<SelectListItem> procesos = lProcesos;
            return Json(mList, JsonRequestBehavior.AllowGet);
        }

        public ActionResult GenerarDocument([Bind(Include = "SelectedEmpresas,SelectedAnos,SelectedProcesos,SelectedMeses")] CascadingDropdownsModel mResults)
        {
            while (MMensual)
            {
                Thread.Sleep(millisecondsTimeout: 5000);
            }
            MMensual = true;
            System.Data.DataTable mResult = SqlClass.Sqlsop.GenerarReporte(mResults);
            WriteExcelWithNpoi(mResult, mResults.SelectedAnos + "-" + mResults.SelectedEmpresas, mResults);
            MMensual = false;
            return RedirectToAction(actionName: "Index");
        }
        public ActionResult GenerarDocumentPro([Bind(Include = "SelectedEmpresas,SelectedProcesos,SelectedMeses")] CascadingDropdownsModel mResults)
        {
            while (MImporte)
            {
                Thread.Sleep(millisecondsTimeout: 5000);
            }
            MImporte = true;
            WriteExcelWith(mResults.SelectedMeses + "-" + mResults.SelectedEmpresas, mResults);
            MImporte = false;
            return RedirectToAction(actionName: "SopradeProcesos", controllerName: "Sop");
        }
        public void WriteExcelWithNpoi(System.Data.DataTable dt, String mNombre, CascadingDropdownsModel mResults)
        {

            IWorkbook workbook;


            workbook = new XSSFWorkbook();

            ISheet sheet1 = workbook.CreateSheet(sheetname: "Sheet 1");


            IRow rows = sheet1.CreateRow(rownum: 0);
            ICell celdaTitle = rows.CreateCell(column: 0);
            sheet1.AddMergedRegion(new NPOI.SS.Util.CellRangeAddress(firstRow: 0, lastRow: 0, firstCol: 0, lastCol: 44));
            NPOI.SS.UserModel.IFont font = workbook.CreateFont();
            font.FontHeightInPoints = 14;
            font.FontName = "Calibri";
            font.Boldweight = (short)FontBoldWeight.Bold;
            ICellStyle celdaStyle = workbook.CreateCellStyle();
            celdaStyle.FillForegroundColor = IndexedColors.White.Index;
            celdaStyle.FillPattern = FillPattern.SolidForeground;
            celdaStyle.SetFont(font);
            celdaTitle.CellStyle = celdaStyle;
            if (mResults.SelectedMeses.Length > 2)
            {
                celdaTitle.SetCellValue("Reporte anual del año " + mResults.SelectedAnos);
            }
            else
            {
                switch (Convert.ToInt16(mResults.SelectedMeses))
                {
                    case 1:
                        celdaTitle.SetCellValue("Reporte del mes de Enero del año " + mResults.SelectedAnos);
                        break;
                    case 2:
                        celdaTitle.SetCellValue("Reporte del mes de Febrero del año " + mResults.SelectedAnos);
                        break;
                    case 3:
                        celdaTitle.SetCellValue("Reporte del mes de Marzo del año " + mResults.SelectedAnos);
                        break;
                    case 4:
                        celdaTitle.SetCellValue("Reporte del mes de Abril del año " + mResults.SelectedAnos);
                        break;
                    case 5:
                        celdaTitle.SetCellValue("Reporte del mes de Mayo del año " + mResults.SelectedAnos);
                        break;
                    case 6:
                        celdaTitle.SetCellValue("Reporte del mes de Junio del año " + mResults.SelectedAnos);
                        break;
                    case 7:
                        celdaTitle.SetCellValue("Reporte del mes de Julio del año " + mResults.SelectedAnos);
                        break;
                    case 8:
                        celdaTitle.SetCellValue("Reporte del mes de Agosto del año " + mResults.SelectedAnos);
                        break;
                    case 9:
                        celdaTitle.SetCellValue("Reporte del mes de Septiembre del año " + mResults.SelectedAnos);
                        break;
                    case 10:
                        celdaTitle.SetCellValue("Reporte del mes de Octubre del año " + mResults.SelectedAnos);
                        break;
                    case 11:
                        celdaTitle.SetCellValue("Reporte del mes de Noviembre del año " + mResults.SelectedAnos);
                        break;
                    case 12:
                        celdaTitle.SetCellValue("Reporte del mes de Noviembre del año " + mResults.SelectedAnos);
                        break;
                }
            }
            IRow rows1 = sheet1.CreateRow(rownum: 1);
            ICell celdaTitle1 = rows1.CreateCell(column: 0);
            sheet1.AddMergedRegion(new NPOI.SS.Util.CellRangeAddress(firstRow: 1, lastRow: 1, firstCol: 0, lastCol: 44));
            NPOI.SS.UserModel.IFont font1 = workbook.CreateFont();
            font1.FontHeightInPoints = 14;
            font1.FontName = "Calibri";
            font1.Boldweight = (short)FontBoldWeight.Bold;
            ICellStyle celdaStyle1 = workbook.CreateCellStyle();
            celdaStyle1.FillForegroundColor = IndexedColors.White.Index;
            celdaStyle1.FillPattern = FillPattern.SolidForeground;
            celdaStyle1.SetFont(font1);
            celdaTitle1.CellStyle = celdaStyle1;
            celdaTitle1.SetCellValue(SqlClass.Sqlsop.Empresa_Title(mResults.SelectedEmpresas));

            IRow row1 = sheet1.CreateRow(rownum: 2);
            for (int j = 0; j < dt.Columns.Count; j++)
            {
                ICell cell = row1.CreateCell(j);
                ICellStyle style = workbook.CreateCellStyle();
                style.FillForegroundColor = IndexedColors.White.Index;
                if (j > 8 && j < 21)
                {
                    style.FillForegroundColor = IndexedColors.LightBlue.Index;
                }
                else if (j > 20 && j < 42)
                {
                    style.FillForegroundColor = IndexedColors.Red.Index;
                }
                else if (j == 42)
                {
                    style.FillForegroundColor = IndexedColors.Green.Index;
                }
                else if (j == 43 || j == 44)
                {
                    style.FillForegroundColor = IndexedColors.Yellow.Index;
                }

                style.FillPattern = FillPattern.SolidForeground;
                cell.CellStyle = style;
                String columnName = dt.Columns[j].ToString();
                cell.SetCellValue(columnName);
                sheet1.AutoSizeColumn(j);
            }

            //loops through data
            for (int i = 0; i < dt.Rows.Count; i++)
            {
                IRow row = sheet1.CreateRow(i + 3);
                for (int j = 0; j < dt.Columns.Count; j++)
                {

                    ICell cell = row.CreateCell(j);
                    ICellStyle style = workbook.CreateCellStyle();
                    style.FillForegroundColor = IndexedColors.White.Index;
                    if (i > 0)
                    {
                        if (i % 2 == 0)
                        {

                            style.FillForegroundColor = IndexedColors.SkyBlue.Index;
                        }
                    }
                    else
                    {
                        style.FillForegroundColor = IndexedColors.SkyBlue.Index;

                    }
                    if (i + 1 == dt.Rows.Count)
                    {
                        style.FillForegroundColor = IndexedColors.LightYellow.Index;
                    }
                    style.FillPattern = FillPattern.SolidForeground;
                    String columnName = dt.Columns[j].ToString();
                    if (j > 8 && j < 45 || j == 5 || j == 6)
                    {
                        cell.SetCellType(CellType.Numeric);
                        cell.SetCellValue(Convert.ToDouble(dt.Rows[i][columnName].ToString()));
                    }
                    else
                    {
                        cell.SetCellValue(dt.Rows[i][columnName].ToString());
                    }
                    cell.CellStyle = style;
                    sheet1.AutoSizeColumn(j);
                }
            }

            using (var exportData = new MemoryStream())
            {
                Response.Clear();
                workbook.Write(exportData);
                Response.ContentType = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet";
                Response.AddHeader(name: "Content-Disposition", value: string.Format(format: "attachment;filename={0}", arg0: "Soprade " + mNombre + "Reporte.xlsx"));
                Response.BinaryWrite(exportData.ToArray());
                Response.End();
            }
        }
        public void WriteExcelWith(String mNombre, CascadingDropdownsModel mResults)
        {
            IWorkbook workbook;

            workbook = new XSSFWorkbook();
            int inicio = 0;
            int fin = 0;
            bool mAguinaldo = false;
            Char delimiter = '-';
            String[] substrings = mResults.SelectedMeses.Split(delimiter);
            foreach (var substring in substrings)
            {
                if (inicio > 0)
                    fin = Convert.ToInt32(substring);
                else
                    inicio = Convert.ToInt32(substring);
            }
            if (fin == 0)
                fin = inicio;
            if ((inicio <= 2015224 && fin >= 2015224) || (inicio <= 2015252 && fin >= 2015252))
                mAguinaldo = true;
            while (fin >= inicio)
            {
                mResults.SelectedMeses = inicio.ToString();
                System.Data.DataTable dt = SqlClass.Sqlsop.GenerarReporteImp(mResults);
                ISheet sheet1 = workbook.CreateSheet(inicio.ToString());
                if (!(mResults.SelectedProcesos.Length > 4))
                {
                    IRow rows = sheet1.CreateRow(rownum: 0);
                    ICell celdaTitle = rows.CreateCell(column: 0);
                    sheet1.AddMergedRegion(new NPOI.SS.Util.CellRangeAddress(firstRow: 0, lastRow: 0, firstCol: 0, lastCol: 44));
                    NPOI.SS.UserModel.IFont font = workbook.CreateFont();
                    font.FontHeightInPoints = 14;
                    font.FontName = "Calibri";
                    font.Boldweight = (short)FontBoldWeight.Bold;
                    ICellStyle celdaStyle = workbook.CreateCellStyle();
                    celdaStyle.FillForegroundColor = IndexedColors.White.Index;
                    celdaStyle.FillPattern = FillPattern.SolidForeground;
                    celdaStyle.SetFont(font);
                    celdaTitle.CellStyle = celdaStyle;
                    celdaTitle.SetCellValue(SqlClass.Sqlsop.ProcesosTitle(mResults, inicio));
                }


                IRow rows1 = sheet1.CreateRow(rownum: 1);
                ICell celdaTitle1 = rows1.CreateCell(column: 0);
                sheet1.AddMergedRegion(new NPOI.SS.Util.CellRangeAddress(firstRow: 1, lastRow: 1, firstCol: 0, lastCol: 44));
                NPOI.SS.UserModel.IFont font1 = workbook.CreateFont();
                font1.FontHeightInPoints = 14;
                font1.FontName = "Calibri";
                font1.Boldweight = (short)FontBoldWeight.Bold;
                ICellStyle celdaStyle1 = workbook.CreateCellStyle();
                celdaStyle1.FillForegroundColor = IndexedColors.White.Index;
                celdaStyle1.FillPattern = FillPattern.SolidForeground;
                celdaStyle1.SetFont(font1);
                celdaTitle1.CellStyle = celdaStyle1;
                celdaTitle1.SetCellValue(SqlClass.Sqlsop.Empresa_Title(mResults.SelectedEmpresas));

                IRow row1 = sheet1.CreateRow(rownum: 2);
                for (int j = 0; j < dt.Columns.Count; j++)
                {
                    ICell cell = row1.CreateCell(j);
                    ICellStyle style = workbook.CreateCellStyle();
                    style.FillForegroundColor = IndexedColors.White.Index;
                    if (j > 8 && j < 21)
                    {
                        style.FillForegroundColor = IndexedColors.LightBlue.Index;
                    }
                    else if (j > 20 && j < 42)
                    {
                        style.FillForegroundColor = IndexedColors.Red.Index;
                    }
                    else if (j == 42)
                    {
                        style.FillForegroundColor = IndexedColors.Green.Index;
                    }
                    else if (j == 43 || j == 44)
                    {
                        style.FillForegroundColor = IndexedColors.Yellow.Index;
                    }

                    style.FillPattern = FillPattern.SolidForeground;
                    cell.CellStyle = style;
                    String columnName = dt.Columns[j].ToString();
                    cell.SetCellValue(columnName);
                    sheet1.AutoSizeColumn(j);
                }

                if (dt.Rows.Count > 1)
                    for (int i = 0; i < dt.Rows.Count; i++)
                    {
                        IRow row = sheet1.CreateRow(i + 3);
                        for (int j = 0; j < dt.Columns.Count; j++)
                        {

                            ICell cell = row.CreateCell(j);
                            ICellStyle style = workbook.CreateCellStyle();
                            style.FillForegroundColor = IndexedColors.White.Index;
                            if (i > 0)
                            {
                                if (i % 2 == 0)
                                {

                                    style.FillForegroundColor = IndexedColors.SkyBlue.Index;
                                }
                            }
                            else
                            {
                                style.FillForegroundColor = IndexedColors.SkyBlue.Index;
                            }
                            if (i + 1 == dt.Rows.Count)
                            {
                                style.FillForegroundColor = IndexedColors.LightYellow.Index;
                            }
                            style.FillPattern = FillPattern.SolidForeground;
                            String columnName = dt.Columns[j].ToString();
                            if (j > 8 && j < 45 || j == 5 || j == 6)
                            {
                                cell.SetCellType(CellType.Numeric);
                                cell.SetCellValue(Convert.ToDouble(dt.Rows[i][columnName].ToString()));
                            }
                            else
                            {
                                cell.SetCellValue(dt.Rows[i][columnName].ToString());
                            }
                            cell.CellStyle = style;
                            sheet1.AutoSizeColumn(j);
                        }
                    }
                if (fin == inicio)
                {
                    if (mAguinaldo)
                    {
                        fin = 2015702;
                        inicio = 2015702;
                        mAguinaldo = false;
                    }
                    else
                        inicio++;
                }
                else
                    inicio++;
            }


            using (var exportData = new MemoryStream())
            {
                Response.Clear();
                workbook.Write(exportData);
                Response.ContentType = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet";
                Response.AddHeader(name: "Content-Disposition", value: string.Format(format: "attachment;filename={0}", arg0: "Soprade " + mNombre + "Reporte.xlsx"));
                Response.BinaryWrite(exportData.ToArray());
                Response.End();
            }
        }
    }
}