using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;
using Reporte.Models;
using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Linq;
using System.Web;
using System.Web.Mvc;

namespace Reporte.Controllers
{
    public class FonInfController : Controller
    {
        List<SelectListItem> lEmpresas = new List<SelectListItem>();
        List<SelectListItem> lReportes = new List<SelectListItem>();
        List<SelectListItem> lBase = new List<SelectListItem>();
        // GET: FonInf
        public ActionResult Index()
        {
            lBase.Add(new SelectListItem()
            {
                Text = "Siap",
                Value = "siap"
            });
            lBase.Add(new SelectListItem()
            {
                Text = "Soprade",
                Value = "sop"
            });
            lReportes.Add(new SelectListItem()
            {
                Text = "Infonavit",
                Value = "info"
            });
            lReportes.Add(new SelectListItem()
            {
                Text = "Fonacot",
                Value = "fona"
            });
            var lAnos = new List<SelectListItem>();
            lAnos = SqlClass.ListasStaticas.ObtenerAnos();
            var lMeses = new List<SelectListItem>();
            lMeses = SqlClass.ListasStaticas.ObtenerMesesSop();
            var lBimestre = new List<SelectListItem>();
            lBimestre = SqlClass.ListasStaticas.BimestresSop();
            var model = new CascadingDropdownsModel
            {
                Base = lBase,
                Anos = lAnos,
                Meses = lMeses,
                Bimestre = lBimestre,
                Reportes = lReportes
            };
            return View(model);
        }
        public ActionResult GetEmpresas(string reportes)
        {
            System.Data.DataTable mEmpresas;
            if (reportes.Equals(value: "siap"))
            {
                mEmpresas = SqlClass.Sqldata.ObtenerEmpInf();
            }
            else
            {
                mEmpresas = SqlClass.Sqlsop.ObtenerEmpInf();
            }
            var lEmpresas = new List<SelectListItem>();
            foreach (DataRow Row in mEmpresas.Rows)
            {
                lEmpresas.Add(new SelectListItem()
                {
                    Text = Row[1].ToString(),
                    Value = Row[0].ToString(),
                });
            }
            var mList = new MultiSelectList(lEmpresas, dataValueField: "Value", dataTextField: "Text");
            List<SelectListItem> empresas = lEmpresas;
            return Json(mList, JsonRequestBehavior.AllowGet);
        }
        public ActionResult GetProcesos(string empresas)
        {
            System.Data.DataTable mProcesos;
            if (empresas.StartsWith(value: "sopa"))
            {
                mProcesos = SqlClass.Sqlsop.ObtenerProInf(empresas.Replace(oldValue: "sopa", newValue: ""));
            }
            else
            {
                mProcesos = SqlClass.Sqldata.ObtenerProinf(empresas.Replace(oldValue: "siap", newValue: ""));
            }
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
        public ActionResult GenerarReporte([Bind(Include = "SelectedBase,SelectedEmpresas,SelectedAnos,SelectedProcesos,typeCheck,SelectedMeses, SelectedAnosMen, SelectedBimestre, PerIni, PerFin,SelectedReporte")] CascadingDropdownsModel mResults)
        {
            DataTable mConsulta;
            if(mResults.SelectedBase.Equals(value: "siap"))
            {
                if(mResults.SelectedReporte.Equals(value: "info"))
                {
                    if(mResults.typeCheck.Equals(value: "Mensual"))
                        mConsulta = SqlClass.Sqldata.GenerarReporteInfMen(mResults);
                    else
                       mConsulta = SqlClass.Sqldata.GenerarReporteInfBim(mResults);

                    WriteExcelSiap(mConsulta, mNombre: "SiapInfonavit", mResults: mResults);
                }
                else
                {
                    if (mResults.typeCheck.Equals(value: "Mensual"))
                        mConsulta = SqlClass.Sqldata.GenerarReporteFonMen(mResults);
                    else
                        mConsulta = SqlClass.Sqldata.GenerarReporteFonBim(mResults);

                    WriteExcelSiap(mConsulta, mNombre: "SiapFonacot", mResults: mResults);
                }
                
            }
            else
            {
                if (mResults.SelectedReporte.Equals(value:"info"))
                {
                    if (mResults.typeCheck.Equals(value:"Mensual"))
                        mConsulta = SqlClass.Sqlsop.GenerarReporteInfMen(mResults);
                    else if (mResults.typeCheck.Equals(value:"Bimestral"))
                        mConsulta = SqlClass.Sqlsop.GenerarReporteInfBim(mResults);
                    else
                        mConsulta = SqlClass.Sqlsop.GenerarReporteInfPer(mResults);
                    WriteExcelSopInf(mConsulta, mNombre: "SopradeInfonavit", mResults: mResults);
                }
                else
                {
                    if (mResults.typeCheck.Equals(value: "Mensual"))
                        mConsulta = SqlClass.Sqlsop.GenerarReporteFonMen(mResults);
                    else if (mResults.typeCheck.Equals(value: "Bimestral"))
                        mConsulta = SqlClass.Sqlsop.GenerarReporteFonBim(mResults);
                    else
                        mConsulta = SqlClass.Sqlsop.GenerarReporteFonPer(mResults);
                    WriteExcelSopFon(mConsulta, mNombre: "SopradeFonacot", mResults: mResults);
                }
            }
            return RedirectToAction(actionName: "Index");
        }
        public void WriteExcelSiap(System.Data.DataTable dt, String mNombre, CascadingDropdownsModel mResults)
        {

            IWorkbook workbook;


            workbook = new XSSFWorkbook();

            ISheet sheet1 = workbook.CreateSheet(sheetname: "Sheet 1");
            IRow row1 = sheet1.CreateRow(rownum: 0);
            for (int j = 0; j < dt.Columns.Count; j++)
            {
                ICell cell = row1.CreateCell(j);
                ICellStyle style = workbook.CreateCellStyle();
                if (mResults.SelectedReporte.Equals(value: "info"))
                    style.FillForegroundColor = IndexedColors.BrightGreen.Index;
                else
                    style.FillForegroundColor = IndexedColors.BlueGrey.Index;
                style.FillPattern = FillPattern.SolidForeground;
                cell.CellStyle = style;
                String columnName = dt.Columns[j].ToString();
                cell.SetCellValue(columnName);
                sheet1.AutoSizeColumn(j);
            }

            //loops through data
            for (int i = 0; i < dt.Rows.Count; i++)
            {
                IRow row = sheet1.CreateRow(i + 1);
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
                    style.FillPattern = FillPattern.SolidForeground;
                    String columnName = dt.Columns[j].ToString();
                    cell.SetCellValue(dt.Rows[i][columnName].ToString());
                    cell.CellStyle = style;
                    sheet1.AutoSizeColumn(j);
                }
            }

            using (var exportData = new MemoryStream())
            {
                Response.Clear();
                workbook.Write(exportData);
                Response.ContentType = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet";
                Response.AddHeader(name: "Content-Disposition", value: string.Format(format: "attachment;filename={0}", arg0: mNombre + "Reporte.xlsx"));
                Response.BinaryWrite(exportData.ToArray());
                Response.End();
            }
        }
        public void WriteExcelSopInf(System.Data.DataTable dt, String mNombre, CascadingDropdownsModel mResults)
        {

            IWorkbook workbook;


            workbook = new XSSFWorkbook();

            ISheet sheet1 = workbook.CreateSheet(sheetname: "Sheet 1");
            IRow row1 = sheet1.CreateRow(rownum: 0);
            for (int j = 0; j < dt.Columns.Count; j++)
            {
                ICell cell = row1.CreateCell(j);
                ICellStyle style = workbook.CreateCellStyle();
                style.FillForegroundColor = IndexedColors.BrightGreen.Index;
                style.FillPattern = FillPattern.SolidForeground;
                cell.CellStyle = style;
                String columnName = dt.Columns[j].ToString();
                cell.SetCellValue(columnName);
                sheet1.AutoSizeColumn(j);
            }

            //loops through data
            for (int i = 0; i < dt.Rows.Count; i++)
            {
                IRow row = sheet1.CreateRow(i + 1);
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
                    style.FillPattern = FillPattern.SolidForeground;
                    String columnName = dt.Columns[j].ToString();
                    cell.SetCellValue(dt.Rows[i][columnName].ToString());
                    cell.CellStyle = style;
                    sheet1.AutoSizeColumn(j);
                }
            }

            using (var exportData = new MemoryStream())
            {
                Response.Clear();
                workbook.Write(exportData);
                Response.ContentType = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet";
                Response.AddHeader(name: "Content-Disposition", value: string.Format(format: "attachment;filename={0}", arg0: mNombre + "Reporte.xlsx"));
                Response.BinaryWrite(exportData.ToArray());
                Response.End();
            }
        }
        public void WriteExcelSopFon(System.Data.DataTable dt, String mNombre, CascadingDropdownsModel mResults)
        {

            IWorkbook workbook;


            workbook = new XSSFWorkbook();

            ISheet sheet1 = workbook.CreateSheet(sheetname: "Sheet 1");
            IRow row1 = sheet1.CreateRow(rownum: 0);
            for (int j = 0; j < dt.Columns.Count; j++)
            {
                ICell cell = row1.CreateCell(j);
                ICellStyle style = workbook.CreateCellStyle();
                style.FillForegroundColor = IndexedColors.BlueGrey.Index;
                style.FillPattern = FillPattern.SolidForeground;
                cell.CellStyle = style;
                String columnName = dt.Columns[j].ToString();
                cell.SetCellValue(columnName);
                sheet1.AutoSizeColumn(j);
            }

            //loops through data
            for (int i = 0; i < dt.Rows.Count; i++)
            {
                IRow row = sheet1.CreateRow(i + 1);
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
                    style.FillPattern = FillPattern.SolidForeground;
                    String columnName = dt.Columns[j].ToString();
                    cell.SetCellValue(dt.Rows[i][columnName].ToString());
                    cell.CellStyle = style;
                    sheet1.AutoSizeColumn(j);
                }
            }

            using (var exportData = new MemoryStream())
            {
                Response.Clear();
                workbook.Write(exportData);
                Response.ContentType = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet";
                Response.AddHeader(name: "Content-Disposition", value: string.Format(format: "attachment;filename={0}", arg0: mNombre + "Reporte.xlsx"));
                Response.BinaryWrite(exportData.ToArray());
                Response.End();
            }
        }
    }
}