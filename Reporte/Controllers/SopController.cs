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
        public bool m_importe = false;
        public bool m_mensual = false;
        List<SelectListItem> l_empresas = new List<SelectListItem>();

        // GET: Sop
        public ActionResult Index()
        {
            return View();
        }
        // GET: Home
        public ActionResult SopradeMensual()
        {
            System.Data.DataTable m_Empresas = SqlClass.sqlsop.ObtenerEmp();

            foreach (DataRow Row in m_Empresas.Rows)
            {
                l_empresas.Add(new SelectListItem()
                {
                    Text = Row[1].ToString(),
                    Value = Row[0].ToString()
                });
            }
            List<SelectListItem> l_anos = new List<SelectListItem>();
            l_anos = SqlClass.ListasStaticas.ObtenerAnos();
            List<SelectListItem> l_meses = new List<SelectListItem>();
            l_meses = SqlClass.ListasStaticas.ObtenerMesesSop();
            var model = new CascadingDropdownsModel
            {
                Empresas = l_empresas,
                Anos = l_anos,
                Meses = l_meses
            };
            return View(model);
        }
        // GET: Home
        public ActionResult SopradeProcesos()
        {
            System.Data.DataTable m_Empresas = SqlClass.sqlsop.ObtenerEmp();
            List<SelectListItem> l_empresas = new List<SelectListItem>();
            foreach (DataRow Row in m_Empresas.Rows)
            {
                l_empresas.Add(new SelectListItem()
                {
                    Text = Row[1].ToString(),
                    Value = Row[0].ToString()
                });
            }
            var model = new CascadingDropdownsModel
            {
                Empresas = l_empresas
            };
            return View(model);
        }
        public ActionResult GetProcesos(string empresas)
        {
            System.Data.DataTable m_Procesos = SqlClass.sqlsop.ObtenerPro(empresas);
            List<SelectListItem> l_Procesos = new List<SelectListItem>();
            string sop_proceso = "";
            foreach (DataRow Row in m_Procesos.Rows)
            {
                if (!sop_proceso.Equals(""))
                {
                    sop_proceso = sop_proceso + ",";
                }
                sop_proceso = sop_proceso + Row[0].ToString();
                l_Procesos.Add(new SelectListItem()
                {
                    Text = Row[1].ToString(),
                    Value = Row[0].ToString(),
                });
            }
            l_Procesos.Add(new SelectListItem()
            {
                Text = "Todos los procesos",
                Value = sop_proceso
            });
            MultiSelectList m_List = new MultiSelectList(l_Procesos, "Value", "Text");
            var procesos = l_Procesos;
            return Json(m_List, JsonRequestBehavior.AllowGet);
        }

        public ActionResult GenerarDocument([Bind(Include = "SelectedEmpresas,SelectedAnos,SelectedProcesos,SelectedMeses")] CascadingDropdownsModel m_Results)
        {
            while (m_mensual)
            {
                Thread.Sleep(5000);
            }
            m_mensual = true;
            System.Data.DataTable m_result = SqlClass.sqlsop.GenerarReporte(m_Results);
            WriteExcelWithNPOI(m_result, m_Results.SelectedAnos + "-" + m_Results.SelectedEmpresas, m_Results);
            m_mensual = false;
            return RedirectToAction("Index");
        }
        public ActionResult GenerarDocumentPro([Bind(Include = "SelectedEmpresas,SelectedProcesos,SelectedMeses")] CascadingDropdownsModel m_Results)
        {
            while (m_importe)
            {
                Thread.Sleep(5000);
            }
            m_importe = true;
            WriteExcelWith(m_Results.SelectedMeses + "-" + m_Results.SelectedEmpresas, m_Results);
            m_importe = false;
            return RedirectToAction("SopradeProcesos", "Sop");
        }
        public void WriteExcelWithNPOI(System.Data.DataTable dt, String m_nombre, CascadingDropdownsModel m_Results)
        {

            IWorkbook workbook;


            workbook = new XSSFWorkbook();

            ISheet sheet1 = workbook.CreateSheet("Sheet 1");


            IRow rows = sheet1.CreateRow(0);
            ICell celda_title = rows.CreateCell(0);
            sheet1.AddMergedRegion(new NPOI.SS.Util.CellRangeAddress(0, 0, 0, 44));
            var font = workbook.CreateFont();
            font.FontHeightInPoints = 14;
            font.FontName = "Calibri";
            font.Boldweight = (short)FontBoldWeight.Bold;
            ICellStyle celda_style = workbook.CreateCellStyle();
            celda_style.FillForegroundColor = IndexedColors.White.Index;
            celda_style.FillPattern = FillPattern.SolidForeground;
            celda_style.SetFont(font);
            celda_title.CellStyle = celda_style;
            if (m_Results.SelectedMeses.Length > 2)
            {
                celda_title.SetCellValue("Reporte anual del año " + m_Results.SelectedAnos);
            }
            else
            {
                switch (Convert.ToInt16(m_Results.SelectedMeses))
                {
                    case 1:
                        celda_title.SetCellValue("Reporte del mes de Enero del año " + m_Results.SelectedAnos);
                        break;
                    case 2:
                        celda_title.SetCellValue("Reporte del mes de Febrero del año " + m_Results.SelectedAnos);
                        break;
                    case 3:
                        celda_title.SetCellValue("Reporte del mes de Marzo del año " + m_Results.SelectedAnos);
                        break;
                    case 4:
                        celda_title.SetCellValue("Reporte del mes de Abril del año " + m_Results.SelectedAnos);
                        break;
                    case 5:
                        celda_title.SetCellValue("Reporte del mes de Mayo del año " + m_Results.SelectedAnos);
                        break;
                    case 6:
                        celda_title.SetCellValue("Reporte del mes de Junio del año " + m_Results.SelectedAnos);
                        break;
                    case 7:
                        celda_title.SetCellValue("Reporte del mes de Julio del año " + m_Results.SelectedAnos);
                        break;
                    case 8:
                        celda_title.SetCellValue("Reporte del mes de Agosto del año " + m_Results.SelectedAnos);
                        break;
                    case 9:
                        celda_title.SetCellValue("Reporte del mes de Septiembre del año " + m_Results.SelectedAnos);
                        break;
                    case 10:
                        celda_title.SetCellValue("Reporte del mes de Octubre del año " + m_Results.SelectedAnos);
                        break;
                    case 11:
                        celda_title.SetCellValue("Reporte del mes de Noviembre del año " + m_Results.SelectedAnos);
                        break;
                    case 12:
                        celda_title.SetCellValue("Reporte del mes de Noviembre del año " + m_Results.SelectedAnos);
                        break;
                };
            }
            IRow rows1 = sheet1.CreateRow(1);
            ICell celda_title1 = rows1.CreateCell(0);
            sheet1.AddMergedRegion(new NPOI.SS.Util.CellRangeAddress(1, 1, 0, 44));
            var font1 = workbook.CreateFont();
            font1.FontHeightInPoints = 14;
            font1.FontName = "Calibri";
            font1.Boldweight = (short)FontBoldWeight.Bold;
            ICellStyle celda_style1 = workbook.CreateCellStyle();
            celda_style1.FillForegroundColor = IndexedColors.White.Index;
            celda_style1.FillPattern = FillPattern.SolidForeground;
            celda_style1.SetFont(font1);
            celda_title1.CellStyle = celda_style1;
            celda_title1.SetCellValue(SqlClass.sqlsop.Empresa_Title(m_Results.SelectedEmpresas));

            IRow row1 = sheet1.CreateRow(2);
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
                Response.AddHeader("Content-Disposition", string.Format("attachment;filename={0}", "Soprade " + m_nombre + "Reporte.xlsx"));
                Response.BinaryWrite(exportData.ToArray());
                Response.End();
            }
        }
        public void WriteExcelWith(String m_nombre, CascadingDropdownsModel m_Results)
        {
            IWorkbook workbook;

            workbook = new XSSFWorkbook();
            int inicio = 0;
            int fin = 0;
            bool m_Aguinaldo = false;
            Char delimiter = '-';
            String[] substrings = m_Results.SelectedMeses.Split(delimiter);
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
                m_Aguinaldo = true;
            while (fin >= inicio)
            {
                m_Results.SelectedMeses = inicio.ToString();
                System.Data.DataTable dt = SqlClass.sqlsop.GenerarReporteImp(m_Results);
                ISheet sheet1 = workbook.CreateSheet(inicio.ToString());
                if (!(m_Results.SelectedProcesos.Length > 4))
                {
                    IRow rows = sheet1.CreateRow(0);
                    ICell celda_title = rows.CreateCell(0);
                    sheet1.AddMergedRegion(new NPOI.SS.Util.CellRangeAddress(0, 0, 0, 44));
                    var font = workbook.CreateFont();
                    font.FontHeightInPoints = 14;
                    font.FontName = "Calibri";
                    font.Boldweight = (short)FontBoldWeight.Bold;
                    ICellStyle celda_style = workbook.CreateCellStyle();
                    celda_style.FillForegroundColor = IndexedColors.White.Index;
                    celda_style.FillPattern = FillPattern.SolidForeground;
                    celda_style.SetFont(font);
                    celda_title.CellStyle = celda_style;
                    celda_title.SetCellValue(SqlClass.sqlsop.ProcesosTitle(m_Results, inicio));
                }


                IRow rows1 = sheet1.CreateRow(1);
                ICell celda_title1 = rows1.CreateCell(0);
                sheet1.AddMergedRegion(new NPOI.SS.Util.CellRangeAddress(1, 1, 0, 44));
                var font1 = workbook.CreateFont();
                font1.FontHeightInPoints = 14;
                font1.FontName = "Calibri";
                font1.Boldweight = (short)FontBoldWeight.Bold;
                ICellStyle celda_style1 = workbook.CreateCellStyle();
                celda_style1.FillForegroundColor = IndexedColors.White.Index;
                celda_style1.FillPattern = FillPattern.SolidForeground;
                celda_style1.SetFont(font1);
                celda_title1.CellStyle = celda_style1;
                celda_title1.SetCellValue(SqlClass.sqlsop.Empresa_Title(m_Results.SelectedEmpresas));

                IRow row1 = sheet1.CreateRow(2);
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
                    if (m_Aguinaldo)
                    {
                        fin = 2015702;
                        inicio = 2015702;
                        m_Aguinaldo = false;
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
                Response.AddHeader("Content-Disposition", string.Format("attachment;filename={0}", "Soprade " + m_nombre + "Reporte.xlsx"));
                Response.BinaryWrite(exportData.ToArray());
                Response.End();
            }
        }
    }
}