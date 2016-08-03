using NPOI.HSSF.Util;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;
using Reporte.Models;
using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Threading;
using System.Web.Mvc;
//using Excel = Microsoft.Office.Interop.Excel;

namespace Reporte.Controllers
{
    public class HomeController : Controller
    {
        private static IDictionary<Guid, int> tasks = new Dictionary<Guid, int>();
        public bool m_importe = false;
        public bool m_mensual = false;
        // GET: Home
        public ActionResult Index()
        {
            return View();
        }
        // GET: Home
        public ActionResult FilterMensual()
        {
            DataTable m_Empresas = SqlClass.sqldata.ObtenerEmp();
            List<SelectListItem> l_empresas = new List<SelectListItem>();
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
            l_meses = SqlClass.ListasStaticas.ObtenerMeses();
            var model = new CascadingDropdownsModel
            {
                Empresas = l_empresas,
                Anos = l_anos,
                Meses = l_meses
            };
            return View(model);
        }
        // GET: Home
        public ActionResult FilterProcesos()
        {
            DataTable m_Empresas = SqlClass.sqldata.ObtenerEmp();
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
            DataTable m_Procesos = SqlClass.sqldata.ObtenerPro(empresas);
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
                    Text = Row[0].ToString() + " - " + Row[1].ToString(),
                    Value = Row[0].ToString()
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
            while(m_mensual)
            {
                Thread.Sleep(5000);
            }
            m_mensual = true;
            DataTable m_result = SqlClass.sqldata.GenerarReporte(m_Results);
            WriteExcelWithNPOI(m_result, m_Results.SelectedAnos + "-" + m_Results.SelectedEmpresas);
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
            WriteExcelWith( m_Results.SelectedMeses + "-" + m_Results.SelectedEmpresas, m_Results);
            m_importe = false;
            return RedirectToAction("Home","FilterProcesos");
        }
        
        public void WriteExcelWithNPOI(DataTable dt, String m_nombre)
        {

            IWorkbook workbook;


            workbook = new XSSFWorkbook();
            
            ISheet sheet1 = workbook.CreateSheet("Sheet 1");

            //make a header row
            IRow row1 = sheet1.CreateRow(0);

            for (int j = 0; j < dt.Columns.Count; j++)
            {
                ICell cell = row1.CreateCell(j);
                ICellStyle style = workbook.CreateCellStyle();
                style.FillForegroundColor = IndexedColors.White.Index;
                if (j > 4 && j < 17)
                {
                    style.FillForegroundColor = IndexedColors.LightBlue.Index;
                }
                else if (j > 16 && j < 37)
                {
                    style.FillForegroundColor = IndexedColors.Red.Index;
                }
                else if (j == 37)
                {
                    style.FillForegroundColor = IndexedColors.Green.Index;
                }
                else if (j == 38 || j == 39)
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
                    cell.CellStyle = style;
                    String columnName = dt.Columns[j].ToString();
                    cell.SetCellValue(dt.Rows[i][columnName].ToString());
                    sheet1.AutoSizeColumn(j);
                }
            }

            using (var exportData = new MemoryStream())
            {
                Response.Clear();
                workbook.Write(exportData);
                Response.ContentType = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet";
                Response.AddHeader("Content-Disposition", string.Format("attachment;filename={0}", m_nombre + "Reporte.xlsx"));
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
            while(fin >= inicio)
            {
                m_Results.SelectedMeses = inicio.ToString();
                DataTable dt = SqlClass.sqldata.GenerarReporteImp(m_Results);
                ISheet sheet1 = workbook.CreateSheet(inicio.ToString());

                //make a header row
                IRow row1 = sheet1.CreateRow(0);
                for (int j = 0; j < dt.Columns.Count; j++)
                {
                    ICell cell = row1.CreateCell(j);
                    ICellStyle style = workbook.CreateCellStyle();
                    style.FillForegroundColor = IndexedColors.White.Index;
                    if (j > 4 && j < 17)
                    {
                        style.FillForegroundColor = IndexedColors.LightBlue.Index;
                    }
                    else if (j > 16 && j < 37)
                    {
                        style.FillForegroundColor = IndexedColors.Red.Index;
                    }
                    else if (j == 37)
                    {
                        style.FillForegroundColor = IndexedColors.Green.Index;
                    }
                    else if (j == 38 || j == 39)
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
                        cell.CellStyle = style;
                        String columnName = dt.Columns[j].ToString();
                        cell.SetCellValue(dt.Rows[i][columnName].ToString());
                        sheet1.AutoSizeColumn(j);
                    }
                }
                inicio++;
            }
            
            
            using (var exportData = new MemoryStream())
            {
                Response.Clear();
                workbook.Write(exportData);
                Response.ContentType = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet";
                Response.AddHeader("Content-Disposition", string.Format("attachment;filename={0}", m_nombre + "Reporte.xlsx"));
                Response.BinaryWrite(exportData.ToArray());
                Response.End();
            }
        }
    }
}