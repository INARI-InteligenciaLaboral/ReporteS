using NPOI.HSSF.UserModel;
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
using Excel = Microsoft.Office.Interop.Excel;

namespace Reporte.Controllers
{
    public class HomeController : Controller
    {
        // GET: Home
        public ActionResult Index()
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
        public ActionResult GetProcesos(string empresas)
        {
            DataTable m_Procesos = SqlClass.sqldata.ObtenerPro(empresas);
            List<SelectListItem> l_Procesos = new List<SelectListItem>();
            foreach (DataRow Row in m_Procesos.Rows)
            {
                l_Procesos.Add(new SelectListItem()
                {
                    Text = Row[0].ToString() + " - " + Row[1].ToString(),
                    Value = Row[0].ToString()
                });
            }
            MultiSelectList m_List = new MultiSelectList(l_Procesos.OrderBy(i => i.Value), "Value", "Text");
            var procesos = l_Procesos;
            return Json(m_List, JsonRequestBehavior.AllowGet);
        }
        
        public ActionResult GenerarDocument([Bind(Include = "SelectedEmpresas,SelectedAnos,SelectedProcesos,SelectedMeses")] CascadingDropdownsModel m_Results)
        {
            DataTable m_result = SqlClass.sqldata.GenerarReporte(m_Results);
            WriteExcelWithNPOI(m_result,"xlsx");
            return Redirect("Index");
        }

        public void WriteExcelWithNPOI(DataTable dt, String extension)
        {

            IWorkbook workbook;

            if (extension == "xlsx")
            {
                workbook = new XSSFWorkbook();
            }
            else
            {
                throw new Exception("This format is not supported");
            }

            ISheet sheet1 = workbook.CreateSheet("Sheet 1");

            //make a header row
            IRow row1 = sheet1.CreateRow(0);

            for (int j = 0; j < dt.Columns.Count; j++)
            {

                ICell cell = row1.CreateCell(j);
                String columnName = dt.Columns[j].ToString();
                cell.SetCellValue(columnName);
            }

            //loops through data
            for (int i = 0; i < dt.Rows.Count; i++)
            {
                IRow row = sheet1.CreateRow(i + 1);
                for (int j = 0; j < dt.Columns.Count; j++)
                {

                    ICell cell = row.CreateCell(j);
                    String columnName = dt.Columns[j].ToString();
                    cell.SetCellValue(dt.Rows[i][columnName].ToString());
                }
            }

            using (var exportData = new MemoryStream())
            {
                Response.Clear();
                workbook.Write(exportData);
                if (extension == "xlsx") //xlsx file format
                {
                    Response.ContentType = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet";
                    Response.AddHeader("Content-Disposition", string.Format("attachment;filename={0}", "SIAP.xlsx"));
                    Response.BinaryWrite(exportData.ToArray());
                }
                Response.End();
            }
        }

    }
}