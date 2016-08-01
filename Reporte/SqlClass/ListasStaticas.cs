using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Web.Mvc;

namespace Reporte.SqlClass
{
    public class ListasStaticas
    {
        public static List<SelectListItem> ObtenerAnos()
        {
            List<SelectListItem> l_anos = new List<SelectListItem>();
            l_anos.Add(
                new SelectListItem()
                {
                    Text = "2010",
                    Value = "2010"
                }
            );
            l_anos.Add(
                new SelectListItem()
                {
                    Text = "2011",
                    Value = "2011"
                }
            );
            l_anos.Add(
                new SelectListItem()
                {
                    Text = "2012",
                    Value = "2012"
                }
            );
            l_anos.Add(
                new SelectListItem()
                {
                    Text = "2013",
                    Value = "2013"
                }
            );
            l_anos.Add(
                new SelectListItem()
                {
                    Text = "2014",
                    Value = "2014"
                }
            );
            l_anos.Add(
                new SelectListItem()
                {
                    Text = "2015",
                    Value = "2015"
                }
            );
            l_anos.Add(
                new SelectListItem()
                {
                    Text = "2016",
                    Value = "2016"
                }
            );
            return l_anos;
        }
        public static List<SelectListItem> ObtenerMeses()
        {
            List<SelectListItem> l_meses = new List<SelectListItem>();
            l_meses.Add(
                new SelectListItem()
                {
                    Text = "Todos",
                    Value = "1,2,3,4,5,6,7,8,9,10,11,12"
                }
            );
            l_meses.Add(
                new SelectListItem()
                {
                    Text = "Enero",
                    Value = "1"
                }
            );
            l_meses.Add(
                new SelectListItem()
                {
                    Text = "Febrero",
                    Value = "2"
                }
            );
            l_meses.Add(
                new SelectListItem()
                {
                    Text = "Marzo",
                    Value = "3"
                }
            );
            l_meses.Add(
                new SelectListItem()
                {
                    Text = "Abril",
                    Value = "4"
                }
            );
            l_meses.Add(
                new SelectListItem()
                {
                    Text = "Mayo",
                    Value = "5"
                }
            );
            l_meses.Add(
                new SelectListItem()
                {
                    Text = "Junio",
                    Value = "6"
                }
            );
            l_meses.Add(
                new SelectListItem()
                {
                    Text = "Julio",
                    Value = "7"
                }
            );
            l_meses.Add(
                new SelectListItem()
                {
                    Text = "Agosto",
                    Value = "8"
                }
            );
            l_meses.Add(
                new SelectListItem()
                {
                    Text = "Setiembre",
                    Value = "9"
                }
            );
            l_meses.Add(
                new SelectListItem()
                {
                    Text = "Octubre",
                    Value = "10"
                }
            );
            l_meses.Add(
                new SelectListItem()
                {
                    Text = "Noviembre",
                    Value = "11"
                }
            );
            l_meses.Add(
                new SelectListItem()
                {
                    Text = "Diciembre",
                    Value = "12"
                }
            );
            return l_meses;
        }
    }
}
