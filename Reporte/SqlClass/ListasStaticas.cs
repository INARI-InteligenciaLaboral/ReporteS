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
            for(int x = DateTime.Now.Year; x >= 2010; x--)
            {
                l_anos.Add(
                new SelectListItem()
                {
                    Text = x.ToString(),
                    Value = x.ToString()
                });
            }
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
                    Text = "Septiembre",
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
        public static List<SelectListItem> ObtenerMesesSop()
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
                    Value = "01"
                }
            );
            l_meses.Add(
                new SelectListItem()
                {
                    Text = "Febrero",
                    Value = "02"
                }
            );
            l_meses.Add(
                new SelectListItem()
                {
                    Text = "Marzo",
                    Value = "03"
                }
            );
            l_meses.Add(
                new SelectListItem()
                {
                    Text = "Abril",
                    Value = "04"
                }
            );
            l_meses.Add(
                new SelectListItem()
                {
                    Text = "Mayo",
                    Value = "05"
                }
            );
            l_meses.Add(
                new SelectListItem()
                {
                    Text = "Junio",
                    Value = "06"
                }
            );
            l_meses.Add(
                new SelectListItem()
                {
                    Text = "Julio",
                    Value = "07"
                }
            );
            l_meses.Add(
                new SelectListItem()
                {
                    Text = "Agosto",
                    Value = "08"
                }
            );
            l_meses.Add(
                new SelectListItem()
                {
                    Text = "Septiembre",
                    Value = "09"
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
        public static List<SelectListItem> BimestresSop()
        {
            List<SelectListItem> l_meses = new List<SelectListItem>();
            l_meses.Add(
                new SelectListItem()
                {
                    Text = "Primero",
                    Value = "1,2"
                }
            );
            l_meses.Add(
                new SelectListItem()
                {
                    Text = "Segundo",
                    Value = "3,4"
                }
            );
            l_meses.Add(
                new SelectListItem()
                {
                    Text = "Tercero",
                    Value = "5,6"
                }
            );
            l_meses.Add(
                new SelectListItem()
                {
                    Text = "Cuarto",
                    Value = "7,8"
                }
            );
            l_meses.Add(
                new SelectListItem()
                {
                    Text = "Quinto",
                    Value = "9,10"
                }
            );
            l_meses.Add(
                new SelectListItem()
                {
                    Text = "Sexto",
                    Value = "11"
                }
            );
            return l_meses;
        }
    }
}
