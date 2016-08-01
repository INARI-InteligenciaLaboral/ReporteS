using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Web.Mvc;

namespace Reporte.Models
{
    public class CascadingDropdownsModel
    {
        public IList<SelectListItem> Empresas { get; set; }

        public string SelectedEmpresas { get; set; }

        public string SelectedProcesos { get; set; }

        public string SelectedAnos { get; set; }

        public string SelectedMeses { get; set; }

        public IList<SelectListItem> Anos { get; set; }

        public IList<SelectListItem> Meses { get; set; }
    }
}
