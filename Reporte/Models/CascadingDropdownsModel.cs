using System.Collections.Generic;
using System.Web.Mvc;

namespace Reporte.Models
{
    public class CascadingDropdownsModel
    {
        public IList<SelectListItem> Base { get; set; }

        public IList<SelectListItem> Empresas { get; set; }

        public IList<SelectListItem> Reportes { get; set; }

        public string SelectedReporte { get; set; }

        public string SelectedBase { get; set; }

        public string SelectedEmpresas { get; set; }

        public string SelectedProcesos { get; set; }

        public string SelectedAnos { get; set; }

        public string SelectedAnosMen { get; set; }

        public string SelectedMeses { get; set; }

        public string SelectedBimestre { get; set; }

        public IList<SelectListItem> Anos { get; set; }

        public IList<SelectListItem> Meses { get; set; }

        public IList<SelectListItem> Bimestre { get; set; }

        public string typeCheck { get; set; }

        public string PerMen { get; set; }

        public string PerIni { get; set; }

        public string PerFin { get; set; }
    }
}
