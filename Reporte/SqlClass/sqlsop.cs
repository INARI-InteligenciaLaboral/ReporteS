using Reporte.Models;
using System;
using System.Collections.Generic;
using System.Data;
using System.Data.SqlClient;
using System.Linq;
using System.Web;

namespace Reporte.SqlClass
{
    public class sqlsop
    {
        public static DataTable ObtenerEmp()
        {
            string m_cadena = "Persist Security Info=False;User ID=sa;Initial Catalog=dbDatosNominaTest;";
            DataTable m_empresas = new DataTable();
            try
            {
                using (SqlConnection m_conexion = new SqlConnection(m_cadena))
                {
                    m_conexion.Open();
                    string m_command = "select emprIDEmpr,emprRazonSocial from genEmpresas";
                    SqlCommand m_adapter = new SqlCommand(m_command, m_conexion);
                    m_empresas.Load(m_adapter.ExecuteReader());
                    m_conexion.Close();
                }
            }
            catch
            { }
            return m_empresas;
        }

        public static DataTable ObtenerPro(string m_empresa)
        {
            string m_cadena = "Persist Security Info=False;User ID=sa;Initial Catalog=dbDatosNominaTest;";
            DataTable m_Procesos = new DataTable();
            try
            {
                using (SqlConnection m_conexion = new SqlConnection(m_cadena))
                {
                    m_conexion.Open();
                    string m_command = "select pcalIDPcal, unneDescripcion from genUnidadesNegocio Inner join nomPlantillasCalculo on unneIDUnne = pcalIDUnne where unneIDEmpr = " + m_empresa;
                    SqlCommand m_adapter = new SqlCommand(m_command, m_conexion);
                    m_Procesos.Load(m_adapter.ExecuteReader());
                    m_conexion.Close();
                }
            }
            catch
            { }
            return m_Procesos;
        }
        public static DataTable GenerarReporte(CascadingDropdownsModel m_Parametros)
        {
            SqlDataAdapter adp;
            DataSet ds = new DataSet();
            DataTable dt = new DataTable();
            string m_cadena = "Persist Security Info=False;User ID=sa;Initial Catalog=dbDatosNominaTest;";
            try
            {
                using (SqlConnection m_conexion = new SqlConnection(m_cadena))
                {
                    m_conexion.Open();
                    SqlCommand cmd = new SqlCommand("P_T_MenSop", m_conexion);
                    cmd.CommandType = CommandType.StoredProcedure;
                    cmd.CommandTimeout = 0;
                    cmd.Parameters.Add("@m_empr", SqlDbType.VarChar, 320).Value = m_Parametros.SelectedEmpresas;
                    cmd.Parameters.Add("@m_mes", SqlDbType.VarChar, 320).Value = m_Parametros.SelectedMeses;
                    cmd.Parameters.Add("@m_ano", SqlDbType.VarChar, 320).Value = m_Parametros.SelectedAnos;
                    adp = new SqlDataAdapter(cmd);
                    ds = new DataSet();
                    adp.Fill(ds);
                    dt = ds.Tables[0];
                    m_conexion.Close();
                }
            }
            catch
            { }
            return dt;
        }
        public static DataTable GenerarReporteImp(CascadingDropdownsModel m_Parametros)
        {
            SqlDataAdapter adp;
            DataSet ds = new DataSet();
            DataTable dt = new DataTable();
            string m_cadena = "Persist Security Info=False;User ID=sa;Initial Catalog=dbDatosNominaTest;";
            try
            {
                using (SqlConnection m_conexion = new SqlConnection(m_cadena))
                {
                    m_conexion.Open();
                    SqlCommand cmd = new SqlCommand("P_T_ImpSop", m_conexion);
                    cmd.CommandTimeout = 0;
                    cmd.CommandType = CommandType.StoredProcedure;
                    cmd.Parameters.Add("@m_procesos", SqlDbType.VarChar, 320).Value = m_Parametros.SelectedMeses;
                    cmd.Parameters.Add("@m_empr", SqlDbType.VarChar, 320).Value = m_Parametros.SelectedEmpresas;
                    adp = new SqlDataAdapter(cmd);
                    ds = new DataSet();
                    adp.Fill(ds);
                    dt = ds.Tables[0];
                    m_conexion.Close();
                }
            }
            catch
            { }
            return dt;
        }
        public static string Empresa_Title(string m_Parametros)
        {
            string m_cadena = "Persist Security Info=False;User ID=sa;Initial Catalog=dbDatosNominaTest;";
            string titlemens = "";
            DataTable m_empresas = new DataTable();
            try
            {
                using (SqlConnection m_conexion = new SqlConnection(m_cadena))
                {
                    m_conexion.Open();
                    string m_command = "select emprIDEmpr,emprRazonSocial from genEmpresas where emprIDEmpr = " + m_Parametros;
                    SqlCommand m_adapter = new SqlCommand(m_command, m_conexion);
                    m_empresas.Load(m_adapter.ExecuteReader());
                    m_conexion.Close();
                }
                foreach (DataRow Row in m_empresas.Rows)
                {
                    titlemens = Row[1].ToString();
                }
            }
            catch
            { }
            return titlemens;
        }
        public static string ProcesosTitle(CascadingDropdownsModel m_Parametros, int periodo)
        {
            string m_cadena = "Persist Security Info=False;User ID=sa;Initial Catalog=dbDatosNominaTest;";
            string titlemens = "";
            DataTable m_empresas = new DataTable();
            try
            {
                using (SqlConnection m_conexion = new SqlConnection(m_cadena))
                {
                    m_conexion.Open();
                    string m_command = "select pcalIDPcal, periFecIni, periFecFin from genUnidadesNegocio Inner join nomPlantillasCalculo on unneIDUnne = pcalIDUnne Inner join nomPeriodos on periIDPcal = pcalIDPcal where periIDPcal = '" + m_Parametros.SelectedProcesos + "' and periIDPeri = '" + periodo.ToString() + "'";
                    SqlCommand m_adapter = new SqlCommand(m_command, m_conexion);
                    m_empresas.Load(m_adapter.ExecuteReader());
                    m_conexion.Close();
                }
                foreach (DataRow Row in m_empresas.Rows)
                {
                    titlemens = "Reporte del Proceso " + m_Parametros.SelectedProcesos.ToString() + " del período " + Row[0].ToString() + " - " + Convert.ToDateTime(Row[1]).ToString("dd-MMM-yyyy") + " al " + Convert.ToDateTime(Row[2]).ToString("dd-MMM-yyyy");
                }
            }
            catch
            { }
            return titlemens;
        }
    }
}