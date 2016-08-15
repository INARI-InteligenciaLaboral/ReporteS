using System;
using System.Data;
using System.Data.SqlClient;
using Reporte.Models;

namespace Reporte.SqlClass
{
    public class sqldata
    {
        public static DataTable ObtenerEmp()
        {
            string m_cadena = "Persist Security Info=False;User ID=usuario_siap;Initial Catalog=SIAP;";
            DataTable m_empresas = new DataTable();
            try
            {
                using (SqlConnection m_conexion = new SqlConnection(m_cadena))
                {
                    m_conexion.Open();
                    string m_command = "SELECT cia_keycia AS KeyEmp, cia_descia AS Empresa FROM nmlocias";
                    SqlCommand m_adapter = new SqlCommand(m_command, m_conexion);
                    m_empresas.Load(m_adapter.ExecuteReader());
                    m_conexion.Close();
                }
            }
            catch
            {}
            return m_empresas;
        }

        public static DataTable ObtenerPro(string m_empresa)
        {
            string m_cadena = "Persist Security Info=False;User ID=usuario_siap;Initial Catalog=SIAP;";
            DataTable m_Procesos = new DataTable();
            try
            {
                using (SqlConnection m_conexion = new SqlConnection(m_cadena))
                {
                    m_conexion.Open();
                    string m_command = "SELECT pro_keypro AS KeyPro, pro_despro AS DesPro FROM nmloproc WHERE pro_keycia = " + m_empresa + " ORDER BY KeyPro";
                    SqlCommand m_adapter = new SqlCommand(m_command, m_conexion);
                    m_Procesos.Load(m_adapter.ExecuteReader());
                    m_conexion.Close();
                }
            }
            catch 
            {}
            return m_Procesos;
        }
        public static DataTable GenerarReporte(CascadingDropdownsModel m_Parametros)
        {
            SqlDataAdapter adp;
            DataSet ds = new DataSet();
            DataTable dt = new DataTable();
            string m_cadena = "Persist Security Info=False;User ID=usuario_siap;Initial Catalog=SIAP;";
            try
            {
                using (SqlConnection m_conexion = new SqlConnection(m_cadena))
                {
                    m_conexion.Open();
                    SqlCommand cmd = new SqlCommand("P_T_Mensual", m_conexion);
                    cmd.CommandType = CommandType.StoredProcedure;
                    cmd.CommandTimeout = 0;
                    cmd.Parameters.Add("@m_procesos", SqlDbType.NVarChar, 320).Value = m_Parametros.SelectedProcesos;
                    cmd.Parameters.Add("@m_mes", SqlDbType.NVarChar, 320).Value = m_Parametros.SelectedMeses;
                    cmd.Parameters.Add("@m_ano", SqlDbType.NVarChar, 320).Value = m_Parametros.SelectedAnos;
                    adp = new SqlDataAdapter(cmd);
                    ds = new DataSet();
                    adp.Fill(ds);
                    dt = ds.Tables[0];
                    m_conexion.Close();
                }
            }
            catch (Exception ex)
            { }
            return dt;
        }
        public static DataTable GenerarReporteImp(CascadingDropdownsModel m_Parametros)
        {
            SqlDataAdapter adp;
            DataSet ds = new DataSet();
            DataTable dt = new DataTable();
            string m_cadena = "Persist Security Info=False;User ID=usuario_siap;Initial Catalog=SIAP;";
            try
            {
                using (SqlConnection m_conexion = new SqlConnection(m_cadena))
                {
                    m_conexion.Open();
                    SqlCommand cmd = new SqlCommand("P_T_Importe", m_conexion);
                    cmd.CommandTimeout = 0;
                    cmd.CommandType = CommandType.StoredProcedure;
                    cmd.Parameters.Add("@m_procesos", SqlDbType.NVarChar, 320).Value = m_Parametros.SelectedProcesos;
                    cmd.Parameters.Add("@m_mes", SqlDbType.NVarChar, 320).Value = m_Parametros.SelectedMeses;
                    adp = new SqlDataAdapter(cmd);
                    ds = new DataSet();
                    adp.Fill(ds);
                    dt = ds.Tables[0];
                    m_conexion.Close();
                }
            }
            catch (Exception ex)
            { }
            return dt;
        }
        public static string Empresa_Title(string m_Parametros)
        {
            string m_cadena = "Persist Security Info=False;User ID=usuario_siap;Initial Catalog=SIAP;";
            string titlemens = "";
            DataTable m_empresas = new DataTable();
            try
            {
                using (SqlConnection m_conexion = new SqlConnection(m_cadena))
                {
                    m_conexion.Open();
                    string m_command = "SELECT cia_keycia AS KeyEmp, cia_descia AS Empresa FROM nmlocias Where cia_keycia = " + m_Parametros;
                    SqlCommand m_adapter = new SqlCommand(m_command, m_conexion);
                    m_empresas.Load(m_adapter.ExecuteReader());
                    m_conexion.Close();
                }
                foreach (DataRow Row in m_empresas.Rows)
                {
                    titlemens = Row[1].ToString();
                }
            }
            catch (Exception ex)
            { }
            return titlemens;
        }
        public static string ProcesosTitle(CascadingDropdownsModel m_Parametros, int periodo)
        {
            string m_cadena = "Persist Security Info=False;User ID=usuario_siap;Initial Catalog=SIAP;";
            string titlemens = "";
            DataTable m_empresas = new DataTable();
            try
            {
                using (SqlConnection m_conexion = new SqlConnection(m_cadena))
                {
                    m_conexion.Open();
                    string m_command = "select substring(per_keyper,6,7), per_fecini, per_fecfin from nmloperi where per_keypro = " + m_Parametros.SelectedProcesos + " and per_keyper = " + periodo.ToString();
                    SqlCommand m_adapter = new SqlCommand(m_command, m_conexion);
                    m_empresas.Load(m_adapter.ExecuteReader());
                    m_conexion.Close();
                }
                foreach (DataRow Row in m_empresas.Rows)
                {
                    titlemens = "Reporte del Proceso " + m_Parametros.SelectedProcesos.ToString() + " del período " + Row[0].ToString() + " - " + Convert.ToDateTime(Row[1]).ToString("dd-MMM-yyyy") + " al " + Convert.ToDateTime(Row[2]).ToString("dd-MMM-yyyy");
                }
            }
            catch (Exception ex)
            { }
            return titlemens;
        }
    }
}