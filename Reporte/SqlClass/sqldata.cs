using System;
using System.Data;
using System.Data.SqlClient;
using Reporte.Models;

namespace Reporte.SqlClass
{
    public class Sqldata
    {
        public static DataTable ObtenerEmp()
        {
            string mCadena = "Persist Security Info=False;";
            var mEmpresas = new DataTable();
            try
            {
                using (SqlConnection mConexion = new SqlConnection(mCadena))
                {
                    mConexion.Open();
                    string mCommand = "SELECT cia_keycia AS KeyEmp, cia_descia AS Empresa FROM nmlocias";
                    var mAdapter = new SqlCommand(mCommand, mConexion);
                    mEmpresas.Load(mAdapter.ExecuteReader());
                    mConexion.Close();
                }
            }
            catch
            {}
            return mEmpresas;
        }
        public static DataTable ObtenerPro(string mEmpresa)
        {
            string mCadena = "Persist Security Info=False;";
            var mProcesos = new DataTable();
            try
            {
                using (SqlConnection mConexion = new SqlConnection(mCadena))
                {
                    mConexion.Open();
                    string mCommand = "SELECT pro_keypro AS KeyPro, pro_despro AS DesPro FROM nmloproc WHERE pro_keycia = " + mEmpresa + " ORDER BY KeyPro";
                    var mAdapter = new SqlCommand(mCommand, mConexion);
                    mProcesos.Load(mAdapter.ExecuteReader());
                    mConexion.Close();
                }
            }
            catch 
            {}
            return mProcesos;
        }
        public static DataTable GenerarReporte(CascadingDropdownsModel mParametros)
        {
            SqlDataAdapter adp;
            var ds = new DataSet();
            var dt = new DataTable();
            string mCadena = "Persist Security Info=False;";
            try
            {
                using (SqlConnection mConexion = new SqlConnection(mCadena))
                {
                    mConexion.Open();
                    var cmd = new SqlCommand(cmdText: "P_T_Mensual", connection: mConexion);
                    cmd.CommandType = CommandType.StoredProcedure;
                    cmd.CommandTimeout = 0;
                    cmd.Parameters.Add(parameterName: "@m_procesos", sqlDbType: SqlDbType.NVarChar, size: 320).Value = mParametros.SelectedProcesos;
                    cmd.Parameters.Add(parameterName: "@m_mes", sqlDbType: SqlDbType.NVarChar, size: 320).Value = mParametros.SelectedMeses;
                    cmd.Parameters.Add(parameterName: "@m_ano", sqlDbType: SqlDbType.NVarChar, size: 320).Value = mParametros.SelectedAnos;
                    adp = new SqlDataAdapter(cmd);
                    ds = new DataSet();
                    adp.Fill(ds);
                    dt = ds.Tables[0];
                    mConexion.Close();
                }
            }
            catch 
            { }
            return dt;
        }
        public static DataTable GenerarReporteImp(CascadingDropdownsModel mParametros)
        {
            SqlDataAdapter adp;
            var ds = new DataSet();
            var dt = new DataTable();
            string mCadena = "Persist Security Info=False;";
            try
            {
                using (SqlConnection mConexion = new SqlConnection(mCadena))
                {
                    mConexion.Open();
                    var cmd = new SqlCommand(cmdText: "P_T_Importe", connection: mConexion);
                    cmd.CommandTimeout = 0;
                    cmd.CommandType = CommandType.StoredProcedure;
                    cmd.Parameters.Add(parameterName: "@m_procesos", sqlDbType: SqlDbType.NVarChar, size: 320).Value = mParametros.SelectedProcesos;
                    cmd.Parameters.Add(parameterName: "@m_mes", sqlDbType: SqlDbType.NVarChar, size: 320).Value = mParametros.SelectedMeses;
                    adp = new SqlDataAdapter(cmd);
                    ds = new DataSet();
                    adp.Fill(ds);
                    dt = ds.Tables[0];
                    mConexion.Close();
                }
            }
            catch 
            { }
            return dt;
        }
        public static string Empresa_Title(string mParametros)
        {
            string mCadena = "Persist Security Info=False;";
            string titlemens = "";
            var mEmpresas = new DataTable();
            try
            {
                using (SqlConnection mConexion = new SqlConnection(mCadena))
                {
                    mConexion.Open();
                    string mCommand = "SELECT cia_keycia AS KeyEmp, cia_descia AS Empresa FROM nmlocias Where cia_keycia = " + mParametros;
                    var mAdapter = new SqlCommand(mCommand, mConexion);
                    mEmpresas.Load(mAdapter.ExecuteReader());
                    mConexion.Close();
                }
                foreach (DataRow Row in mEmpresas.Rows)
                {
                    titlemens = Row[1].ToString();
                }
            }
            catch 
            { }
            return titlemens;
        }
        public static string ProcesosTitle(CascadingDropdownsModel mParametros, int periodo)
        {
            string mCadena = "Persist Security Info=False;";
            string titlemens = "";
            var mEmpresas = new DataTable();
            try
            {
                using (SqlConnection mConexion = new SqlConnection(mCadena))
                {
                    mConexion.Open();
                    string mCommand = "select substring(per_keyper,6,7), per_fecini, per_fecfin from nmloperi where per_keypro = " + mParametros.SelectedProcesos + " and per_keyper = " + periodo.ToString();
                    var mAdapter = new SqlCommand(mCommand, mConexion);
                    mEmpresas.Load(mAdapter.ExecuteReader());
                    mConexion.Close();
                }
                foreach (DataRow Row in mEmpresas.Rows)
                {
                    titlemens = "Reporte del Proceso " + mParametros.SelectedProcesos.ToString() + " del período " + Row[0].ToString() + " - " + Convert.ToDateTime(Row[1]).ToString(format: "dd-MMM-yyyy")+ " al " + Convert.ToDateTime(Row[2]).ToString(format: "dd-MMM-yyyy");
                }
            }
            catch 
            { }
            return titlemens;
        }
        public static DataTable ObtenerEmpInf()
        {
            string mCadena = "Persist Security Info=False;";
            var mEmpresas = new DataTable();
            try
            {
                using (SqlConnection mConexion = new SqlConnection(mCadena))
                {
                    mConexion.Open();
                    string mCommand = "SELECT 'siap' + cia_keycia AS KeyEmp, cia_descia AS Empresa FROM nmlocias";
                    var mAdapter = new SqlCommand(mCommand, mConexion);
                    mEmpresas.Load(mAdapter.ExecuteReader());
                    mConexion.Close();
                }
            }
            catch
            { }
            return mEmpresas;
        }
        public static DataTable ObtenerProinf(string mEmpresa)
        {
            string mCadena = "Persist Security Info=False;";
            var mProcesos = new DataTable();
            try
            {
                using (SqlConnection mConexion = new SqlConnection(mCadena))
                {
                    mConexion.Open();
                    string mCommand = "SELECT pro_keypro AS KeyPro, CAST(pro_keypro AS CHAR(4)) + '-' + pro_despro AS DesPro FROM nmloproc WHERE pro_keycia = " + mEmpresa + " ORDER BY KeyPro";
                    var mAdapter = new SqlCommand(mCommand, mConexion);
                    mProcesos.Load(mAdapter.ExecuteReader());
                    mConexion.Close();
                }
            }
            catch
            { }
            return mProcesos;
        }
        public static DataTable GenerarReporteInfMen(CascadingDropdownsModel mParametros)
        {
            SqlDataAdapter adp;
            var ds = new DataSet();
            var dt = new DataTable();
            string mCadena = "Persist Security Info=False;";
            try
            {
                using (SqlConnection mConexion = new SqlConnection(mCadena))
                {
                    mConexion.Open();
                    var cmd = new SqlCommand(cmdText: "P_T_SiapInfMen", connection: mConexion);
                    cmd.CommandTimeout = 0;
                    cmd.CommandType = CommandType.StoredProcedure;
                    cmd.Parameters.Add(parameterName: "@m_anos", sqlDbType: SqlDbType.VarChar, size: 4).Value = mParametros.SelectedAnosMen;
                    cmd.Parameters.Add(parameterName: "@m_mes", sqlDbType: SqlDbType.VarChar, size: 7).Value = mParametros.SelectedMeses;
                    cmd.Parameters.Add(parameterName: "@m_proceso", sqlDbType: SqlDbType.VarChar, size: 320).Value = mParametros.SelectedProcesos;
                    adp = new SqlDataAdapter(cmd);
                    ds = new DataSet();
                    adp.Fill(ds);
                    dt = ds.Tables[0];
                    mConexion.Close();
                }
            }
            catch
            { }
            return dt;
        }
        public static DataTable GenerarReporteInfBim(CascadingDropdownsModel mParametros)
        {
            SqlDataAdapter adp;
            var ds = new DataSet();
            var dt = new DataTable();
            string mCadena = "Persist Security Info=False;";
            try
            {
                using (SqlConnection mConexion = new SqlConnection(mCadena))
                {
                    mConexion.Open();
                    var cmd = new SqlCommand(cmdText: "P_T_SiapInfMen", connection: mConexion);
                    cmd.CommandTimeout = 0;
                    cmd.CommandType = CommandType.StoredProcedure;
                    cmd.Parameters.Add(parameterName: "@m_anos", sqlDbType: SqlDbType.VarChar, size: 4).Value = mParametros.SelectedAnos;
                    cmd.Parameters.Add(parameterName: "@m_mes", sqlDbType: SqlDbType.VarChar, size: 7).Value = mParametros.SelectedBimestre;
                    cmd.Parameters.Add(parameterName: "@m_proceso", sqlDbType: SqlDbType.VarChar, size: 320).Value = mParametros.SelectedProcesos;
                    adp = new SqlDataAdapter(cmd);
                    ds = new DataSet();
                    adp.Fill(ds);
                    dt = ds.Tables[0];
                    mConexion.Close();
                }
            }
            catch
            { }
            return dt;
        }
        public static DataTable GenerarReporteFonMen(CascadingDropdownsModel mParametros)
        {
            SqlDataAdapter adp;
            var ds = new DataSet();
            var dt = new DataTable();
            string mCadena = "Persist Security Info=False;";
            try
            {
                using (SqlConnection mConexion = new SqlConnection(mCadena))
                {
                    mConexion.Open();
                    var cmd = new SqlCommand(cmdText: "P_T_SiapFonMen", connection: mConexion);
                    cmd.CommandTimeout = 0;
                    cmd.CommandType = CommandType.StoredProcedure;
                    cmd.Parameters.Add(parameterName: "@m_anos", sqlDbType: SqlDbType.VarChar, size: 4).Value = mParametros.SelectedAnosMen;
                    cmd.Parameters.Add(parameterName: "@m_mes", sqlDbType: SqlDbType.VarChar, size: 7).Value = mParametros.SelectedMeses;
                    cmd.Parameters.Add(parameterName: "@m_proceso", sqlDbType: SqlDbType.VarChar, size: 320).Value = mParametros.SelectedProcesos;
                    adp = new SqlDataAdapter(cmd);
                    ds = new DataSet();
                    adp.Fill(ds);
                    dt = ds.Tables[0];
                    mConexion.Close();
                }
            }
            catch
            { }
            return dt;
        }
        public static DataTable GenerarReporteFonBim(CascadingDropdownsModel mParametros)
        {
            SqlDataAdapter adp;
            var ds = new DataSet();
            var dt = new DataTable();
            string mCadena = "Persist Security Info=False;";
            try
            {
                using (SqlConnection mConexion = new SqlConnection(mCadena))
                {
                    mConexion.Open();
                    var cmd = new SqlCommand(cmdText: "P_T_SiapFonMen", connection: mConexion);
                    cmd.CommandTimeout = 0;
                    cmd.CommandType = CommandType.StoredProcedure;
                    cmd.Parameters.Add(parameterName: "@m_anos", sqlDbType: SqlDbType.VarChar, size: 4).Value = mParametros.SelectedAnos;
                    cmd.Parameters.Add(parameterName: "@m_mes", sqlDbType: SqlDbType.VarChar, size: 7).Value = mParametros.SelectedBimestre;
                    cmd.Parameters.Add(parameterName: "@m_proceso", sqlDbType: SqlDbType.VarChar, size: 320).Value = mParametros.SelectedProcesos;
                    adp = new SqlDataAdapter(cmd);
                    ds = new DataSet();
                    adp.Fill(ds);
                    dt = ds.Tables[0];
                    mConexion.Close();
                }
            }
            catch
            { }
            return dt;
        }
    }
}