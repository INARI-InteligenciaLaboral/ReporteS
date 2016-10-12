using Reporte.Models;
using System;
using System.Data;
using System.Data.SqlClient;

namespace Reporte.SqlClass
{
    public class Sqlsop
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
                    string mCommand = "select emprIDEmpr,emprRazonSocial from genEmpresas";
                    var mAdapter = new SqlCommand(mCommand, mConexion);
                    mEmpresas.Load(mAdapter.ExecuteReader());
                    mConexion.Close();
                }
            }
            catch
            { }
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
                    string mCommand = "select pcalIDPcal, unneDescripcion from genUnidadesNegocio Inner join nomPlantillasCalculo on unneIDUnne = pcalIDUnne where unneIDEmpr = " + mEmpresa;
                    var mAdapter = new SqlCommand(mCommand, mConexion);
                    mProcesos.Load(mAdapter.ExecuteReader());
                    mConexion.Close();
                }
            }
            catch
            { }
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
                    var cmd = new SqlCommand(cmdText: "P_T_MenSop", connection: mConexion);
                    cmd.CommandType = CommandType.StoredProcedure;
                    cmd.CommandTimeout = 0;
                    cmd.Parameters.Add(parameterName: "@m_empr", sqlDbType: SqlDbType.VarChar, size: 320).Value = mParametros.SelectedEmpresas;
                    cmd.Parameters.Add(parameterName: "@m_mes", sqlDbType: SqlDbType.VarChar, size: 320).Value = mParametros.SelectedMeses;
                    cmd.Parameters.Add(parameterName: "@m_ano", sqlDbType: SqlDbType.VarChar, size: 320).Value = mParametros.SelectedAnos;
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
                    var cmd = new SqlCommand(cmdText: "P_T_ImpSop", connection: mConexion);
                    cmd.CommandTimeout = 0;
                    cmd.CommandType = CommandType.StoredProcedure;
                    cmd.Parameters.Add(parameterName: "@m_procesos", sqlDbType: SqlDbType.VarChar, size: 320).Value = mParametros.SelectedMeses;
                    cmd.Parameters.Add(parameterName: "@m_empr", sqlDbType: SqlDbType.VarChar, size: 320).Value = mParametros.SelectedEmpresas;
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
                    string mCommand = "select emprIDEmpr,emprRazonSocial from genEmpresas where emprIDEmpr = " + mParametros;
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
                    string mCommand = "select pcalIDPcal, periFecIni, periFecFin from genUnidadesNegocio Inner join nomPlantillasCalculo on unneIDUnne = pcalIDUnne Inner join nomPeriodos on periIDPcal = pcalIDPcal where periIDPcal = '" + mParametros.SelectedProcesos + "' and periIDPeri = '" + periodo.ToString() + "'";
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
                    string mCommand = "select 'sopa' + emprIDEmpr,emprRazonSocial from genEmpresas";
                    var mAdapter = new SqlCommand(mCommand, mConexion);
                    mEmpresas.Load(mAdapter.ExecuteReader());
                    mConexion.Close();
                }
            }
            catch
            { }
            return mEmpresas;
        }

        public static DataTable ObtenerProInf(string mEmpresa)
        {
            string mCadena = "Persist Security Info=False;";
            var mProcesos = new DataTable();
            try
            {
                using (SqlConnection mConexion = new SqlConnection(mCadena))
                {
                    mConexion.Open();
                    string mCommand = "select pcalIDPcal, unneDescripcion from genUnidadesNegocio Inner join nomPlantillasCalculo on unneIDUnne = pcalIDUnne where unneIDEmpr = " + mEmpresa;
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
                    var cmd = new SqlCommand(cmdText: "P_T_SopInfMen", connection: mConexion);
                    cmd.CommandTimeout = 0;
                    cmd.CommandType = CommandType.StoredProcedure;
                    cmd.Parameters.Add(parameterName: "@m_ano", sqlDbType: SqlDbType.VarChar, size: 4).Value = mParametros.SelectedAnosMen;
                    cmd.Parameters.Add(parameterName: "@m_mes", sqlDbType: SqlDbType.VarChar, size: 10).Value = mParametros.SelectedMeses;
                    cmd.Parameters.Add(parameterName: "@m_procesos", sqlDbType: SqlDbType.VarChar, size: 320).Value = mParametros.SelectedProcesos;
                    cmd.Parameters.Add(parameterName: "@m_fechamax", sqlDbType: SqlDbType.Date).Value = Convert.ToDateTime(mParametros.SelectedAnosMen + "/" + mParametros.SelectedMeses + "/28");
                    adp = new SqlDataAdapter(cmd);
                    ds = new DataSet();
                    adp.Fill(ds);
                    dt = ds.Tables[0];
                    mConexion.Close();
                }
            }
            catch
            {
                
            }
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
                    var cmd = new SqlCommand(cmdText: "P_T_SopInfMen", connection: mConexion);
                    cmd.CommandTimeout = 0;
                    cmd.CommandType = CommandType.StoredProcedure;
                    cmd.Parameters.Add(parameterName: "@m_ano", sqlDbType: SqlDbType.VarChar, size: 4).Value = mParametros.SelectedAnos;
                    cmd.Parameters.Add(parameterName: "@m_mes", sqlDbType: SqlDbType.VarChar, size: 10).Value = mParametros.SelectedBimestre;
                    cmd.Parameters.Add(parameterName: "@m_procesos", sqlDbType: SqlDbType.VarChar, size: 320).Value = mParametros.SelectedProcesos;
                    if(mParametros.SelectedBimestre.Equals(value:"1,2"))
                        cmd.Parameters.Add(parameterName: "@m_fechamax", sqlDbType: SqlDbType.Date).Value = Convert.ToDateTime(mParametros.SelectedAnos + "/02/28");
                    else if (mParametros.SelectedBimestre.Equals(value: "3,4"))
                        cmd.Parameters.Add(parameterName: "@m_fechamax", sqlDbType: SqlDbType.Date).Value = Convert.ToDateTime(mParametros.SelectedAnos + "/04/28");
                    else if (mParametros.SelectedBimestre.Equals(value: "5,6"))
                        cmd.Parameters.Add(parameterName: "@m_fechamax", sqlDbType: SqlDbType.Date).Value = Convert.ToDateTime(mParametros.SelectedAnos + "/06/28");
                    else if (mParametros.SelectedBimestre.Equals(value: "7,8"))
                        cmd.Parameters.Add(parameterName: "@m_fechamax", sqlDbType: SqlDbType.Date).Value = Convert.ToDateTime(mParametros.SelectedAnos + "/08/28");
                    else if (mParametros.SelectedBimestre.Equals(value: "9,10"))
                        cmd.Parameters.Add(parameterName: "@m_fechamax", sqlDbType: SqlDbType.Date).Value = Convert.ToDateTime(mParametros.SelectedAnos + "/10/28");
                    else if (mParametros.SelectedBimestre.Equals(value: "11,12"))
                        cmd.Parameters.Add(parameterName: "@m_fechamax", sqlDbType: SqlDbType.Date).Value = Convert.ToDateTime(mParametros.SelectedAnos + "/12/28");
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
        public static DataTable GenerarReporteInfPer(CascadingDropdownsModel mParametros)
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
                    var cmd = new SqlCommand(cmdText: "P_T_SopInfPer", connection: mConexion);
                    cmd.CommandTimeout = 0;
                    cmd.CommandType = CommandType.StoredProcedure;
                    cmd.Parameters.Add(parameterName: "@m_perini", sqlDbType: SqlDbType.VarChar, size: 16).Value = mParametros.PerIni;
                    if(String.IsNullOrEmpty(mParametros.PerFin))
                        cmd.Parameters.Add(parameterName: "@m_perfin", sqlDbType: SqlDbType.VarChar, size: 16).Value = mParametros.PerIni;
                    else
                        cmd.Parameters.Add(parameterName: "@m_perfin", sqlDbType: SqlDbType.VarChar, size: 16).Value = mParametros.PerFin;
                    cmd.Parameters.Add(parameterName: "@m_procesos", sqlDbType: SqlDbType.VarChar, size: 320).Value = mParametros.SelectedProcesos;
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
        public static DataTable GenerarReporteFonPer(CascadingDropdownsModel mParametros)
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
                    var cmd = new SqlCommand(cmdText: "P_T_SopFonPer", connection: mConexion);
                    cmd.CommandTimeout = 0;
                    cmd.CommandType = CommandType.StoredProcedure;
                    cmd.Parameters.Add(parameterName: "@m_perini", sqlDbType: SqlDbType.VarChar, size: 16).Value = mParametros.PerIni;
                    if (String.IsNullOrEmpty(mParametros.PerFin))
                        cmd.Parameters.Add(parameterName: "@m_perfin", sqlDbType: SqlDbType.VarChar, size: 16).Value = mParametros.PerIni;
                    else
                        cmd.Parameters.Add(parameterName: "@m_perfin", sqlDbType: SqlDbType.VarChar, size: 16).Value = mParametros.PerFin;
                    cmd.Parameters.Add(parameterName: "@m_procesos", sqlDbType: SqlDbType.VarChar, size: 320).Value = mParametros.SelectedProcesos;
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
                    var cmd = new SqlCommand(cmdText: "P_T_SopFonMen", connection: mConexion);
                    cmd.CommandTimeout = 0;
                    cmd.CommandType = CommandType.StoredProcedure;
                    cmd.Parameters.Add(parameterName: "@m_ano", sqlDbType: SqlDbType.VarChar, size: 4).Value = mParametros.SelectedAnosMen;
                    cmd.Parameters.Add(parameterName: "@m_mes", sqlDbType: SqlDbType.VarChar, size: 10).Value = mParametros.SelectedMeses;
                    cmd.Parameters.Add(parameterName: "@m_procesos", sqlDbType: SqlDbType.VarChar, size: 320).Value = mParametros.SelectedProcesos;
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
                    var cmd = new SqlCommand(cmdText: "P_T_SopFonMen", connection: mConexion);
                    cmd.CommandTimeout = 0;
                    cmd.CommandType = CommandType.StoredProcedure;
                    cmd.Parameters.Add(parameterName: "@m_ano", sqlDbType: SqlDbType.VarChar, size: 4).Value = mParametros.SelectedAnos;
                    cmd.Parameters.Add(parameterName: "@m_mes", sqlDbType: SqlDbType.VarChar, size: 10).Value = mParametros.SelectedBimestre;
                    cmd.Parameters.Add(parameterName: "@m_procesos", sqlDbType: SqlDbType.VarChar, size: 320).Value = mParametros.SelectedProcesos;
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