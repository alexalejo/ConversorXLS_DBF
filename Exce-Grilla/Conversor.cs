using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using System.Data.OleDb;
using System.Data.Odbc;
namespace Exce_Grilla
{
    public partial class Conversor : Form
    {
        #region Constructor
        public Conversor()
        {
            InitializeComponent();
        }
        #endregion

        #region variables

        DataTable dtHojas;
        OleDbConnection oleConexion = new OleDbConnection();
        string strConexion = string.Empty;
        bool procesando = false;
        string strRutaDbf = string.Empty;
        string strRutaXls = string.Empty;
        string strSelect = string.Empty;
        
            
        #endregion

        #region Eventos
        private void btnExaminar_Click(object sender, EventArgs e)
        {
            ofdArchivo.Filter = "Archivos *.xls|*.xls";// +"Textos|*.txt|" + "Documentos|*.doc|" + "Imagenes|*.bmp;*.jpg;*.gif|" + "Todos los archivos|*.*";
            ofdArchivo.InitialDirectory = System.Configuration.ConfigurationSettings.AppSettings["PathXls"]; 
            ofdArchivo.FileName = string.Empty;
            ofdArchivo.Title = "Buscar archivo Excel";
            if (ofdArchivo.ShowDialog() == DialogResult.OK )
            {
                lRuta.Text =  ofdArchivo.FileName ;
                cmbHojas.Items.Clear();
                limpiarGrilla();
            }
            if (lRuta.Text != string.Empty)
            {
                try
                {
                    strConexion = @"Provider=Microsoft.Jet.OLEDB.4.0;Data Source= " + this.lRuta.Text + ";Extended Properties=Excel 8.0;";
                    oleConexion.ConnectionString = strConexion;
                    oleConexion.Open();
                    dtHojas = oleConexion.GetSchema("TABLES");
                    DataRow[] drRows = dtHojas.Select("TABLE_NAME like '*$' OR TABLE_NAME like '*$'''");
                    //Cargo el combo con nombre de las hojas
                    foreach (DataRow objRow in drRows)
                    {
                        string strSheet = objRow["TABLE_NAME"].ToString();
                        strSheet = strSheet.StartsWith("'") ? strSheet.Remove(0, 1) : strSheet;
                        strSheet = strSheet.IndexOf("$'") > 0 ? strSheet.Remove(strSheet.IndexOf('$'), 2) : strSheet.Remove(strSheet.IndexOf('$'), 1);
                        cmbHojas.Items.Add(strSheet);
                    }
                    cmbHojas.Text = "TARIFA";
                }
                catch (Exception)
                {
                    MessageBox.Show("Error, El archivo excel puede ser que no exista o tenga una estructura diferente.");
                }
                finally
                {
                    if (oleConexion != null && oleConexion.State != ConnectionState.Closed)
                        oleConexion.Close();
                }
            }
        }
        private void btnComprobar_Click(object sender, EventArgs e)
        {
            if (cmbHojas.Items.Count > 0 && dgDatos.DataSource == null)
            {              
                try
                {
                    Cursor.Current = Cursors.WaitCursor;
                    abrirExcel();
                }
                catch (Exception ex)
                {
                    MessageBox.Show(ex.Message);
                    return;
                }
                Cursor.Current = Cursors.Default;
            }
            else if (cmbHojas.Items.Count > 0 && dgDatos.DataSource != null)
                MessageBox.Show("Hoja seleccionada ya se encuentra abierta", "Información");
            else
                MessageBox.Show("No hay abierto ningún archivo excel!", "Información");
        }

        private void btnSalir_Click(object sender, EventArgs e)
        {
            this.Close();
        }
        private void btnprocesar_Click(object sender, EventArgs e)
        {
            
            borrarDbf();
            if (dgDatos.Rows.Count > 0)
               grabarDbf();
            else
               MessageBox.Show("No se encuentran registros a procesar!", "información");
        }

        private void cmbHojas_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (!procesando)
            {
                limpiarGrilla();
                if (oleConexion != null && oleConexion.State != ConnectionState.Closed)
                    oleConexion.Close();
            }
        }
        #endregion

        #region metodos
         
        void borrarDbf()
        {
            abrirdbf("stocks");
            abrirdbf("Articulo");
            dgDatos.DataSource = dsDocumentos.Articulo;
            strRutaDbf = System.Configuration.ConfigurationSettings.AppSettings["PathDbfs"];
            string strSql = string.Empty;// Modo == "DELETE" ? "DELETE FROM Articulo where cref =" : "INSERT"; 
            OdbcConnection cnConexion = null;
            OdbcCommand sComando = null;
            try
            {
                cnConexion = new OdbcConnection("Driver={Microsoft dBASE Driver (*.dbf)};DriverID=277;Dbq=   " + strRutaDbf + ";");
                if (cnConexion.State != ConnectionState.Open)
                    cnConexion.Open();
                for (int i = 0; i < dgDatos.Rows.Count - 1; i++)
                {
                    strSql = "DELETE FROM Articulo where cref =" + " '" + dgDatos.Rows[i].Cells[0].Value.ToString() + "' ";
                    sComando = new OdbcCommand(strSql, cnConexion);
                    sComando.ExecuteNonQuery();
                    strSql = "DELETE FROM stocks where cref =" + " '" + dgDatos.Rows[i].Cells[0].Value.ToString() + "' ";
                    sComando = new OdbcCommand(strSql, cnConexion);
                    sComando.ExecuteNonQuery();
                    sComando = null;
                }
            }
            catch (Exception ex)
            {
                throw new Exception(ex.Message);
            }
            finally
            {
                if (cnConexion != null && cnConexion.State != ConnectionState.Closed)
                {
                    cnConexion.Close();
                    cnConexion.Dispose();
                    cnConexion = null;
                }
                if (sComando != null)
                {
                    sComando.Dispose();
                    sComando = null;
                }
            }
        }

        void grabarDbf()
        {
            strRutaDbf = System.Configuration.ConfigurationSettings.AppSettings["PathDbfs"];
            string strSqlArt = string.Empty;
            string strSqlStoc = string.Empty;
            int totreg = 0;
            OdbcConnection cnConexion = null;
            OdbcCommand sComando = null;
            try
            {
                string cadena = "GEUR86N";
                
                brProgreso.Minimum = 0;
                brProgreso.Maximum = dgDatos.Rows.Count;
                brProgreso.Step = 1;
                byte cont = 0;
                cnConexion = new OdbcConnection("Driver={Microsoft dBASE Driver (*.dbf)};DriverID=277;Dbq=   " + strRutaDbf + ";");
                if (cnConexion.State != ConnectionState.Open)
                    cnConexion.Open();
                Cursor.Current = Cursors.WaitCursor;
                for (int i = 0; i < dgDatos.Rows.Count; i++)
                {
                   if (dgDatos.Rows[i].Cells[1].Value.ToString() == "Referencia" || dgDatos.Rows[i].Cells[2].Value.ToString() == "Descripción")
                        i++;
                    if (dgDatos.Rows[i].Cells[5].Value.ToString() == "PENDIENTE")
                    {
                        if (dgDatos.Rows[i].Cells[0].Value.ToString() == "INSERTAR")
                        {
                            string fami = System.Configuration.ConfigurationSettings.AppSettings["CODFAM"];
                            strSqlArt = "insert into Articulo (CREF,CDETALLE,CCODFAM, NPZOREP ,NCOSTEPROM,NCOSTEDIV, NSTOCKMIN,NSTOCKMAX, CTIPOIVA, NPVP,  NPREMAYOR, " +
                            " NDTO1,  NDTO2, NDTO3,  NDTO4, NDTO5, NDTO6, NPENDSER, NPENDREC , CCODDIV, NENTACU,CCODPRO, LSELECT, NUNIDADES," +
                            " NETIQUETAS,  LINTERNET, LUNIDADES, NTIPOCODE, LKIT, LPESO,NTARA, NBENEFPVP,NBENEFMAY, LPRECPROP1,  LPRECPROP2, NPCONIVA, NPMCONIVA, LCTRLSTOCK, " +
                            " NPREPV,LACTUALIZA,DACTUALIZA, NSALACU, LANTICIPO,LBMPTACTIL, NCANT_TAC,NPERIGARAN,CENTALM)";
                            strSqlArt = strSqlArt + " Values ('" + dgDatos.Rows[i].Cells[1].Value.ToString() + "','" + dgDatos.Rows[i].Cells[2].Value.ToString().Replace("'", "''") + "','" + fami + "', " + 0 + "" +
                            " ," + dgDatos.Rows[i].Cells[3].Value.ToString().Replace(",", ".") + " ," + dgDatos.Rows[i].Cells[3].Value.ToString().Replace(",", ".") + "" +
                            " , " + 0 + " , " + 0 + " ,'" + cadena.Substring(0, 1).ToString() + "'," + (dgDatos.Rows[i].Cells[4].Value.ToString() == string.Empty ? "0" : dgDatos.Rows[i].Cells[4].Value.ToString().Replace(",", ".")) + "" +
                            " , " + 0 + ", " + 0 + "," + 0 + "," + 0 + ", " + 0 + " , " + 0 + " , " + 0 + " , " + 0 + ", " + 0 + ",'" + cadena.Substring(1, 3).ToString() + "'" +
                            " , " + 0 + " ,'" + cadena.Substring(4, 2) + "', " + 0 + " ," + 0 + ", " + 1 + " ," + 0 + " ," + 0 + "," + 1 + "," + 0 + ", " + 0 + "" +
                            " , " + 0 + "," + 0 + "," + 0 + " , " + 0 + " ," + 0 + " ," + 16 + " ," + 0 + "," + 1 + " ," + 0 + "," + 1 + ",'" + DateTime.Now + "'," + 0 + "" +
                            " ," + 0 + ", " + 0 + ", " + 0 + "  ," + 0 + ",'" + cadena.Substring(6, 1).ToString() + "')";
                            string alma = System.Configuration.ConfigurationSettings.AppSettings["CCODALM"];
                            strSqlStoc = "insert into stocks (CCODALM,CREF) Values ('"+ alma +"','" + dgDatos.Rows[i].Cells[1].Value.ToString() + "')";
                        
                        }
                        else if (dgDatos.Rows[i].Cells[0].Value.ToString() == "MODIFICAR")
                        {
                            strSqlArt = "update Articulo SET NCOSTEPROM = " + dgDatos.Rows[i].Cells[3].Value.ToString().Replace(",", ".") + ", " +
                                " NCOSTEDIV = " + dgDatos.Rows[i].Cells[3].Value.ToString().Replace(",", ".") + " ," +
                                " NPVP  = " + (dgDatos.Rows[i].Cells[4].Value.ToString() == string.Empty ? "0" : dgDatos.Rows[i].Cells[4].Value.ToString().Replace(",", ".")) + " " +
                                //" NPCONIVA = " + dCalculo +" " +
                                " where CREF = '" + dgDatos.Rows[i].Cells[1].Value.ToString() + "' ";
                            strSqlStoc = string.Empty;
                        }
                    } 
               if (cont < 10 && dgDatos.Rows[i].Cells[1].Value.ToString() != string.Empty && dgDatos.Rows[i].Cells[2].Value.ToString() != string.Empty && strSqlStoc != string.Empty)
               {
                    insertarMsj(i, 1);
                    cont = 0;
                    brProgreso.PerformStep();
                    totreg += 1;
                    sComando = new OdbcCommand(strSqlArt, cnConexion);
                    sComando.ExecuteNonQuery();
                    if (strSqlStoc != string.Empty)
                    {
                        sComando = new OdbcCommand(strSqlStoc, cnConexion);
                        sComando.ExecuteNonQuery();
                    }
                    else if (cont < 10)
                    {
                        cont += 1;
                        insertarMsj(i, 2);
                    }
                    else
                        break;
                    sComando = null;
               }
               }
            }
            catch (Exception ex)
            {
                throw new Exception(ex.Message);
            }
            finally
            {
                if (cnConexion != null && cnConexion.State != ConnectionState.Closed)
                {
                    cnConexion.Close();
                    cnConexion.Dispose();
                    cnConexion = null;
                }
                if (sComando != null)
                {
                    sComando.Dispose();
                    sComando = null;
                }
                MessageBox.Show("Se procesaron   " + totreg + "    registros!", "información");
                brProgreso.Value = 0;
            }
        }
        void procesar()
        {
            strRutaDbf = System.Configuration.ConfigurationSettings.AppSettings["PathDbfs"];
            strConexion = "Driver={Microsoft dBASE Driver (*.dbf)};DriverID=277;Dbq=" + strRutaDbf + ";";
            OdbcConnection bdConexion = new OdbcConnection(strConexion);
            string cadena = "GEUR86N";
            string cad = string.Empty;
            int totreg = 0;
            try
            {
                bdConexion.Open();
                Cursor.Current = Cursors.WaitCursor;
                brProgreso.Minimum = 0;
                brProgreso.Maximum = dgDatos.Rows.Count;
                brProgreso.Step = 1;
                byte cont = 0;
                
                //double dCalculo = 0; 
                for (int i = 0; i < dgDatos.Rows.Count-1; i++)
                {
                    //double dCalculo = 0;
                    //dCalculo = (Convert.ToDouble(dgDatos.Rows[i].Cells[4].Value.ToString()) * ((Convert.ToDouble(dsDocumentos.Iva.Rows[0][2].ToString()) / 100) + 1)); //*if (((System.Data.InternalDataCollectionBase)(((Exce_Grilla.dsDocumentos.Excel1DataTable)((System.Data.DataTable)(dsDocumentos.Excel1))).Columns)).Count == 28)
                    if (dgDatos.Rows[i].Cells[1].Value.ToString() == "Referencia" || dgDatos.Rows[i].Cells[2].Value.ToString() == "Descripción")
                        i++;
                    if (dgDatos.Rows[i].Cells[5].Value.ToString() == "PENDIENTE")
                    {
                        if (dgDatos.Rows[i].Cells[0].Value.ToString() == "INSERTAR")
                        {
                            cad = "insert into Articulo (CREF,CDETALLE, NPZOREP ,NCOSTEPROM,NCOSTEDIV, NSTOCKMIN,NSTOCKMAX, CTIPOIVA, NPVP,  NPREMAYOR, " +
                            " NDTO1,  NDTO2, NDTO3,  NDTO4, NDTO5, NDTO6, NPENDSER, NPENDREC , CCODDIV, NENTACU,CCODPRO, LSELECT, NUNIDADES," +
                            " NETIQUETAS,  LINTERNET, LUNIDADES, NTIPOCODE, LKIT, LPESO,NTARA, NBENEFPVP,NBENEFMAY, LPRECPROP1,  LPRECPROP2, NPCONIVA, NPMCONIVA, LCTRLSTOCK, " +
                            " NPREPV,LACTUALIZA,DACTUALIZA, NSALACU, LANTICIPO,LBMPTACTIL, NCANT_TAC,NPERIGARAN,CENTALM)";
                            cad = cad + " Values ('" + dgDatos.Rows[i].Cells[1].Value.ToString() + "','" + dgDatos.Rows[i].Cells[2].Value.ToString().Replace("'", "''") + "', " + 0 + "" +
                            " ," + dgDatos.Rows[i].Cells[3].Value.ToString().Replace(",", ".") + " ," + dgDatos.Rows[i].Cells[3].Value.ToString().Replace(",", ".") + "" +
                            " , " + 0 + " , " + 0 + " ,'" + cadena.Substring(0, 1).ToString() + "'," + (dgDatos.Rows[i].Cells[4].Value.ToString() == string.Empty ? "0" : dgDatos.Rows[i].Cells[4].Value.ToString().Replace(",", ".")) + "" +
                            " , " + 0 + ", " + 0 + "," + 0 + "," + 0 + ", " + 0 + " , " + 0 + " , " + 0 + " , " + 0 + ", " + 0 + ",'" + cadena.Substring(1, 3).ToString() + "'" +
                            " , " + 0 + " ,'" + cadena.Substring(4, 2) + "', " + 0 + " ," + 0 + ", " + 1 + " ," + 0 + " ," + 0 + "," + 1 + "," + 0 + ", " + 0 + "" +
                            " , " + 0 + "," + 0 + "," + 0 + " , " + 0 + " ," + 0 + " ," + 16 + " ," + 0 + "," + 1 + " ," + 0 + "," + 1 + ",'" + DateTime.Now + "'," + 0 + "" +
                            " ," + 0 + ", " + 0 + ", " + 0 + "  ," + 0 + ",'" + cadena.Substring(6, 1).ToString() + "')";
                        }
                        else if (dgDatos.Rows[i].Cells[0].Value.ToString() == "MODIFICAR")
                        {
                            cad = "update Articulo SET NCOSTEPROM = " + dgDatos.Rows[i].Cells[3].Value.ToString().Replace(",", ".") + ", " +
                                " NCOSTEDIV = " + dgDatos.Rows[i].Cells[3].Value.ToString().Replace(",", ".") + " ," +
                                " NPVP  = " + (dgDatos.Rows[i].Cells[4].Value.ToString() == string.Empty ? "0" : dgDatos.Rows[i].Cells[4].Value.ToString().Replace(",", ".")) + " " +
                                //" NPCONIVA = " + dCalculo +" " +
                                " where CREF = '" + dgDatos.Rows[i].Cells[1].Value.ToString() + "' ";
                        }
                    }

                    if (cont < 10 && dgDatos.Rows[i].Cells[2].Value.ToString() != string.Empty && dgDatos.Rows[i].Cells[3].Value.ToString() != string.Empty && cad != string.Empty)
                    {
                        OdbcCommand cmdInstruccion = new OdbcCommand(cad, bdConexion);
                        cad = string.Empty;
                        cmdInstruccion.ExecuteNonQuery();
                        //insEnExcel(); 
                        //i += 22;
                        insertarMsj(i, 1);
                        cont = 0;
                        brProgreso.PerformStep();
                        totreg += 1;
                    }
                    else if (cont < 10)
                    {
                        cont += 1;
                        insertarMsj(i, 2);
                    }
                    else
                        break;
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("jo " + ex.Message);
            }
            finally
            {
                Cursor.Current = Cursors.Default;
                if (oleConexion != null && oleConexion.State != ConnectionState.Closed)
                    bdConexion.Close();
                MessageBox.Show("Se procesaron   " + totreg+  "    registros!", "información");
                brProgreso.Value = 0;
            }
        }
        void insertarMsj(int fila, byte modo)
        {
            if (dgDatos.Rows[fila].Cells[5].ToString() != "PROCESADO")
                dgDatos.Rows[fila].Cells[5].Value = modo == 1 ? "PROCESADO" : "FALLIDO";

        }
        void colocarTit()
        {
            if (dgDatos.RowCount > 1)
            {
                dgDatos.Columns[1].HeaderCell.Value = "REFERENCIA";
                dgDatos.Columns[2].HeaderCell.Value = "DESCRIPCION";
                dgDatos.Columns[3].HeaderCell.Value = "PRECIO COMPRA CLIENTE";
                dgDatos.Columns[4].HeaderCell.Value = "PVP";
                dgDatos.Columns[5].HeaderCell.Value = "ESTADO";
            }
        }
        void limpiarGrilla()
        {
            dgDatos.DataSource = null;
            dsDocumentos.Excel.Clear();
            dsDocumentos.Excel1.Clear();
        }
        void abrirdbf(string tabla)
        {
            strRutaDbf = System.Configuration.ConfigurationSettings.AppSettings["PathDbfs"];
            strSelect = "SELECT * FROM "+ tabla +";" ;
            strConexion = "Driver={Microsoft dBASE Driver (*.dbf)};DriverID=277;Dbq=" + strRutaDbf+ ";";
            OdbcConnection dbConexionDbf = new System.Data.Odbc.OdbcConnection(strConexion);
            try
            {
                dbConexionDbf.Open();
                OdbcDataAdapter da = new System.Data.Odbc.OdbcDataAdapter(strSelect, dbConexionDbf);
                da.Fill(dsDocumentos,tabla);
                dbConexionDbf.Close();
            }
            catch (Exception ex)
            {
                MessageBox.Show("Error al abrir la base de datos\n" + ex.Message);
                return ;
            }
            finally
            {
                if (dbConexionDbf != null && oleConexion.State != ConnectionState.Closed)
                    dbConexionDbf.Close();
            }
        }
        void Mensaje()
        {
            Cursor.Current = Cursors.WaitCursor;
            foreach (DataRow drRow in dsDocumentos.Excel.Rows)
            {
                //if (solo si no existe la columa en excel)
                    drRow["F4"] = "PENDIENTE";
                DataRow[] accionRow = dsDocumentos.Articulo.Select("CREF = '" + drRow[1] + "'");
                //drRow["F4"] = string.Empty;
                if (accionRow.Count() > 0)
                    drRow["ACCION"] = "MODIFICAR";
                else
                    drRow["ACCION"] = "INSERTAR";
            }
            dgDatos.DataSource = dsDocumentos.Excel;
            Cursor.Current = Cursors.Default;
        }
        void insEnExcel()
        {
            string file = lRuta.Text;
            string connectionString = @"Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" + file + ";Extended Properties='Excel 8.0;HDR=Yes;IMEX=0;'";
            System.Data.OleDb.OleDbConnection oCon = new System.Data.OleDb.OleDbConnection(connectionString);
            oCon.Open();
            string q = "INSERT INTO [TaRIFA$[i]:AA[i] VALUES (Procesado)";
            int r = new System.Data.OleDb.OleDbCommand(q, oCon).ExecuteNonQuery();
            oCon.Close();
        }
        void abrirExcel()
        {
            try
            {
                OleDbDataReader drMiReader;
                dsDocumentos.Excel.Clear();
                dsDocumentos.Excel1.Clear();
                dgDatos.Rows.Clear();
                oleConexion.Open();
                strSelect = "SELECT * FROM [" + cmbHojas.Text + "$] ";
                OleDbCommand miComando = new OleDbCommand(strSelect, oleConexion);
                drMiReader = miComando.ExecuteReader(CommandBehavior.CloseConnection);
                dsDocumentos.Excel1.Load(drMiReader, LoadOption.OverwriteChanges); // verifica si la hoja es la correcta y cargar planilla *****
                if (dsDocumentos.Excel1.Rows[0].ItemArray[2].ToString() == "Referencia" && dsDocumentos.Excel1.Rows[0].ItemArray[3].ToString() == "Descripción")//&& dsDocumentos.Excel.Rows[REG].ItemArray[3].ToString() == "GP")
                {
                    strSelect = ((System.Data.InternalDataCollectionBase)(((Exce_Grilla.dsDocumentos.Excel1DataTable)((System.Data.DataTable)(dsDocumentos.Excel1))).Columns)).Count == 28 ? "" +
                 " SELECT F2,F3,F17,F23,F4 FROM [" + cmbHojas.Text + "$]" : "SELECT F2,F3,F17,F24,F28 FROM [" + cmbHojas.Text + "$]";
                    OleDbCommand miComando1 = new OleDbCommand(strSelect, oleConexion);
                    oleConexion.Open();
                    drMiReader = miComando1.ExecuteReader(CommandBehavior.CloseConnection);
                    dsDocumentos.Excel.Load(drMiReader, LoadOption.OverwriteChanges);
                    abrirdbf("Articulo");
                    //abrirdbf("Ivas");
                    dgDatos.DataSource = dsDocumentos.Excel;
                    if (dgDatos.Rows.Count != 0)
                    { colocarTit(); Mensaje(); }
                }
                else
                {
                    MessageBox.Show("La Hoja de datos seleccionada es incorrecta!", "Información");
                    limpiarGrilla();
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
                //MessageBox.Show(ex.Message);
            }
            finally
            {
                if (oleConexion != null && oleConexion.State != ConnectionState.Closed)
                    oleConexion.Close();
            }
            //   //  otra manera de ver datos
            //dgDatos.DataSource = dsDocumentos.Excel;
            //OleDbDataAdapter adapter = new OleDbDataAdapter();
            //OleDbCommand selectCommand = new OleDbCommand();
            //selectCommand.CommandText = "SELECT * FROM [" + cmbHojas.Text + "$] ";
            //selectCommand.Connection = oleConexion;
            //adapter.SelectCommand = selectCommand;
            //dsDocumentos.Excel.Clear();
            //adapter.Fill(dsDocumentos.Excel);
            //selectCommand.CommandText = strSelect;
            //oleConexion.Open();
            //selectCommand.Connection = oleConexion;
            //adapter.SelectCommand = selectCommand;
            //dsDocumentos.Excel.Clear()  ;
            //adapter.Fill(dsDocumentos.Excel);
        }

        #region no usado
        //void usarlinQ()
        //{
        //    DataTable deExcel = dsDocumentos.Excel;
        //    DataTable deDbf = dsDocumentos.Dbf;
        //    IEnumerable<DataRow> Consulta = from miHoja in deExcel.AsEnumerable()
        //                                    join miDbf in deDbf.AsEnumerable()
        //                                    on miHoja.Field<string>("F2") equals
        //                                        miDbf.Field<string>("CREF")
        //                                    select miHoja;
        //    dsDocumentos.Resultado.Columns.Add("ACCION", typeof(string));
        //    dgDatos.DataSource = Consulta.CopyToDataTable<DataRow>();
        //    Cursor.Current = Cursors.WaitCursor;
        //    foreach (DataRow drRow in dsDocumentos.Excel.Rows)
        //    {
        //        DataRow[] accionRow = dsDocumentos.Dbf.Select("CREF = '" + drRow[2] + "'");
        //        if (accionRow.Count() > 0)
        //            drRow["ACCION"] = "MODIFICA";
        //        else
        //            drRow["ACCION"] = "INSERTAR";
        //    }
        //    dgDatos.DataSource = dsDocumentos.Excel;
        //    Cursor.Current = Cursors.Default;
        //}
        void usarlinQ(string consulta)
        {
            Cursor.Current = Cursors.WaitCursor;
            //dsDocumentos.Resultado.Columns.Add("ACCION", typeof(string));
            //dgDatos.DataSource = Consulta.CopyToDataTable<DataRow>();
            //foreach (DataRow drRow in dsDocumentos.Excel.Rows)
            //{
            //    DataRow[] accionRow = dsDocumentos.Dbf.Select("CREF = '" + drRow[2] + "'");
            //    if (accionRow.Count() > 0)
            //        drRow["ACCION"] = "MODIFICA";
            //    else
            //        drRow["ACCION"] = "INSERTAR";
            //}
            //dgDatos.DataSource = dsDocumentos.Excel;
            Cursor.Current = Cursors.Default;
        }
        private bool validarDatos()
        {
            // hacer las querys que traigan la inf del stock.dbf
            //crear una funcion que me inserte en la grilla el 1er. campo accion con Insert o Update segun corresponda con el dbf
            return true;
        }
        #endregion

        private void button1_Click(object sender, EventArgs e)
        {
            borrarDbf();

        }


        #endregion

        private void button2_Click(object sender, EventArgs e)
        {
            abrirdbf("stocks");
            dgDatos.DataSource = dsDocumentos.stocks;

        }

    }
    }

        


