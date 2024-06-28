using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Text;
using System.Windows.Forms;
using System.Net;
using DevExpress.XtraEditors;
using System.Diagnostics;
using System.IO;
using f2.clases;

/////////////////////////////////////////Nancy Nuñez PNormal 4755///////////////////////////////////////////////////////////////
namespace f2
{
    public partial class Form1 : Form
    {
        #region SETEOS DE BASE
        /// <summary>
        /// estable si el form esta en produccion
        /// </summary>
        private bool esProduccion;
        /// <summary>
        /// url del webService del serviceBus
        /// </summary>
        private string urlServiceBus = "";
        /// <summary>
        /// tns del server de oracle es ip + pipe + serviceName, default dbitacua
        /// </summary>
        private string tnsORacle = "";
        /// <summary>
        /// el usuario de oracle que viene de finansys
        /// </summary>
        private string usuarioOracle = "";
        /// <summary>
        /// el password de oracle
        /// </summary>
        private string passOracle = "";
        /// <summary>
        /// los parametro recibidos desde finansys
        /// </summary>
        private string parametros = "";
        /// <summary>
        /// determina si se puede cerrar el formulario
        /// </summary>
        private bool sePuedeCerrar;
        /// <summary>
        /// el nombre de la aplicacion
        /// </summary>
        private string nombreAplicacion = "Finansys 2";
        /// <summary>
        /// seteos de formularios
        /// </summary>
        private void iniciosVarios()
        {
            //Inicializaciones de base
            //this.Icon = f2.Properties.Resources.pyg;
            //this.Text = "Finansys 2";
            //this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedSingle;
            //this.MaximizeBox = false;
            //this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
        }
        #endregion    

        #region TABLERO DE CONTROL DE DAEMONS - 08/11/2017 - 11949
        private DaemonControl.DaemonControl _controlDaemon;
        #endregion

        //private businessLayer capaNegocios = new businessLayer();
        private string v_ruta_temp = "";
        private string config_procedures = "";
        //private string v_systempath = "";
        DataSet ds_config_procedures;
        string v_cfg_destinatario = "", v_cfg_copia = "", v_logfile_name = "";
        string v_linea_log1 = "", v_linea_log2 = "";
        string smpt_user = "", smpt_pass = "", smpt_server = "", smpt_mailfrom = "";
        DateTime vt_hoy = DateTime.Now;
        private bool vt_corrio = false;
        string v_hora_ini = "", v_minuto_ini = "";
        FormPadre frm = new FormPadre();
        /// <summary>
        /// constructor del formulario
        /// </summary>
        public Form1()
        {
            InitializeComponent();
            //seteos standart de clase
            iniciosVarios();
            //establecer la version del archivo
            f2.objFinanasys2 laSesion = new objFinanasys2();
            laSesion.version = "2017110811949";// "2017102711925";
            lbVersion.Text = laSesion.version;
            this.Tag = (Object)laSesion;
        }         

        /// <summary>
        /// Ocurre antes de cerrar el formulario
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void Form1_FormClosing(object sender, FormClosingEventArgs e)
        {
            //if (sePuedeCerrar)
            //{
            //    this.DialogResult = MessageBox.Show("Cerrar reporte PBViajero?",
            //        this.nombreAplicacion,
            //        MessageBoxButtons.YesNo,
            //        MessageBoxIcon.Question,
            //        MessageBoxDefaultButton.Button1);
            //    if (this.DialogResult != DialogResult.Yes)
            //        e.Cancel = true;
            //}
            //else { e.Cancel = true; }

            if (this.sePuedeCerrar)
            {
                if (MessageBox.Show("¿Confirma que desea cerrar la ventana?", "Atención",
                MessageBoxButtons.YesNo, MessageBoxIcon.Question,
                MessageBoxDefaultButton.Button1) == DialogResult.No)
                {
                    e.Cancel = true;
                }
            }
            else
            {
                e.Cancel = true;
            }
        }
        
        private void Form1_Shown(object sender, EventArgs e)
        {
        }
        
        private void Form1_Load(object sender, EventArgs e){

            esProduccion = true;
            if (esProduccion)
            {
            //    //recupero los parametros recibidos desde finansys
            //    f2.objFinanasys2 laSesion = (f2.objFinanasys2)this.Tag;
            //    usuarioOracle = laSesion.usuario;
            //    passOracle = laSesion.userPassword;
            //    parametros = laSesion.parametros;
            //}
            //else
            //{
                ///aqui podemos colocar los datos para trabajar en un sandBox
                usuarioOracle = Properties.Settings.Default.ora_user.ToString();
                passOracle = Properties.Settings.Default.ora_pass.ToString();
                parametros = "";
                //COMENTAR DESPUES
                tnsORacle = Properties.Settings.Default.ora_tns.ToString();
            }

            frm.v_systempath = "C:\\Finansys2\\demonio_genera_judiciales";// Environment.CurrentDirectory.ToString();

         smpt_user = Properties.Settings.Default.smpt_user;
           smpt_pass = Properties.Settings.Default.smpt_pass;
            smpt_server = Properties.Settings.Default.smpt_server;
            smpt_mailfrom = Properties.Settings.Default.smpt_mailfrom;

            v_hora_ini = Properties.Settings.Default.c_hora_ini.ToString();
            v_minuto_ini = Properties.Settings.Default.c_minuto_ini.ToString();

            v_cfg_destinatario = Properties.Settings.Default.c_destinatario.ToString();
            v_cfg_copia = Properties.Settings.Default.c_copia.ToString();
            
            v_ruta_temp = Properties.Settings.Default.c_ruta_temp.ToString();

            config_procedures = frm.v_systempath + "\\" + Properties.Settings.Default.config_procedures.ToString();
            //trae los pkg a utilizar
            ds_config_procedures = Utilidades.redXml2(config_procedures);

            mpLabura.Visible = false;
            mpON.Visible = false;

            //aqui se estable la capa de negocios
            frm.capaNegocios = new businessLayer(usuarioOracle, passOracle, tnsORacle);

            #region TABLERO DE CONTROL DE DAEMONS - 08/11/2017 - 11949
            //string v_autostart = System.Configuration.ConfigurationSettings.AppSettings["autostart"].ToString();
            //if (v_autostart == "SI")
            //    arrancarDemonio();
            _controlDaemon = new DaemonControl.DaemonControl("16b24e0ad1f7990fe16545a3ba1e42ea", this.ProductName);
            arrancarDemonio();
            #endregion
        }

        //aqui generamos nuestro reporte, lo adjuntamos en excel y lo enviamos por mail
        private void button1_Click(object sender, EventArgs e)
        {
            arrancarDemonio();
           
        }

        private void arrancarDemonio()
        {

            if (!tmEjecuta1vez.Enabled)
            {
                btVerFilesConfig.Text = "APAGAR";

                mpON.Visible = true;

                tmEjecuta1vez.Enabled = true;
                tmEjecuta1vez.Start();
                tmMarcaPresencia.Enabled = true;
                tmMarcaPresencia.Start();

            }
            else
            {
                btVerFilesConfig.Text = "PRENDER";

                mpON.Visible = false;

                tmEjecuta1vez.Enabled = false;
                tmEjecuta1vez.Stop();
                tmMarcaPresencia.Enabled = false;
                tmMarcaPresencia.Stop();
            }
        }

        private void arrancarProceso()
        {
            mpLabura.Visible = true;
            this.UseWaitCursor = true;

            bgkProcesa.RunWorkerAsync();
        }

        private void correrAlertar()
        {
            //string v_nombre = "", v_comando = "";
            string v_nombre = "", v_comando = "", v_nombre_archivo = "";
            int vi = 0;
            if (ds_config_procedures.Tables.Count > 0)
            {
                if (ds_config_procedures.Tables[0].Rows.Count > 0)
                {
                    //Aca se cargan valores que le seran enviador al business layer
                    foreach (DataRow dr in ds_config_procedures.Tables[0].Rows)
                    {
                        vi++;
                        v_nombre = dr["NOMBRE"].ToString();
                        v_comando = dr["COMANDO"].ToString();
                        v_nombre_archivo = dr["NOMBRE_ARCHIVO"].ToString();
                        v_linea_log2 = "EJECUTANDO - " + dr["COMANDO"].ToString();

                        //v_destinatario = dr["MAIL_DESTINO"].ToString();
                        //v_copia = dr["MAIL_COPIA"].ToString();


                        log1(traerFecha() + "INI " + v_comando+"\r\n");
                        log1(traerFecha() + "ALERTA " + v_nombre + "\r\n");
                        
                        bgkProcesa.ReportProgress(vi);
                        generarAlerta(v_comando, v_nombre, v_nombre_archivo);

                        log1(traerFecha() + "FIN " + v_comando + "\r\n");
                    }
                }
            }
        }

        //aqui generamos nuestro reporte, lo adjuntamos en txt y lo enviamos por mail
        private void generarAlerta(string p_procedimiento, string p_titulo, string p_nombre_archivo)
        {
            DataTable dt_reporte = new DataTable("REPORTE");
            string v_html = "", v_ret_msg = "";
            string v_ruta_xls = v_ruta_temp + p_nombre_archivo;
            string v_destino = "";
            string v_copia = "";

            string v_paquete = p_procedimiento;

            frm.capaNegocios.sp_reportes_alertas(v_paquete, ref dt_reporte, ref v_html, ref v_ret_msg, ref v_destino, ref v_copia, this.frm.v_systempath + "\\" + v_logfile_name);

            //Si el stored procedure devolvio ok (es decir, trajo algo) se procede a generar un reporte en Excel
            if (v_ret_msg == "OK")
            {
                if (dt_reporte.Rows.Count > 0)
                {
                    //se genera el reporte
                    XlsExporter xls = new XlsExporter(dt_reporte, p_nombre_archivo);
                    xls.borrarFile(v_ruta_xls);
                    xls.saveFile();

                    //MiniMail mini = new MiniMail("10.1.1.160", "EnvioAutomatico", "Auto.123456");
                    MiniMail mini = new MiniMail(smpt_server, smpt_user, smpt_pass);
                    mini.fromMail = smpt_mailfrom; //email que sera usado para enviar el correo
                    mini.toMail = v_destino; //em
                    mini.ccMail = v_copia;
                    mini.IsBodyHTML = true;
                    mini.subjectMail = p_titulo;
                    mini.bodyMail = v_html;
                    mini.RutaAttach = v_ruta_xls;
                    mini.enviarDirecto();

                }
                else
                {
                    log1(traerFecha() + "La alerta no genero resultados.\r\n");
                }
            }
            else
            {
                log1(traerFecha() + "ERROR " + v_ret_msg + "\r\n");
                MiniMail mini = new MiniMail(smpt_server, smpt_user, smpt_pass);
                mini.fromMail = smpt_mailfrom;
                mini.toMail = v_cfg_destinatario;
                mini.ccMail = v_cfg_copia;
                mini.IsBodyHTML = true;
                mini.subjectMail = p_titulo;
                mini.bodyMail = "El programa arrojó una excepción:<br />" + v_ret_msg;
                mini.RutaAttach = "";
                mini.enviarDirecto();

                #region TABLERO DE CONTROL DE DAEMONS - 08/11/2017 - 11949
                _controlDaemon.AddMessage(DaemonControl.WsDaemon.TipoMensaje.Error, "El programa arrojó una excepción: " + v_ret_msg, DateTime.Now);
                #endregion
            }

        }

        private void bgkProcesa_DoWork(object sender, DoWorkEventArgs e)
        {
            correrAlertar();
        }

        private void bgkProcesa_RunWorkerCompleted(object sender, RunWorkerCompletedEventArgs e)
        {
            mpLabura.Visible = false;
            this.UseWaitCursor = false;
            teProcActual.Text = "Esperando siguiente proceso";
        }

        private void bgkProcesa_ProgressChanged(object sender, ProgressChangedEventArgs e)
        {
            teProcActual.Text = v_linea_log2;
        }


        private void log1(string p_texto)
        {
            CheckForIllegalCrossThreadCalls = false;
            teLog1.AppendText(p_texto);
        }

        public string traerFecha()
        {
            return DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss") + " ";
        }

        private void tmEjecuta1vez_Tick(object sender, EventArgs e)
        {
            
            DateTime vt_now = DateTime.Now;
            DateTime vt_hoy2 = DateTime.Now;
            string v_log = "";

            // si hoy es diferente a la fecha guardada
            //verifica si es un nuevo dia y sino corre el proceso
            //si es un nuevo dia setea vt_hoy2 
            if (vt_hoy2.ToString("yyyy-MM-dd") != vt_now.ToString("yyyy-MM-dd"))
            {
                //grabo la la nueva fecha que es hoy
                vt_hoy2 = new DateTime(vt_hoy2.Year, vt_hoy2.Month, vt_hoy2.Day, int.Parse(v_hora_ini), int.Parse(v_minuto_ini), 0);
                vt_corrio = false; //marco como que no corrio
            }
            else
            {
                //si las fechas sin iguales, verifico contra la hora
                if (vt_corrio == false && vt_now > vt_hoy) //&& vt_now.ToString("mm") == vt_hoy.ToString("mm")
                {
                    vt_corrio = true;
                    
                    //PONER AQUI LO QUE TIENE QUE CORRER UNA SOLA VEZ AL DIA
                    //correr el proceso
                    
                    arrancarProceso();
                    //teLog1.AppendText("ya corri");
                }
            }

            if (lbTimer1.Visible)
                lbTimer1.Visible = false;
            else
                lbTimer1.Visible = true;
        
        }


        private void btnSalir_Click(object sender, EventArgs e)
        {
            sePuedeCerrar = true;
            this.Close();
        }

        private void tmMarcaPresencia_Tick(object sender, EventArgs e)
        {
            this.frm.capaNegocios.controlSeguridad();
        }
    }

    #region VALIDACION CERTIFICADO SSL
    /// <summary>
    /// clase para confiar en cualquier certificado
    /// </summary>
    public class MyPolicy : ICertificatePolicy
    {
        public bool CheckValidationResult(ServicePoint srvPoint,
            System.Security.Cryptography.X509Certificates.X509Certificate certificate,
            WebRequest request, Int32 certificateProblem)
        {
            return true;
        }
    }
#endregion
}
