using System;
using System.Collections;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Data.OleDb;
using System.IO;
using Excel = Microsoft.Office.Interop.Excel;
using Outlook = Microsoft.Office.Interop.Outlook;
using DAOLibrary.DAO;
using DAOLibrary.DAO.TicketMAX;
using DAOLibrary.DTO;
using DAOLibrary.DTO.Ticket_MAX;
using DAOLibrary.DAO.CambiosMAX;
using DAOLibrary.DAO.Ticket_CA;
using DAOLibrary.DTO.Ticket_CA;
using System.Deployment.Application;

namespace SubeTickets
{
    public partial class Form1 : Form
    {
        ArrayList sedesNuevas;
        ArrayList Feriados;
        ArrayList flujo;
        ArrayList listaFiltrado;
        ArrayList listaProblemas;
        ArrayList otrostickets;
        ArrayList ClientesMAX;
        ArrayList ServiciosMAX;
        ArrayList FuenteMAX;
        ArrayList EstadoMAX;
        ArrayList ClienteCambio;
        ArrayList ServicioCambio;
        ArrayList EstadoCambio;
        ArrayList listOcts;
        ArrayList nuevosTickets;
        ArrayList oldTickets;
        ArrayList empresas;
        

        /*ArrayList para CA*/

        ArrayList EstadoCA;
        ArrayList SedeCA;
        ArrayList empresasUsuarioCA;
        ArrayList CategoriaTicketCA;
        ArrayList CategoriaConsultas;
        ArrayList MetodoCA;
        ArrayList GrupoTicketCA;
        ArrayList ReportadorCA;
        ArrayList TicketsActualizar;
        ArrayList TicketsActualizarNuevos;
        ArrayList grupoTransacciones;
        ArrayList gruposSeleccionados;
        ArrayList ticketsPruebita;


        int CantidadActualizar = 0;
        /* FIN */

        Excel.Application xlApp;
        Excel.Workbook xlWorkBook;

        public Form1()
        {
            InitializeComponent();
            sedesNuevas = new ArrayList();
            Feriados = new ArrayList();
            ClientesMAX = new ArrayList();
            ServiciosMAX = new ArrayList();
            FuenteMAX = new ArrayList();
            EstadoMAX = new ArrayList();

            ClienteCambio = new ArrayList();
            ServicioCambio = new ArrayList();
            EstadoCambio = new ArrayList();
            listaFiltrado = new ArrayList();
            listaProblemas = new ArrayList();
            empresas = new ArrayList();

            /* ARRAYLIST CA */

            //ControlTicketCA = new ArrayList();
            EstadoCA = new ArrayList();
            SedeCA = new ArrayList();
            //TipoCA = new ArrayList();
            CategoriaTicketCA = new ArrayList();
            CategoriaConsultas = new ArrayList();
            MetodoCA = new ArrayList();
            //CategoriaCA = new ArrayList();
            GrupoTicketCA = new ArrayList();
            ReportadorCA = new ArrayList();
            empresasUsuarioCA = new ArrayList();
            listOcts = new ArrayList();
            grupoTransacciones = new ArrayList();
            gruposSeleccionados = new ArrayList();
            ticketsPruebita = new ArrayList();


            /* FIN */

            Feriados.Add(new DateTime(2016, 1, 1));
            Feriados.Add(new DateTime(2016, 3, 24));
            Feriados.Add(new DateTime(2016, 3, 25));
            Feriados.Add(new DateTime(2016, 5, 1));
            Feriados.Add(new DateTime(2016, 6, 29));
            Feriados.Add(new DateTime(2016, 7, 28));
            Feriados.Add(new DateTime(2016, 7, 29));
            Feriados.Add(new DateTime(2016, 8, 30));
            Feriados.Add(new DateTime(2016, 10, 8));
            Feriados.Add(new DateTime(2016, 11, 1));
            Feriados.Add(new DateTime(2016, 12, 8));
            Feriados.Add(new DateTime(2016, 12, 25));

        }
        private void Form1_Load(object sender, EventArgs e)
        {
            rellenarTablasCA();

            if (ApplicationDeployment.IsNetworkDeployed)
            {
                this.Text = "Sube Ticket version: " + ApplicationDeployment.CurrentDeployment.CurrentVersion;
            }
        }
      
        //Métodos Generales
        #region
        private void releaseObject(object obj)
        {
            try
            {
                System.Runtime.InteropServices.Marshal.ReleaseComObject(obj);
                obj = null;
            }
            catch (Exception ex)
            {
                obj = null;
                MessageBox.Show("Unable to release the Object " + ex.ToString());
            }
            finally
            {
                GC.Collect();
            }
        }
        private Boolean esFeriado(DateTime fecha)
        {
            foreach (DateTime x in Feriados)
            {
                if (new DateTime(x.Year, x.Month, x.Day) == new DateTime(fecha.Year, fecha.Month, fecha.Day))
                {
                    return true;
                }
            }
            return false;
        }
        private decimal DiferenciaHoras(DateTime? menor, DateTime? mayor, Boolean HorasLaborales)
        {
            decimal horas = 0;

            if (menor > mayor)
            {
                return -1;
            }

            if (mayor == null || menor == null)
            {
                return 0;
            }
            TimeSpan ts = mayor.Value - menor.Value;
            if (HorasLaborales)
            {
                menor = preparaHorasLaborales(menor);
                mayor = preparaHorasLaborales(mayor);
                ts = mayor.Value - menor.Value;

                if (new DateTime(menor.Value.Year, menor.Value.Month, menor.Value.Day) == new DateTime(mayor.Value.Year, mayor.Value.Month, mayor.Value.Day))
                {
                    horas = horas + ts.Hours;
                    horas = horas + (decimal)(ts.Minutes / 60.0);
                    horas = horas + (decimal)(ts.Seconds / 3600.0);
                }
                else
                {
                    DateTime fechaaux = menor.Value;
                    DateTime findejornada = new DateTime(2022, 12, 30, 18, 0, 0);

                    while (fechaaux <= mayor)
                    {
                        if (fechaaux.DayOfWeek == DayOfWeek.Sunday || fechaaux.DayOfWeek == DayOfWeek.Saturday || esFeriado(fechaaux))
                        {
                            fechaaux = fechaaux.AddDays(1);
                        }
                        else
                        {
                            DateTime horaauxiliar = new DateTime(2022, 12, 30, fechaaux.Hour, fechaaux.Minute, fechaaux.Second);

                            if (new DateTime(fechaaux.Year, fechaaux.Month, fechaaux.Day) == new DateTime(mayor.Value.Year, mayor.Value.Month, mayor.Value.Day))
                            {
                                findejornada = new DateTime(2022, 12, 30, mayor.Value.Hour, mayor.Value.Minute, mayor.Value.Second);
                            }

                            if (horaauxiliar > findejornada)
                            {
                                fechaaux = fechaaux.AddDays(1);
                            }
                            else
                            {
                                horaauxiliar = preparaHorasLaborales(horaauxiliar);
                                ts = findejornada - horaauxiliar;
                                horas = horas + ts.Hours;
                                horas = horas + (decimal)(ts.Minutes / 60.0);
                                horas = horas + (decimal)(ts.Seconds / 3600.0);

                            }
                            fechaaux = fechaaux.AddDays(1);
                            fechaaux = new DateTime(fechaaux.Year, fechaaux.Month, fechaaux.Day, 9, 0, 0);
                        }

                    }

                }
            }
            else
            {
                horas = (ts.Days * 24);
                horas = horas + ts.Hours;
                horas = horas + (decimal)(ts.Minutes / 60.0);
                horas = horas + (decimal)(ts.Seconds / 3600.0);
            }
            return horas;
        }
        public DateTime preparaHorasLaborales(DateTime? x)
        {


            if (x.Value.Hour < 9)
            {
                x = new DateTime(x.Value.Year, x.Value.Month, x.Value.Day, 9, 0, 0);
            }
            if (x.Value.Hour >= 18)
            {
                x = new DateTime(x.Value.Year, x.Value.Month, x.Value.Day, 18, 0, 0);
            }

            return x.Value;
        }
        private string limpiarTexto(string texto)
        {
            texto = texto.Replace("´", " ");
            return texto.Replace("'", " ");
        }
        #endregion

        //Lógica Tickets CA
        #region
        private void button2_Click(object sender, EventArgs e)
        {
            string Chosen_File = "";
            openFileDialog1.Title = "Ingresa la Solicitud";
            openFileDialog1.FileName = "";
            openFileDialog1.Filter = "Archivos Excel *.xls|*.xls*";
            openFileDialog1.ShowDialog();

            Chosen_File = openFileDialog1.FileName;

            if (Chosen_File == "")
            {
                MessageBox.Show("No ha Seleccionado ningun Archivo");
            }
            else
            {
                //Sentencias Excel
                label1.Text = "Buscando Sedes nuevas..";
                object misValue = System.Reflection.Missing.Value;
                xlApp = new Excel.Application();
                xlWorkBook = xlApp.Workbooks.Open(Chosen_File, 0, true, 5, "", "", true, Microsoft.Office.Interop.Excel.XlPlatform.xlWindows, "\t", false, false, 0, true, 1, 0);
                int lastRow = -1;

                //Buscar sedes Nuevas
                foreach (Excel.Worksheet element in xlWorkBook.Worksheets)
                {

                    progressBar1.Value = 0;
                    lastRow = element.UsedRange.Rows.Count;
                    Excel.Range rango = (Excel.Range)element.get_Range("B6", "BB" + lastRow);
                    progressBar1.Maximum = rango.Rows.Count;

                    for (int row = 1; row <= rango.Rows.Count; row++)
                    {
                        if ((rango.Cells[row, 26] as Excel.Range).Value2 != null)
                        {
                            if (encontrarCodigoSede((rango.Cells[row, 26] as Excel.Range).Value2.ToString()) == 0 && buscarSedeNuevas((rango.Cells[row, 26] as Excel.Range).Value2.ToString()) < 0)
                            {
                                sedesNuevas.Add((rango.Cells[row, 26] as Excel.Range).Value2.ToString());
                            }
                        }
                        progressBar1.Value += 1;
                    }
                }

                if (sedesNuevas.Count > 0)
                {
                    DAOEmpresa daoempresa = new DAOEmpresa();
                    DataSet dsempresa = daoempresa.selectEmpresas();
                    cboempresaSE.DataSource = dsempresa.Tables[0];
                    cboempresaSE.ValueMember = "codEmpresa";
                    cboempresaSE.DisplayMember = "nomEmpresa";

                    TicketsActualizar = new ArrayList();
                    TicketsActualizarNuevos = new ArrayList();

                    this.Size = new System.Drawing.Size(610, 367);
                    groupBox1.Visible = true;
                    txtnomsede.Text = sedesNuevas[0].ToString();

                }
                else
                {
                    groupBox1.Visible = false;
                    this.Size = new System.Drawing.Size(610, 148);

                    lastRow = lastRow - 6;

                    progressBar1.Value = 0;
                    label1.Text = "Leyendo las filas  1/" + lastRow;
                    progressBar1.Maximum = lastRow;

                    TicketsActualizar = new ArrayList();
                    TicketsActualizarNuevos = new ArrayList();

                    backgroundWorker4.RunWorkerAsync();
                }
            }
        }
        
        private void backgroundWorker4_DoWork(object sender, DoWorkEventArgs e)
        {
            DateTime start = DateTime.Now;
            e.Result = "";

            object misValue = System.Reflection.Missing.Value;

            foreach (Excel.Worksheet element in xlWorkBook.Worksheets)
            {

                progressBar1.Value = 0;
                int lastRow = -1;
                lastRow = element.get_Range("B" + element.Rows.Count).get_End(Excel.XlDirection.xlUp).Row;


                Excel.Range rango = (Excel.Range)element.get_Range("B6", "BB" + lastRow);
                CantidadActualizar = rango.Rows.Count;

                if (lastRow < 5)
                {
                    //Recorriendo Rango de Datos
                    for (int row = 1; row <= rango.Rows.Count; row++)
                    {
                        DTOTicket_CA ticket = new DTOTicket_CA();

                        string tipo = (rango.Cells[row, 4] as Excel.Range).Value2.ToString();
                        ticket.CodTicket = (int)Int32.Parse((rango.Cells[row, 1] as Excel.Range).Value2.ToString());

                        if (tipo == "Incidente")
                        {
                            ticket.TipoCA_ticket = "I";
                            ticket.CodTipo = 3;
                        }
                        if (tipo == "OC con Tareas")
                        {
                            ticket.TipoCA_ticket = "OC_T";
                            ticket.CodTipo = 4;
                        }
                        if (tipo == "OC sin Tareas")
                        {
                            ticket.TipoCA_ticket = "OC";
                            ticket.CodTipo = 1;
                        }
                        if (tipo == "Problema")
                        {
                            ticket.TipoCA_ticket = "PR";
                            ticket.CodTipo = 5;
                        }
                        if (tipo == "Solicitud")
                        {
                            ticket.TipoCA_ticket = "S";
                            ticket.CodTipo = 1;
                        }

                        //En caso fuera una OC_T
                        if (ticket.TipoCA_ticket == "OC_T")
                        {
                            int valor = buscarEnLista((int)Int32.Parse((rango.Cells[row, 1] as Excel.Range).Value2.ToString()), listOcts);

                            //Existe ya la OC_T, solo se inserta la tarea
                            if (valor != -1)
                            {
                                if ((rango.Cells[row, 47] as Excel.Range).Value2.ToString() != "Tarea de inicio de grupo" && (rango.Cells[row, 47] as Excel.Range).Value2.ToString() != "Tarea de finalización de grupo")
                                {
                                    ticket = guardarTarea(ticket, rango, row, TicketsActualizarNuevos, TicketsActualizar);
                                }
                            }
                            //No existe la OC_T, hay que crearla primero y luego la tarea.
                            else
                            {
                                //Insertar OC_T
                                DTOTicket_CA ordendecambio = new DTOTicket_CA();

                                ordendecambio.CodTicket = (int)Int32.Parse((rango.Cells[row, 1] as Excel.Range).Value2.ToString());
                                ordendecambio.TipoCA_ticket = "OC_T";
                                ordendecambio.CodTipo = 4;
                                ordendecambio = guardarTicket(ordendecambio, rango, row, TicketsActualizarNuevos, TicketsActualizar);

                                if (ordendecambio != null)
                                {
                                    listOcts.Add(ordendecambio);
                                }

                                //Insertar Tarea
                                string a = (rango.Cells[row, 47] as Excel.Range).Value2.ToString();
                                if ((a != "Tarea de inicio de grupo") && (a != "Tarea de finalización de grupo"))
                                {
                                    ticket = guardarTarea(ticket, rango, row, TicketsActualizarNuevos, TicketsActualizar);
                                }
                            }
                        }
                        else
                        //Demas Tipos de Ticket
                        {
                            ticket = guardarTicket(ticket, rango, row, TicketsActualizarNuevos, TicketsActualizar);
                        }

                        backgroundWorker4.ReportProgress(row - 1, DateTime.Now);

                        if (backgroundWorker4.CancellationPending)
                        {
                            e.Cancel = true;
                            return;
                        }
                    }
                }

               
                releaseObject(element);

            }//fin de recorrido de hojas


            xlWorkBook.Close(false, misValue, misValue);
            xlApp.Quit();

            releaseObject(xlWorkBook);
            releaseObject(xlApp);

            TimeSpan duration = DateTime.Now - start;
            e.Result = "Duracion: " + duration.TotalMinutes.ToString() + "m.";
        }
        private void backgroundWorker4_ProgressChanged(object sender, ProgressChangedEventArgs e)
        {
            progressBar1.Value = e.ProgressPercentage;
            label1.Text = "Leyendo las filas  " + e.ProgressPercentage + "/" + CantidadActualizar;
            DateTime time = Convert.ToDateTime(e.UserState);
        }
        private void backgroundWorker4_RunWorkerCompleted(object sender, RunWorkerCompletedEventArgs e)
        {
            if (e.Cancelled)
            {
                MessageBox.Show("The task has been cancelled");
            }
            else if (e.Error != null)
            {
                MessageBox.Show("Error. Details: " + (e.Error as Exception).ToString());
            }
            else
            {
                //TicketsActualizar.Clear();
                MessageBox.Show("Se Leyeron : " + CantidadActualizar + " Tickets en: " + e.Result);
                label1.Text = "Actualizando Tickets... Total:" + TicketsActualizar.Count;
                progressBar1.Value = 0;
                progressBar1.Maximum = TicketsActualizar.Count;
                CantidadActualizar = TicketsActualizar.Count;
                backgroundWorker1.RunWorkerAsync();
            }
        }

        private void backgroundWorker1_DoWork(object sender, DoWorkEventArgs e)
        {
            DateTime start = DateTime.Now;
            e.Result = "";
            DAOTicket_CA daoTicket = new DAOTicket_CA();
            int row = CantidadActualizar;
            foreach (DTOTicket_CA upd in TicketsActualizar)
            {
                daoTicket.UpdateTicket(upd);

                backgroundWorker1.ReportProgress(row - 1, DateTime.Now);

                if (backgroundWorker1.CancellationPending)
                {
                    e.Cancel = true;
                    return;
                }
                row--;
            }

            TimeSpan duration = DateTime.Now - start;
            e.Result = "Duracion: " + duration.TotalMinutes.ToString() + "m.";

        }
        private void backgroundWorker1_ProgressChanged(object sender, ProgressChangedEventArgs e)
        {
            progressBar1.Value = e.ProgressPercentage;
            label1.Text = "Actualizando Tickets... " + e.ProgressPercentage + "/" + CantidadActualizar;
            DateTime time = Convert.ToDateTime(e.UserState);

        }
        private void backgroundWorker1_RunWorkerCompleted(object sender, RunWorkerCompletedEventArgs e)
        {
            if (e.Cancelled)
            {
                MessageBox.Show("The task has been cancelled");
            }
            else if (e.Error != null)
            {
                MessageBox.Show("Error. Details: " + (e.Error as Exception).ToString());
            }
            else
            {
                //TicketsActualizar.Clear();
                MessageBox.Show("Se Actualizaron : " + TicketsActualizar.Count + " Tickets en: " + e.Result);
                label1.Text = "Insertando tickets Nuevos.. Total: " + TicketsActualizarNuevos.Count;
                progressBar1.Value = 0;
                progressBar1.Maximum = TicketsActualizarNuevos.Count;
                CantidadActualizar = TicketsActualizarNuevos.Count;
                backgroundWorker5.RunWorkerAsync();
            }
        }


        private void backgroundWorker5_DoWork(object sender, DoWorkEventArgs e)
        {
            DateTime start = DateTime.Now;
            e.Result = "";
            int row = CantidadActualizar;
            DAOTicket_CA dao = new DAOTicket_CA();

            foreach (DTOTicket_CA ticket in TicketsActualizarNuevos)
            {
                if (ticket.TipoCA_ticket == "T" && esMDA(ticket.CodGrupo_ticket) == 1)
                {
                    ticket.ID_control = 0;
                    ticket.Motivo_mal_ticket = "NINGUNO";
                }

                if (ticket.TipoCA_ticket == "OC_T")
                {
                    Boolean ok = dao.EliminarOCxOCT(ticket);
                }

                dao.InsertTicket(ticket);

                backgroundWorker5.ReportProgress(row - 1, DateTime.Now);
                if (backgroundWorker5.CancellationPending)
                {
                    e.Cancel = true;
                    return;
                }
                row--;
            }

            TimeSpan duration = DateTime.Now - start;
            e.Result = "Duracion: " + duration.TotalMinutes.ToString() + "m.";
        }
        private void backgroundWorker5_ProgressChanged(object sender, ProgressChangedEventArgs e)
        {
            progressBar1.Value = e.ProgressPercentage;
            label1.Text = "Insertando tickets ... " + e.ProgressPercentage + "/" + CantidadActualizar;
            DateTime time = Convert.ToDateTime(e.UserState);
        }
        private void backgroundWorker5_RunWorkerCompleted(object sender, RunWorkerCompletedEventArgs e)
        {
            if (e.Cancelled)
            {
                MessageBox.Show("The task has been cancelled");
            }
            else if (e.Error != null)
            {
                MessageBox.Show("Error. Details: " + (e.Error as Exception).ToString());
            }
            else
            {
                TicketsActualizar.Clear();
                MessageBox.Show("Se insertaron :" + TicketsActualizarNuevos.Count + " Filas" + e.Result);

                progressBar1.Maximum = listOcts.Count;
                label1.Text = "Actualizando Tareas ";

                progressBar1.Value = 0;
                foreach (DTOTicket_CA ltOct in listOcts)
                {
                    update_OCT(ltOct);
                    progressBar1.Value += 1;
                }
                progressBar1.Value = 0;
                MessageBox.Show("Tareas actualizadas");

                //Ultima Subida...

                label1.Text = "Insertando tickets finales";
                progressBar1.Value = 0;
                progressBar1.Maximum = listaFiltrado.Count;

                int resultado = cargarDatos(listaFiltrado, progressBar1);

                MessageBox.Show("Se insertaron :" + resultado + " Filas, SE ACABOOOOO DANIEL- SAN ");


            }
        }


        //Métodos Auxiliares - Tickets CA
        private int buscarSedeNuevas(string nombre)
        {
            foreach (string sede in sedesNuevas)
            {
                if (sede == nombre)
                {
                    return 1;
                }
            }
            return -1;
        }
        public int encontrarCodigoSede(string nombre)
        {
            int codigo = 0;
            DTOSede sede = new DTOSede();
            sede.NomSede = nombre;
            DAOSede daosede = new DAOSede();
            DataSet dtr = new DataSet();

            dtr = daosede.selectSede(sede);
            if (dtr.Tables[0] != null && dtr.Tables[0].Rows.Count > 0)
            {
                codigo = (int)Int32.Parse(dtr.Tables[0].Rows[0]["codSede"].ToString());
            }
            return codigo;
        }
        public int buscarEnLista(int codigo, ArrayList lista)
        {
            int posicion = -1;

            foreach (DTOTicket_CA ticket in lista)
            {

                if (ticket.CodTicket == codigo)
                {
                    posicion = codigo;
                    return posicion;
                }
            }


            return posicion;
        }
        public Boolean definirsiesConsulta(string nombre)
        {

            foreach (DTOCategoriaCA_Ticket categoria in CategoriaConsultas)
            {
                if (categoria.NomCategoriaCA_ticket == nombre)
                {
                    return true;
                }
            }

            return false;
        }
        public int esMDA(int? codigo)
        {
            int valor = 0;
            DAOGrupo_Ticket daogrupo_ticket = new DAOGrupo_Ticket();
            DataTable dt = daogrupo_ticket.selectGrupo_TicketTipMDA();

            foreach (DataRow rw in dt.Rows)
            {
                int codigoGrupo = (int)Int32.Parse(rw[0].ToString());

                if (codigo == codigoGrupo)
                {
                    return 1;
                }
            }

            return valor;

        }
        public int cargarDatos(ArrayList lista, ProgressBar p)
        {
            int q = 0;
            //int error = 0;

            DAOTicket_CA dao = new DAOTicket_CA();


            foreach (DTOTicket_CA ticket in lista)
            {
                if (ticket.TipoCA_ticket == "T" && esMDA(ticket.CodGrupo_ticket) == 1)
                {
                    ticket.ID_control = 0;
                    ticket.Motivo_mal_ticket = "NINGUNO";
                }

                if (dao.InsertTicket(ticket) > 0)
                {
                    q++;
                }
                p.Value += 1;
            }

            return q;
        }
        private Boolean update_OCT(DTOTicket_CA ticket_OCT)
        {
            DAOTicket_CA doct = new DAOTicket_CA();
            DataTable dtt = new DataTable();

            dtt = doct.MaxTareaOCT_CA(ticket_OCT);

            int grupo_tarea = (int)Int32.Parse(dtt.Rows[0]["codGrupo_ticket"].ToString());

            if (grupo_tarea != ' ' || grupo_tarea != null)
            {
                doct.Update_OCT(ticket_OCT, grupo_tarea);
            }
            return true;
        }
        private int hayqueactualizarCA(DTOTicket_CA ticket)
        {
            DAOTicket_CA dao = new DAOTicket_CA();
            DateTime fecUltMod = new DateTime(2011, 01, 01);
            DataTable dt = dao.MostrarDatos(ticket);
            if (dt.Rows.Count <= 0)
            {
                return -1;
            }

            if (dt.Rows[0]["ultmod"].ToString() != "")
            {
                fecUltMod = DateTime.ParseExact(dt.Rows[0]["ultmod"].ToString(), "yyyy-MM-dd HH:mm:ss", null);
            }


            if (ticket.Fec_ultmod_ticket > fecUltMod)
            {
                return 1;
            }
            else
            {
                return 0;
            }


        }
        private void rellenarTablasCA()
        {
            DAOCategoriaCA_Ticket daoCategoriaTicket = new DAOCategoriaCA_Ticket();

            //ESTADO
            DAOEstado daoEstado = new DAOEstado();
            DataTable dtEstado = daoEstado.MostrarDatosEstado();
            foreach (DataRow rw in dtEstado.Rows)
            {
                DTOEstado estado = new DTOEstado();
                estado.CodEstado = (int)Int32.Parse(rw["codEstado"].ToString());
                estado.NomEstado = rw["nomEstado"].ToString();
                EstadoCA.Add(estado);
            }

            //GRUPO_TICKET
            DAOGrupo_Ticket daoGrupoTicket = new DAOGrupo_Ticket();
            DataTable dtGrupoTicket = daoGrupoTicket.MostrarDatosGrupo_Ticket();
            foreach (DataRow rw in dtGrupoTicket.Rows)
            {
                DTOGrupo_Ticket grupoTicket = new DTOGrupo_Ticket();
                grupoTicket.CodGrupo_ticket = (int)Int32.Parse(rw["codGrupo_ticket"].ToString());
                grupoTicket.AbreGrupo_ticket = rw["abreGrupo_ticket"].ToString();
                grupoTicket.TipGrupo_ticket = rw["tipGrupo_ticket"].ToString();
                GrupoTicketCA.Add(grupoTicket);
            }

            //METODO
            DAOMetodo daoMetodo = new DAOMetodo();
            DataTable dtMetodo = daoMetodo.MostrarDatosMetodos();
            foreach (DataRow rw in dtMetodo.Rows)
            {
                DTOMetodo metodo = new DTOMetodo();
                metodo.CodMetodo = (int)Int32.Parse(rw["codMetodo"].ToString());
                metodo.NomMetodo = rw["nomMetodo"].ToString();
                MetodoCA.Add(metodo);
            }

            //SEDE
            DAOSede daosede = new DAOSede();
            DataTable dtsede = daosede.MostrarDatosSede();
            foreach (DataRow rw in dtsede.Rows)
            {
                DTOSede sede = new DTOSede();
                sede.CodSede = (int)UInt32.Parse(rw["codSede"].ToString());
                sede.NomSede = rw["nomSede"].ToString();
                SedeCA.Add(sede);
            }

            //CATEGORIA
            DAOCategoriaCA_Ticket daocategoria = new DAOCategoriaCA_Ticket();
            DataTable dtcategoria = daocategoria.MostrarDatosCategoria_Ticket();
            foreach (DataRow rw in dtcategoria.Rows)
            {
                DTOCategoriaCA_Ticket categoria = new DTOCategoriaCA_Ticket();
                categoria.CodCategoriaCA_ticket = (int)UInt32.Parse(rw["codCategoriaCA_ticket"].ToString());
                categoria.NomCategoriaCA_ticket = rw["nomCategoriaCA_ticket"].ToString();
                categoria.AuxCategoriaCA_ticket = rw["auxCategoriaCA_ticket"].ToString();
                CategoriaTicketCA.Add(categoria);

                if (categoria.AuxCategoriaCA_ticket.Equals("Consulta") || categoria.AuxCategoriaCA_ticket.Equals("Incidente Masivo"))
                {
                    CategoriaConsultas.Add(categoria);
                }
            }

            //EMPRESA_USUARIO
            DAOEmpresa_Usuario daoempresausuario = new DAOEmpresa_Usuario();
            DataTable dtempresausuario = daoempresausuario.selectEmpresasUsuario().Tables[0];

            foreach (DataRow rw in dtempresausuario.Rows)
            {
                DTOEmpresa_Usuario empresaUsuario = new DTOEmpresa_Usuario();
                empresaUsuario.Id_emp_usu = (int)UInt32.Parse(rw["id_emp_usu"].ToString());
                empresaUsuario.Nom_emp_usu = rw["nom_emp_usu"].ToString();
                empresasUsuarioCA.Add(empresaUsuario);
            }

            //REPORTADOR
            DAOReportador daoreport = new DAOReportador();
            DataTable dtreport = daoreport.MostrarDatosReport();
            foreach (DataRow rw in dtreport.Rows)
            {
                DTOReportador reportador = new DTOReportador();
                reportador.Id_reportador = (int)Int32.Parse(rw["id_reportador"].ToString());
                reportador.NomReportador = rw["nomReportador"].ToString();
                ReportadorCA.Add(reportador);
            }

            //LISTA DE TRANSACCIONES
            DAOGrupoTrans daotransacciones = new DAOGrupoTrans();
            DataTable dttransacciones = daotransacciones.ListarLista();
            foreach (DataRow rw in dttransacciones.Rows)
            {
                DTOGrupoTrans transaccion = new DTOGrupoTrans();
                transaccion.CodGrupo_ticket = (int)Int32.Parse(rw["codgrupo_ticket"].ToString());
                transaccion.GrupoTrans = rw["grupoTrans"].ToString();
                grupoTransacciones.Add(transaccion);
            }

            DataTable dttransaccionesSeleccionadas = daotransacciones.ListarTransacciones();
            foreach (DataRow rw in dttransacciones.Rows)
            {
                int codGrupo_ticket = (int)Int32.Parse(rw["codgrupo_ticket"].ToString());
                gruposSeleccionados.Add(codGrupo_ticket);
            }
        }
        private int? buscarTablasCA(string tipo, string nombre)
        {
            int? m = -1;

            nombre = nombre.Trim().ToUpper();

            switch (tipo)
            {
                // ESTADO
                case "E":
                    foreach (DTOEstado estado in EstadoCA)
                    {
                        if (estado.NomEstado.Trim().ToUpper().Equals(nombre))
                        {
                            return estado.CodEstado;
                        }
                    }
                    if (m == -1)
                    {
                        DAOEstado daoestado = new DAOEstado();
                        DTOEstado nuevoEstado = new DTOEstado();
                        nuevoEstado.CodEstado = (int)Int32.Parse(daoestado.selectEstadoMayor().Tables[0].Rows[0][0].ToString()) + 1;
                        nuevoEstado.NomEstado = nombre;
                        daoestado.InsertEstado(nuevoEstado);
                        EstadoCA.Add(nuevoEstado);
                        m = nuevoEstado.CodEstado;
                    }
                    break;
                case "G":
                    foreach (DTOGrupo_Ticket grupoTicket in GrupoTicketCA)
                    {
                        if (grupoTicket.AbreGrupo_ticket.Trim().ToUpper().Equals(nombre))
                        {
                            return grupoTicket.CodGrupo_ticket;
                        }
                    }
                    if (m == -1)
                    {
                        DAOGrupo_Ticket daogrupoTicket = new DAOGrupo_Ticket();
                        DTOGrupo_Ticket nuevoGrupoTicket = new DTOGrupo_Ticket();
                        nuevoGrupoTicket.CodGrupo_ticket = (int)Int32.Parse(daogrupoTicket.selectGrupo_TicketMayor().Tables[0].Rows[0][0].ToString()) + 1;
                        nuevoGrupoTicket.AbreGrupo_ticket = nombre;
                        daogrupoTicket.InsertGrupo_Ticket(nuevoGrupoTicket);
                        GrupoTicketCA.Add(nuevoGrupoTicket);
                        m = nuevoGrupoTicket.CodGrupo_ticket;
                    }
                    break;
                case "M":
                    foreach (DTOMetodo metodo in MetodoCA)
                    {
                        if (metodo.NomMetodo.Trim().ToUpper().Equals(nombre))
                        {
                            return metodo.CodMetodo;
                        }
                    }
                    if (m == -1)
                    {
                        DAOMetodo daometod = new DAOMetodo();
                        DTOMetodo nuevometodo = new DTOMetodo();
                        nuevometodo.CodMetodo = (int)Int32.Parse(daometod.selectMetodoMayor().Tables[0].Rows[0][0].ToString()) + 1;
                        nuevometodo.NomMetodo = nombre;
                        daometod.InsertMetodo(nuevometodo);
                        MetodoCA.Add(nuevometodo);
                        m = nuevometodo.CodMetodo;
                    }
                    break;
                case "S":
                    foreach (DTOSede sede in SedeCA)
                    {
                        if (sede.NomSede.Trim().ToUpper().Equals(nombre))
                        {
                            return sede.CodSede;
                        }
                    }
                    if (m == -1)
                    {
                        DAOSede daosede = new DAOSede();
                        DTOSede nuevosede = new DTOSede();
                        nuevosede.CodSede = (int)Int32.Parse(daosede.selectSedeMayor().Tables[0].Rows[0][0].ToString()) + 1;
                        nuevosede.NomSede = nombre;
                        nuevosede.CodEmpresa = 27;
                        daosede.InsertSede(nuevosede);
                        SedeCA.Add(nuevosede);
                        m = nuevosede.CodSede;
                    }
                    break;
                case "R":
                    foreach (DTOReportador reportador in ReportadorCA)
                    {
                        if (reportador.NomReportador.Trim().ToUpper().Equals(nombre))
                        {
                            return reportador.Id_reportador;
                        }
                    }
                    if (m == -1)
                    {
                        DAOReportador daoreportador = new DAOReportador();
                        DTOReportador nuevoreportador = new DTOReportador();
                        nuevoreportador.Id_reportador = (int)Int32.Parse(daoreportador.selectReportadorMayor().Tables[0].Rows[0][0].ToString()) + 1;
                        nuevoreportador.NomReportador = nombre;
                        daoreportador.InsertReportador(nuevoreportador);
                        ReportadorCA.Add(nuevoreportador);
                        m = nuevoreportador.Id_reportador;
                    }
                    break;
                case "C":
                    foreach (DTOCategoriaCA_Ticket categoria in CategoriaTicketCA)
                    {
                        if (categoria.NomCategoriaCA_ticket.Trim().ToUpper().Equals(nombre))
                        {
                            return categoria.CodCategoriaCA_ticket;
                        }
                    }
                    if (m == -1)
                    {
                        DAOCategoriaCA_Ticket daocategoria = new DAOCategoriaCA_Ticket();
                        DTOCategoriaCA_Ticket nuevocategoria = new DTOCategoriaCA_Ticket();
                        nuevocategoria.CodCategoriaCA_ticket = (int)Int32.Parse(daocategoria.selectCategoriaCA_TicketMayor().Tables[0].Rows[0][0].ToString()) + 1;
                        nuevocategoria.NomCategoriaCA_ticket = nombre;
                        nuevocategoria.AuxCategoriaCA_ticket = "Atencion";
                        daocategoria.InsertCategoriaCA_Ticket(nuevocategoria);
                        CategoriaTicketCA.Add(nuevocategoria);
                        m = nuevocategoria.CodCategoriaCA_ticket;
                    }
                    break;
                case "U":
                    foreach (DTOEmpresa_Usuario empresaUsuario in empresasUsuarioCA)
                    {
                        if (empresaUsuario.Nom_emp_usu.Trim().ToUpper().Equals(nombre))
                        {
                            return empresaUsuario.Id_emp_usu;
                        }
                    }
                    if (m == -1)
                    {
                        DAOEmpresa_Usuario dao = new DAOEmpresa_Usuario();
                        DTOEmpresa_Usuario nuevo = new DTOEmpresa_Usuario();
                        nuevo.Id_emp_usu = (int)Int32.Parse(dao.selectEmpresaUsuarioMayor().Tables[0].Rows[0][0].ToString()) + 1;
                        nuevo.Nom_emp_usu = nombre;
                        dao.InsertEmpresaUsuario(nuevo);
                        empresasUsuarioCA.Add(nuevo);
                        m = nuevo.Id_emp_usu;
                    }
                    break;

            }
            return m;
        }
        private Boolean celdaValidar(Excel.Range rango, int row, int col)
        {
            if ((rango.Cells[row, col] as Excel.Range).Value2 != null)
            {
                if (!(rango.Cells[row, col] as Excel.Range).Value2.ToString().Trim().Equals(""))
                {
                    return true;
                }
                else
                {
                    return false;
                }
            }
            else
            {
                return false;
            }
        }
        public string obtenerGrupoTransaccion(int codcategoria, string transaccion)
        {

            foreach (DTOGrupoTrans g in grupoTransacciones)
            {
                if (g.CodGrupo_ticket == codcategoria)
                {
                    if (transaccion.ToUpper().Contains(g.GrupoTrans.ToUpper()))
                    {
                        return g.GrupoTrans;
                    }
                }
            }
            return null;
        }
        public Boolean esCategoriaSeleccionada(int categoria)
        {

            foreach (int t in gruposSeleccionados)
            {
                if (t == categoria)
                {
                    return true;
                }
            }
            return false;
        }

        private DTOTicket_CA guardarTarea(DTOTicket_CA t, Excel.Range rango, int row, ArrayList listaNuevas, ArrayList listaUpdate)
        {
            DTOTicket_CA ticket = new DTOTicket_CA();
            ticket.CodTicket = t.CodTicket;
            ticket.Secuencia_tarea = (int)Int32.Parse((rango.Cells[row, 46] as Excel.Range).Value2.ToString());
            ticket.CodTipo = 1;
            ticket.TipoCA_ticket = "T";
            ticket.Prioridad_ticket = (rango.Cells[row, 6] as Excel.Range).Value2.ToString();

            if (celdaValidar(rango, row, 19))
            {
                ticket.Fec_ultmod_ticket = DateTime.FromOADate(Double.Parse((rango.Cells[row, 19] as Excel.Range).Value2.ToString()));
            }

            int u1 = hayqueactualizarCA(ticket);

            if (u1 == 0)
            {
                return null;
            }

            if (celdaValidar(rango, row, 9))
            {
                ticket.Fech_ape_ticket = DateTime.FromOADate(Double.Parse((rango.Cells[row, 9] as Excel.Range).Value2.ToString()));
            }

            if (celdaValidar(rango, row, 20))
            {
                ticket.Solicitante_ticket = (rango.Cells[row, 20] as Excel.Range).Value2.ToString();
            }

            if (celdaValidar(rango, row, 23))
            {
                ticket.Id_emp_usu = buscarTablasCA("U", (rango.Cells[row, 23] as Excel.Range).Value2.ToString());
            }
            else
            {
                ticket.Id_emp_usu = 103;
            }

            if (celdaValidar(rango, row, 25))
            {
                ticket.CodSedeUsuario = buscarTablasCA("S", (rango.Cells[row, 25] as Excel.Range).Value2.ToString());
            }
            else
            {
                ticket.CodSedeUsuario = 287;
            }

            if (celdaValidar(rango, row, 26))
            {
                ticket.CodSede = buscarTablasCA("S", (rango.Cells[row, 26] as Excel.Range).Value2.ToString());
            }
            else
            {
                ticket.CodSede = 287;
            }

            if (celdaValidar(rango, row, 28))
            {
                ticket.Repor_por_ticket = buscarTablasCA("R", (rango.Cells[row, 28] as Excel.Range).Value2.ToString());
            }
            else
            {
                ticket.Repor_por_ticket = null;
            }

            if (celdaValidar(rango, row, 29))
            {
                ticket.CodCategoriaCA_ticket = buscarTablasCA("C", (rango.Cells[row, 29] as Excel.Range).Value2.ToString());
                if (definirsiesConsulta((rango.Cells[row, 29] as Excel.Range).Value2.ToString()))
                {
                    ticket.CodTipo = 6; //Es una consulta
                }
            }
            else
            {
                ticket.CodCategoriaCA_ticket = 120;
            }



            if (celdaValidar(rango, row, 31))
            {
                ticket.CodGrupo_propi_ticket = buscarTablasCA("G", (rango.Cells[row, 31] as Excel.Range).Value2.ToString());
            }
            else
            {
                ticket.CodGrupo_propi_ticket = 14;
            }

            if (celdaValidar(rango, row, 32))
            {
                ticket.CodGrupo_ticket = buscarTablasCA("G", (rango.Cells[row, 32] as Excel.Range).Value2.ToString());
            }
            else
            {
                ticket.CodGrupo_ticket = 14;
            }

            ticket.Pendiente_por_ticket = "";

            if (celdaValidar(rango, row, 38))
            {
                ticket.CodMetodo = buscarTablasCA("M", (rango.Cells[row, 38] as Excel.Range).Value2.ToString());
            }
            else
            {
                ticket.CodMetodo = 1;
            }

            if (celdaValidar(rango, row, 40))
            {
                ticket.Met_resol_ticket = (rango.Cells[row, 40] as Excel.Range).Value2.ToString();
            }
            else
            {
                ticket.Met_resol_ticket = null;
            }

            if (celdaValidar(rango, row, 41))
            {
                ticket.Socie_sap = (rango.Cells[row, 41] as Excel.Range).Value2.ToString();
            }
            else
            {
                ticket.Socie_sap = null;
            }

            if (celdaValidar(rango, row, 42))
            {
                ticket.Comp_sap = (rango.Cells[row, 42] as Excel.Range).Value2.ToString();
            }
            else
            {
                ticket.Comp_sap = null;
            }

            if (celdaValidar(rango, row, 43))
            {
                ticket.Transaccion = (rango.Cells[row, 43] as Excel.Range).Value2.ToString();
            }
            else
            {
                ticket.Transaccion = null;
            }

            //Validar Transacción
            if (esCategoriaSeleccionada(ticket.CodGrupo_ticket.Value))
            {
                if (ticket.Transaccion != null)
                {
                    ticket.GrupoTrans = obtenerGrupoTransaccion(ticket.CodGrupo_ticket.Value, ticket.Transaccion);
                }
                else
                {
                    ticket.GrupoTrans = null;
                }
            }
            else
            {
                ticket.GrupoTrans = null;
            }


            ticket.Cancelado_por_ticket = "";
            ticket.Solucion_ticket = "";

            if (celdaValidar(rango, row, 47))
            {
                ticket.Resu_ticket = (rango.Cells[row, 47] as Excel.Range).Value2.ToString();
                ticket.Des_ticket = (rango.Cells[row, 47] as Excel.Range).Value2.ToString();
            }
            else
            {
                ticket.Resu_ticket = "";
                ticket.Des_ticket = "";
            }

            if (celdaValidar(rango, row, 48))
            {
                ticket.CodEstado = buscarTablasCA("E", (rango.Cells[row, 48] as Excel.Range).Value2.ToString());
            }
            else
            {
                ticket.CodEstado = null;
            }

            ticket.Causa_ticket = "";

            if (celdaValidar(rango, row, 50))
            {
                ticket.NomEspecialista = (rango.Cells[row, 50] as Excel.Range).Value2.ToString();
            }
            else
            {
                ticket.NomEspecialista = "";
            }

            if (celdaValidar(rango, row, 52))
            {
                ticket.Fec_tran_ticket = DateTime.FromOADate(Double.Parse((rango.Cells[row, 52] as Excel.Range).Value2.ToString()));
            }
            else
            {
                ticket.Fec_tran_ticket = null;
            }

            if (celdaValidar(rango, row, 53))
            {
                ticket.Fec_res_ticket = DateTime.FromOADate(Double.Parse((rango.Cells[row, 53] as Excel.Range).Value2.ToString()));
                ticket.Fec_cie_ticket = DateTime.FromOADate(Double.Parse((rango.Cells[row, 53] as Excel.Range).Value2.ToString()));
            }
            else
            {
                ticket.Fec_res_ticket = null;
                ticket.Fec_cie_ticket = null;
            }

            DateTime? auxiliando = ticket.Fec_cie_ticket;

            if (ticket.CodGrupo_ticket == 10 && ticket.Fec_tran_ticket != null && auxiliando != null)
            {
                ticket.T_lab_ticket = DiferenciaHoras(ticket.Fec_tran_ticket, auxiliando, ticket.CodTipo == 1 || ticket.CodTipo == 6);
            }
            if (ticket.CodGrupo_ticket != 10 && auxiliando != null)
            {
                ticket.T_lab_ticket = DiferenciaHoras(ticket.Fech_ape_ticket, auxiliando, ticket.CodTipo == 1 || ticket.CodTipo == 6);
            }

            if (u1 == 1)
            {
                listaUpdate.Add(ticket);
            }
            else
            {
                listaNuevas.Add(ticket);
            }

            return ticket;
        }
        private DTOTicket_CA guardarTicket(DTOTicket_CA t, Excel.Range rango, int row, ArrayList listaNuevas, ArrayList listaUpdate)
        {
            DTOTicket_CA ticket = new DTOTicket_CA();
            ticket.CodTicket = t.CodTicket;
            ticket.Secuencia_tarea = 0;
            ticket.CodTipo = t.CodTipo;
            ticket.TipoCA_ticket = t.TipoCA_ticket;
            ticket.Prioridad_ticket = (rango.Cells[row, 6] as Excel.Range).Value2.ToString();

            if (celdaValidar(rango, row, 19))
            {
                ticket.Fec_ultmod_ticket = DateTime.FromOADate(Double.Parse((rango.Cells[row, 19] as Excel.Range).Value2.ToString()));

            }

            int u1 = hayqueactualizarCA(ticket);

            if (u1 == 0)
            {
                return null;
            }

            if (celdaValidar(rango, row, 9))
            {
                ticket.Fech_ape_ticket = DateTime.FromOADate(Double.Parse((rango.Cells[row, 9] as Excel.Range).Value2.ToString()));
            }

            if (celdaValidar(rango, row, 14))
            {
                ticket.Fec_res_ticket = DateTime.FromOADate(Double.Parse((rango.Cells[row, 14] as Excel.Range).Value2.ToString()));
            }
            else
            {
                ticket.Fec_res_ticket = null;
            }

            if (celdaValidar(rango, row, 15))
            {
                ticket.Fec_cie_ticket = DateTime.FromOADate(Double.Parse((rango.Cells[row, 15] as Excel.Range).Value2.ToString()));
            }
            else
            {
                ticket.Fec_cie_ticket = null;
            }

            if (celdaValidar(rango, row, 20))
            {
                ticket.Solicitante_ticket = (rango.Cells[row, 20] as Excel.Range).Value2.ToString();
            }

            if (celdaValidar(rango, row, 23))
            {
                ticket.Id_emp_usu = buscarTablasCA("U", (rango.Cells[row, 23] as Excel.Range).Value2.ToString());
            }
            else
            {
                ticket.Id_emp_usu = 103;
            }

            if (celdaValidar(rango, row, 25))
            {
                ticket.CodSedeUsuario = buscarTablasCA("S", (rango.Cells[row, 25] as Excel.Range).Value2.ToString());
            }
            else
            {
                ticket.CodSedeUsuario = 287;
            }

            if (celdaValidar(rango, row, 26))
            {
                ticket.CodSede = buscarTablasCA("S", (rango.Cells[row, 26] as Excel.Range).Value2.ToString());
            }
            else
            {
                ticket.CodSede = 287;
            }

            if (ticket.CodTipo == 5 && ticket.CodSede != 287)
            {
                ticket.Id_emp_usu = buscarEmpresaxSede(ticket.CodSede.Value);
            }


            if (celdaValidar(rango, row, 27))
            {
                ticket.NomEspecialista = (rango.Cells[row, 27] as Excel.Range).Value2.ToString();
            }
            else
            {
                ticket.NomEspecialista = "";
            }

            if (celdaValidar(rango, row, 28))
            {
                ticket.Repor_por_ticket = buscarTablasCA("R", (rango.Cells[row, 28] as Excel.Range).Value2.ToString());
            }
            else
            {
                ticket.Repor_por_ticket = null;
            }

            if (celdaValidar(rango, row, 29))
            {
                ticket.CodCategoriaCA_ticket = buscarTablasCA("C", (rango.Cells[row, 29] as Excel.Range).Value2.ToString());
                if (definirsiesConsulta((rango.Cells[row, 29] as Excel.Range).Value2.ToString()))
                {
                    ticket.CodTipo = 6; //Es una consulta
                }
            }
            else
            {
                ticket.CodCategoriaCA_ticket = 120;
            }

            if (celdaValidar(rango, row, 31))
            {
                ticket.CodGrupo_propi_ticket = buscarTablasCA("G", (rango.Cells[row, 31] as Excel.Range).Value2.ToString());
            }
            else
            {
                ticket.CodGrupo_propi_ticket = 14;
            }

            if (celdaValidar(rango, row, 32))
            {
                ticket.CodGrupo_ticket = buscarTablasCA("G", (rango.Cells[row, 32] as Excel.Range).Value2.ToString());
            }
            else
            {
                ticket.CodGrupo_ticket = 14;
            }

            if (celdaValidar(rango, row, 34))
            {
                ticket.Resu_ticket = (rango.Cells[row, 34] as Excel.Range).Value2.ToString();
            }
            else
            {
                ticket.Resu_ticket = "";
            }

            if (celdaValidar(rango, row, 35))
            {
                ticket.Des_ticket = (rango.Cells[row, 35] as Excel.Range).Value2.ToString();
            }
            else
            {
                ticket.Des_ticket = "";
            }

            if (celdaValidar(rango, row, 36))
            {
                ticket.CodEstado = buscarTablasCA("E", (rango.Cells[row, 36] as Excel.Range).Value2.ToString());
            }
            else
            {
                ticket.CodEstado = null;
            }

            if (celdaValidar(rango, row, 37))
            {
                ticket.Pendiente_por_ticket = (rango.Cells[row, 37] as Excel.Range).Value2.ToString();
            }
            else
            {
                ticket.Pendiente_por_ticket = null;
            }

            if (celdaValidar(rango, row, 38))
            {
                ticket.CodMetodo = buscarTablasCA("M", (rango.Cells[row, 38] as Excel.Range).Value2.ToString());
            }
            else
            {
                ticket.CodMetodo = 1;
            }

            if (celdaValidar(rango, row, 40))
            {
                ticket.Met_resol_ticket = (rango.Cells[row, 40] as Excel.Range).Value2.ToString();
            }
            else
            {
                ticket.Met_resol_ticket = null;
            }

            if (celdaValidar(rango, row, 41))
            {
                ticket.Socie_sap = (rango.Cells[row, 41] as Excel.Range).Value2.ToString();
            }
            else
            {
                ticket.Socie_sap = null;
            }

            if (celdaValidar(rango, row, 42))
            {
                ticket.Comp_sap = (rango.Cells[row, 42] as Excel.Range).Value2.ToString();
            }
            else
            {
                ticket.Comp_sap = null;
            }

            if (celdaValidar(rango, row, 43))
            {
                ticket.Transaccion = (rango.Cells[row, 43] as Excel.Range).Value2.ToString();
            }
            else
            {
                ticket.Transaccion = null;
            }

            //Validar Transacción
            if (esCategoriaSeleccionada(ticket.CodGrupo_ticket.Value))
            {
                if (ticket.Transaccion != null)
                {
                    ticket.GrupoTrans = obtenerGrupoTransaccion(ticket.CodGrupo_ticket.Value, ticket.Transaccion);
                }
                else
                {
                    ticket.GrupoTrans = null;
                }
            }
            else
            {
                ticket.GrupoTrans = null;
            }


            if (celdaValidar(rango, row, 44))
            {
                ticket.Cancelado_por_ticket = (rango.Cells[row, 44] as Excel.Range).Value2.ToString();
            }
            else
            {
                ticket.Cancelado_por_ticket = null;
            }

            if (celdaValidar(rango, row, 45))
            {
                ticket.Solucion_ticket = (rango.Cells[row, 45] as Excel.Range).Value2.ToString();
            }
            else
            {
                ticket.Solucion_ticket = "";
            }

            if (celdaValidar(rango, row, 49))
            {
                ticket.Causa_ticket = (rango.Cells[row, 49] as Excel.Range).Value2.ToString();
            }
            else
            {
                ticket.Causa_ticket = "";
            }

            DateTime? auxiliando = ticket.Fec_cie_ticket;

            if (ticket.CodGrupo_ticket == 10 && ticket.Fec_tran_ticket != null && auxiliando != null)
            {
                ticket.T_lab_ticket = DiferenciaHoras(ticket.Fec_tran_ticket, auxiliando, ticket.CodTipo == 1 || ticket.CodTipo == 6);
            }
            if (ticket.CodGrupo_ticket != 10 && auxiliando != null)
            {
                ticket.T_lab_ticket = DiferenciaHoras(ticket.Fech_ape_ticket, auxiliando, ticket.CodTipo == 1 || ticket.CodTipo == 6);
            }

            if (u1 == 1)
            {
                listaUpdate.Add(ticket);
            }
            else
            {
                listaNuevas.Add(ticket);
            }

            return ticket;

        }
        
        //Sede Nueva
        private void button3_Click(object sender, EventArgs e)
        {
            DTOSede sede = new DTOSede();
            DAOSede dao = new DAOSede();

            sede.NomSede = txtnomsede.Text;
            sede.CodEmpresa = (int)Int32.Parse(cboempresaSE.SelectedValue.ToString());
            sede.UbiSede = cboubicacion.SelectedItem.ToString().Substring(0, 1);
            sede.PaisSede = txtpais.Text;
            sede.DepaSede = txtdepartamento.Text;
            sede.DistSede = txtdistrito.Text;
            sede.DirecSede = txtdireccion.Text;


            if (rbtonsiteSI.Checked)
            {
                sede.OnSiteSede = 1;
            }
            if (rbtonsiteNO.Checked)
            {
                sede.OnSiteSede = 0;
            }


            sede.CodSede = (int)Int32.Parse(dao.selectSedeMayor().Tables[0].Rows[0][0].ToString()) + 1;

            dao.InsertSede(sede);
            SedeCA.Add(sede);
            sedesNuevas.RemoveAt(0);

            cboempresaSE.SelectedIndex = 0;
            rbtonsiteSI.Checked = true;
            rbtonsiteNO.Checked = false;
            cboubicacion.SelectedIndex = 0;
            txtnomsede.Text = "";
            txtpais.Text = "";
            txtdepartamento.Text = "";
            txtdistrito.Text = "";
            txtdireccion.Text = "";


            if (sedesNuevas.Count > 0)
            {
                txtnomsede.Text = sedesNuevas[0].ToString();
            }
            else
            {
                groupBox1.Visible = false;
                this.Size = new System.Drawing.Size(610, 148);
                backgroundWorker4.RunWorkerAsync();
            }


        }

        private int buscarEmpresaxSede(int codsede)
        {
            int empresa = -1;

            DAOSede daosede = new DAOSede();
            DTOSede sede = new DTOSede();
            sede.CodSede = codsede;
            DataSet ds = daosede.selectSede2(sede);

            empresa = (int)Int32.Parse(ds.Tables[0].Rows[0]["codEmpresa"].ToString());

            return empresa;
        }
        #endregion

        //Logica Tickets Maximo
        #region

        private void button8_Click(object sender, EventArgs e)
        {
            string Chosen_File = "";
            openFileDialog1.Title = "Ingresa la Solicitud";
            openFileDialog1.FileName = "";
            openFileDialog1.Filter = "Archivos Excel *.xls|*.xls*";
            openFileDialog1.ShowDialog();

            Chosen_File = openFileDialog1.FileName;

            if (Chosen_File == "")
            {
                MessageBox.Show("No ha Seleccionado ningun Archivo");
            }
            else
            {

                //Sentencias Excel
                label1.Text = "Cargando Tablas Auxiliares....";
                rellenarTablasMaximo();
                object misValue = System.Reflection.Missing.Value;
                xlApp = new Excel.Application();
                xlWorkBook = xlApp.Workbooks.Open(Chosen_File, 0, true, 5, "", "", true, Microsoft.Office.Interop.Excel.XlPlatform.xlWindows, "\t", false, false, 0, true, 1, 0);
                int lastRow = -1;
                foreach (Excel.Worksheet element in xlWorkBook.Worksheets)
                {

                    lastRow = element.get_Range("A" + element.Rows.Count).get_End(Excel.XlDirection.xlUp).Row;
                }
                //lastRow = lastRow - 6;

                progressBar1.Value = 0;
                label1.Text = "Leyendo las filas  0/" + lastRow;
                progressBar1.Maximum = lastRow;

                nuevosTickets = new ArrayList();
                oldTickets = new ArrayList();

                backgroundWorker2.RunWorkerAsync();

            }
        }
        

        private void backgroundWorker2_DoWork(object sender, DoWorkEventArgs e)
        {
            DateTime start = DateTime.Now;
            e.Result = "";
            object misValue = System.Reflection.Missing.Value;

            foreach (Excel.Worksheet element in xlWorkBook.Worksheets)
            {
                //progressBar1.Value = 0;
                //Excel.Range rng = (Excel.Range)element.get_Range("A7", "A99999");
                int lastRow = -1;
                lastRow = element.get_Range("A" + element.Rows.Count).get_End(Excel.XlDirection.xlUp).Row;

                //Leyendo Filas
                if (lastRow != -1)
                {
                    //Definiendo Rango

                    Excel.Range rango = (Excel.Range)element.get_Range("A2", "BS" + lastRow);

                    //label1.Text = "Leyendo las filas";
                    CantidadActualizar = rango.Rows.Count;
                    //progressBar1.Maximum = rango.Rows.Count;

                    //Recorriendo Rango de Datos
                    for (int row = 1; row <= rango.Rows.Count; row++)
                    {
                        DTOTicketMAX ticket = new DTOTicketMAX();
                        DAOTicketMAX daoTicket = new DAOTicketMAX();

                        if ((rango.Cells[row, 1] as Excel.Range).Value2 != null)
                        {
                            ticket.NumTicket = (rango.Cells[row, 1] as Excel.Range).Value2.ToString();
                        }

                        if ((rango.Cells[row, 19] as Excel.Range).Value2 != null)
                        {
                            ticket.FechaEstadoActual = DateTime.FromOADate(Double.Parse((rango.Cells[row, 19] as Excel.Range).Value2.ToString()));
                        }

                        int existencia = ActualizarTicketMax(ticket);

                        if (existencia != 0)
                        {
                            if ((rango.Cells[row, 2] as Excel.Range).Value2 != null)
                            {
                                ticket.Clase = (rango.Cells[row, 2] as Excel.Range).Value2.ToString();
                            }
                            if ((rango.Cells[row, 3] as Excel.Range).Value2 != null)
                            {
                                ticket.Id_estado = buscarTablasMaximo("E", (rango.Cells[row, 3] as Excel.Range).Value2.ToString());
                            }
                            if ((rango.Cells[row,4] as Excel.Range).Value2 != null)
                            {
                                ticket.Id_fuente = buscarTablasMaximo("F", (rango.Cells[row, 4] as Excel.Range).Value2.ToString());
                            }
                            if ((rango.Cells[row, 5] as Excel.Range).Value2 != null)
                            {
                                ticket.Usu = (rango.Cells[row, 5] as Excel.Range).Value2.ToString();
                            }

                            if ((rango.Cells[row, 24] as Excel.Range).Value2 != null)
                            {
                                ticket.Id_servicio = buscarTablasMaximo("S", (rango.Cells[row, 24] as Excel.Range).Value2.ToString());
                            }

                            if ((rango.Cells[row, 6] as Excel.Range).Value2 != null)
                            {
                                ticket.Asignado = (rango.Cells[row, 6] as Excel.Range).Value2.ToString();
                            }
                            
                            if ((rango.Cells[row, 7] as Excel.Range).Value2 != null)
                            {
                                ticket.Id_cliente = buscarTablasMaximo("C", (rango.Cells[row, 7] as Excel.Range).Value2.ToString());
                            }

                            if ((rango.Cells[row, 8] as Excel.Range).Value2 != null)
                            {
                                ticket.Ci = (rango.Cells[row, 8] as Excel.Range).Value2.ToString();
                            }

                            if ((rango.Cells[row, 9] as Excel.Range).Value2 != null)
                            {
                                ticket.Usu_informo = (rango.Cells[row, 9] as Excel.Range).Value2.ToString();
                            }

                            if ((rango.Cells[row, 10] as Excel.Range).Value2 != null)
                            {
                                ticket.Resumen = limpiarTexto((rango.Cells[row, 10] as Excel.Range).Value2.ToString());
                            }
                            if ((rango.Cells[row, 11] as Excel.Range).Value2 != null)
                            {
                                ticket.Detallle = limpiarTexto((rango.Cells[row, 11] as Excel.Range).Value2.ToString());
                            }

                            if ((rango.Cells[row, 12] as Excel.Range).Value2 != null)
                            {
                                ticket.Id_prioridad = (int)Int32.Parse((rango.Cells[row, 12] as Excel.Range).Value2.ToString());
                            }
                            if ((rango.Cells[row, 15] as Excel.Range).Value2 != null)
                            {
                                ticket.FechaInformada = DateTime.FromOADate(Double.Parse((rango.Cells[row, 15] as Excel.Range).Value2.ToString()));
                            }
                            if ((rango.Cells[row, 16] as Excel.Range).Value2 != null)
                            {
                                ticket.InicioIndisposicion = DateTime.FromOADate(Double.Parse((rango.Cells[row, 16] as Excel.Range).Value2.ToString()));
                            }
                            if ((rango.Cells[row, 17] as Excel.Range).Value2 != null)
                            {
                                ticket.InicioReal = DateTime.FromOADate(Double.Parse((rango.Cells[row, 17] as Excel.Range).Value2.ToString()));
                            }
                            if ((rango.Cells[row, 18] as Excel.Range).Value2 != null)
                            {
                                ticket.FinReal = DateTime.FromOADate(Double.Parse((rango.Cells[row, 18] as Excel.Range).Value2.ToString()));
                            }

                            if ((rango.Cells[row, 20] as Excel.Range).Value2 != null)
                            {
                                ticket.FechaCierre = DateTime.FromOADate(Double.Parse((rango.Cells[row, 20] as Excel.Range).Value2.ToString()));
                            }

                            if ((rango.Cells[row, 21] as Excel.Range).Value2 != null)
                            {
                                ticket.Sintoma = limpiarTexto((rango.Cells[row, 21] as Excel.Range).Value2.ToString());
                            }
                            if ((rango.Cells[row, 22] as Excel.Range).Value2 != null)
                            {
                                ticket.Causa = limpiarTexto((rango.Cells[row, 22] as Excel.Range).Value2.ToString());
                            }
                            if ((rango.Cells[row, 23] as Excel.Range).Value2 != null)
                            {
                                ticket.Solucion = limpiarTexto((rango.Cells[row, 23] as Excel.Range).Value2.ToString());
                            }
                            if ((rango.Cells[row, 13] as Excel.Range).Value2 != null)
                            {
                                ticket.TicketPadre = (rango.Cells[row, 13] as Excel.Range).Value2.ToString();
                            }
                            if ((rango.Cells[row, 14] as Excel.Range).Value2 != null)
                            {
                                ticket.ClaseTicketPadre = (rango.Cells[row, 14] as Excel.Range).Value2.ToString();
                            }

                            ticket.TiempoInicio = DiferenciaHoras(ticket.InicioIndisposicion, ticket.InicioReal, ticket.Clase != "INCIDENT");
                            ticket.TiempoSolucion = DiferenciaHoras(ticket.InicioReal, ticket.FinReal, ticket.Clase != "INCIDENT");

                            if (existencia == -1)
                            {
                                nuevosTickets.Add(ticket);
                            }
                            else
                            {
                                oldTickets.Add(ticket);
                            }
                        }

                        backgroundWorker2.ReportProgress(row - 1, DateTime.Now);
                        if (backgroundWorker2.CancellationPending)
                        {
                            e.Cancel = true;
                            return;
                        }

                    }
                }
                releaseObject(element);
            }

            xlWorkBook.Close(false, misValue, misValue);
            xlApp.Quit();

            releaseObject(xlWorkBook);
            releaseObject(xlApp);

            TimeSpan duration = DateTime.Now - start;
            e.Result = "Duracion: " + duration.TotalMinutes.ToString() + "m.";


        }
        private void backgroundWorker2_ProgressChanged(object sender, ProgressChangedEventArgs e)
        {
            progressBar1.Value = e.ProgressPercentage;
            label1.Text = "Leyendo las filas  " + e.ProgressPercentage + "/" + CantidadActualizar;
            DateTime time = Convert.ToDateTime(e.UserState);
        }
        private void backgroundWorker2_RunWorkerCompleted(object sender, RunWorkerCompletedEventArgs e)
        {
            if (e.Cancelled)
            {
                MessageBox.Show("The task has been cancelled");
            }
            else if (e.Error != null)
            {
                MessageBox.Show("Error. Details: " + (e.Error as Exception).ToString());
            }
            else
            {
                MessageBox.Show("Lectura Completa!. Results: " + e.Result.ToString());

                //Actualizando oldTickets
                label1.Text = "Actualizando " + oldTickets.Count + " Tickets";
                progressBar1.Value = 0;
                progressBar1.Maximum = oldTickets.Count;
                CantidadActualizar = oldTickets.Count;

                backgroundWorker9.RunWorkerAsync();

                //DAOTicketMAX daoMAX = new DAOTicketMAX();
                //foreach (DTOTicketMAX ticket in oldTickets)
                //{
                //    daoMAX.UpdateTicketMAX(ticket);
                //    progressBar1.Value += 1;
                //}

                //Insertando nuevos tickets
                //label1.Text = "Insertando " + nuevosTickets.Count + " Tickets";
                //progressBar1.Value = 0;
                //progressBar1.Maximum = nuevosTickets.Count;

                //foreach (DTOTicketMAX ticket in nuevosTickets)
                //{
                //    daoMAX.InsertTicketMAX(ticket);
                //    progressBar1.Value += 1;
                //}

                //MessageBox.Show("Carga Completa");

            }
        }

        private void backgroundWorker9_DoWork(object sender, DoWorkEventArgs e)
        {
            DAOTicketMAX daoMAX = new DAOTicketMAX();
            int aux = 0;
            DateTime start = DateTime.Now;
            e.Result = "";

            foreach (DTOTicketMAX ticket in oldTickets)
            {
                daoMAX.UpdateTicketMAX(ticket);

                backgroundWorker9.ReportProgress(aux, DateTime.Now);
                if (backgroundWorker9.CancellationPending)
                {
                    e.Cancel = true;
                    return;
                }

                aux++;

            }

            TimeSpan duration = DateTime.Now - start;
            e.Result = "Duracion: " + duration.TotalMinutes.ToString() + "m.";
        }
        private void backgroundWorker9_ProgressChanged(object sender, ProgressChangedEventArgs e)
        {
            progressBar1.Value = e.ProgressPercentage;
            label1.Text = "Actualizando  " + e.ProgressPercentage + "/" + CantidadActualizar;
            DateTime time = Convert.ToDateTime(e.UserState);
        }
        private void backgroundWorker9_RunWorkerCompleted(object sender, RunWorkerCompletedEventArgs e)
        {
             if (e.Cancelled)
            {
                MessageBox.Show("The task has been cancelled");
            }
             else if (e.Error != null)
             {
                 MessageBox.Show("Error. Details: " + (e.Error as Exception).ToString());
             }
             else
             {
                 label1.Text = "Insertando " + nuevosTickets.Count + " Tickets";
                 progressBar1.Value = 0;
                 progressBar1.Maximum = nuevosTickets.Count;
                 CantidadActualizar = nuevosTickets.Count;

                 backgroundWorker10.RunWorkerAsync();
             }
        }

        private void backgroundWorker10_DoWork(object sender, DoWorkEventArgs e)
        {
            DAOTicketMAX daoMAX = new DAOTicketMAX();
            int aux = 0;
            DateTime start = DateTime.Now;
            e.Result = "";

            foreach (DTOTicketMAX ticket in nuevosTickets)
            {
                daoMAX.InsertTicketMAX(ticket);

                backgroundWorker10.ReportProgress(aux, DateTime.Now);
                if (backgroundWorker10.CancellationPending)
                {
                    e.Cancel = true;
                    return;
                }

                aux++;
            }

            TimeSpan duration = DateTime.Now - start;
            e.Result = "Duracion: " + duration.TotalMinutes.ToString() + "m.";

        }
        private void backgroundWorker10_ProgressChanged(object sender, ProgressChangedEventArgs e)
        {
            progressBar1.Value = e.ProgressPercentage;
            label1.Text = "Insertando  " + e.ProgressPercentage + "/" + CantidadActualizar;
            DateTime time = Convert.ToDateTime(e.UserState);
        }
        private void backgroundWorker10_RunWorkerCompleted(object sender, RunWorkerCompletedEventArgs e)
        {
            if (e.Cancelled)
            {
                MessageBox.Show("The task has been cancelled");
            }
            else if (e.Error != null)
            {
                MessageBox.Show("Error. Details: " + (e.Error as Exception).ToString());
            }
            else
            {
                MessageBox.Show("Carga Completa");
            }
        }


        //Métodos Auxiliares - Tickets Maximo
        private DateTime corregirFechas(string fecha)
        {
            //int x = fecha.IndexOf(".");
            string entero = fecha.Substring(0, fecha.IndexOf(".") + 1);
            string decimales = fecha.Substring(fecha.IndexOf(".") + 1);
            if (decimales.Length < 3)
            {
                if (decimales.Length == 2)
                {
                    decimales = String.Concat(decimales, "0");
                }
                if (decimales.Length == 1)
                {
                    decimales = String.Concat(decimales, "00");
                }
                if (decimales.Length == 0) { decimales = String.Concat(decimales, "000"); }

            }
            fecha = String.Concat(entero, decimales);
            DateTime myDate = new DateTime(1992, 12, 24);
            try
            {
                myDate = DateTime.ParseExact(fecha, "yyyy-MM-dd HH:mm:ss.fff",
                                      System.Globalization.CultureInfo.InvariantCulture);
            }
            catch (Exception e)
            {
                MessageBox.Show("Error la fecha es: " + fecha);
            }


            return myDate;
        }
        private int ActualizarTicketMax(DTOTicketMAX ticket)
        {
            DAOTicketMAX dao = new DAOTicketMAX();
            DataTable dt = dao.MostrarTicketMAX(ticket.NumTicket);

            if (dt.Rows.Count > 0)
            {
                if (dt.Rows[0]["actual"].ToString() != "" || dt.Rows[0]["actual"] == null)
                {
                    DateTime fechaEstadoActual = DateTime.ParseExact(dt.Rows[0]["actual"].ToString(), "yyyy-MM-dd HH:mm:ss", null);
                    if (fechaEstadoActual < ticket.FechaEstadoActual)
                    {
                        return 1;
                    }
                    else
                    {
                        return 0;
                    }
                }
                else
                {
                    return 1;
                }
            }
            else
            {
                return -1;
            }
        }
        private void rellenarTablasMaximo()
        {
            DAOClienteMAX daocliente = new DAOClienteMAX();
            DataTable dtcliente = daocliente.MostrarDatos();
            foreach (DataRow rw in dtcliente.Rows)
            {
                DTOClienteMAX cliente = new DTOClienteMAX();
                cliente.Id_cliente = (int)Int32.Parse(rw["id_cliente"].ToString());
                cliente.Nom_cliente = rw["nom_cliente"].ToString();
                cliente.Cod_clente = rw["cod_cliente"].ToString();
                ClientesMAX.Add(cliente);
            }

            DAOFuenteMAX daofuente = new DAOFuenteMAX();
            DataTable dtfuente = daofuente.MostrarDatos();
            foreach (DataRow rw in dtfuente.Rows)
            {
                DTOFuenteMAX fuente = new DTOFuenteMAX();
                fuente.Id_fuente = (int)Int32.Parse(rw["id_fuente"].ToString());
                fuente.Nom_fuente = rw["nom_fuente"].ToString();
                FuenteMAX.Add(fuente);
            }

            DAOServicioMAX daoservicio = new DAOServicioMAX();
            DataTable dtservicio = daoservicio.MostrarDatos();
            foreach (DataRow rw in dtservicio.Rows)
            {
                DTOServicioMAX servicio = new DTOServicioMAX();
                servicio.Id_servicio = (int)Int32.Parse(rw["id_servicio"].ToString());
                servicio.Nom_servicio = rw["nom_servicio"].ToString();
                ServiciosMAX.Add(servicio);
            }

            DAOEstadoMAX daoestado = new DAOEstadoMAX();
            DataTable dtestado = daoestado.MostrarDatos();
            foreach (DataRow rw in dtestado.Rows)
            {
                DTOEstadoMAX estado = new DTOEstadoMAX();
                estado.Id_estado = (int)Int32.Parse(rw["id_estado"].ToString());
                estado.Nom_estado = rw["nom_estado"].ToString();
                EstadoMAX.Add(estado);
            }

            DAOClienteCambio daoclienteCambio = new DAOClienteCambio();
            DataTable dtclienteCambio = daoclienteCambio.MostrarDatos();
            foreach (DataRow rw in dtclienteCambio.Rows)
            {
                DTOClienteCambio clienteC = new DTOClienteCambio();
                clienteC.Id_cliente = (int)Int32.Parse(rw["id_cliente_cambio"].ToString());
                clienteC.Nom_cliente = rw["nom_cliente_cambio"].ToString();
                clienteC.Cod_clente = rw["cod_cliente_cambio"].ToString();
                ClienteCambio.Add(clienteC);
            }

            DAOEstadoCambio daoestadoC = new DAOEstadoCambio();
            DataTable dtestadoC = daoestadoC.MostrarDatos();
            foreach (DataRow rw in dtestadoC.Rows)
            {
                DTOEstadoCambio estado = new DTOEstadoCambio();
                estado.Id_estado = (int)Int32.Parse(rw["id_estadoCambio"].ToString());
                estado.Nom_estado = rw["nom_estadoCambio"].ToString();
                estado.Nom_abreviadoCambio = rw["nom_abreviadoCambio"].ToString();
                EstadoCambio.Add(estado);
            }

            DAOServicioCambio daoservicioC = new DAOServicioCambio();
            DataTable dtservicioC = daoservicioC.MostrarDatos();
            foreach (DataRow rw in dtservicioC.Rows)
            {
                DTOServicioCambio servicio = new DTOServicioCambio();
                servicio.Id_servicio = (int)Int32.Parse(rw["id_servicio_cambio"].ToString());
                servicio.Nom_servicio = rw["nom_servicio_cambio"].ToString();
                servicio.Detal_servicio = rw["cod_servicio_cambio"].ToString();
            }
        }
        private int? buscarTablasMaximo(string tipo, string nombre)
        {
            int? j = -1;

            nombre = nombre.Trim().ToUpper();

            switch (tipo)
            {
                case "C":
                    foreach (DTOClienteMAX cliente in ClientesMAX)
                    {
                        if (cliente.Nom_cliente.Trim().ToUpper().Equals(nombre))
                        {
                            return cliente.Id_cliente;
                        }
                    }
                    if (j == -1)
                    {
                        DAOClienteMAX daocliente = new DAOClienteMAX();
                        DTOClienteMAX nuevoCliente = new DTOClienteMAX();
                        nuevoCliente.Id_cliente = (int)Int32.Parse(daocliente.selectClienteMayor().Tables[0].Rows[0][0].ToString()) + 1;
                        nuevoCliente.Nom_cliente = nombre;
                        daocliente.InsertCliente(nuevoCliente);
                        ClientesMAX.Add(nuevoCliente);
                        j = nuevoCliente.Id_cliente;
                    }
                    break;
                case "F":
                    foreach (DTOFuenteMAX fuente in FuenteMAX)
                    {
                        if (fuente.Nom_fuente.Trim().ToUpper().Equals(nombre))
                        {
                            return fuente.Id_fuente;
                        }
                    }
                    if (j == -1)
                    {
                        DAOFuenteMAX daofuente = new DAOFuenteMAX();
                        DTOFuenteMAX nuevaFuente = new DTOFuenteMAX();
                        nuevaFuente.Id_fuente = (int)Int32.Parse(daofuente.selectFuenteMayor().Tables[0].Rows[0][0].ToString()) + 1;
                        nuevaFuente.Nom_fuente = nombre;
                        daofuente.InsertFuente(nuevaFuente);
                        FuenteMAX.Add(nuevaFuente);
                        j = nuevaFuente.Id_fuente;
                    }
                    break;
                case "S":
                    foreach (DTOServicioMAX servicio in ServiciosMAX)
                    {
                        if (servicio.Nom_servicio.Trim().ToUpper().Equals(nombre))
                        {
                            return servicio.Id_servicio;
                        }
                    }
                    if (j == -1)
                    {
                        DAOServicioMAX daoservicio = new DAOServicioMAX();
                        DTOServicioMAX nuevoservicio = new DTOServicioMAX();
                        nuevoservicio.Id_servicio = (int)Int32.Parse(daoservicio.selectServicioMayor().Tables[0].Rows[0][0].ToString()) + 1;
                        nuevoservicio.Nom_servicio = nombre;
                        daoservicio.InsertServicio(nuevoservicio);
                        ServiciosMAX.Add(nuevoservicio);
                        j = nuevoservicio.Id_servicio;
                    }
                    break;
                case "E":
                    foreach (DTOEstadoMAX estado in EstadoMAX)
                    {
                        if (estado.Nom_estado.Trim().ToUpper().Equals(nombre))
                        {
                            return estado.Id_estado;
                        }
                    }
                    if (j == -1)
                    {
                        DAOEstadoMAX daoestado = new DAOEstadoMAX();
                        DTOEstadoMAX nuevoestado = new DTOEstadoMAX();
                        nuevoestado.Id_estado = (int)Int32.Parse(daoestado.selectEstadoMayor().Tables[0].Rows[0][0].ToString()) + 1;
                        nuevoestado.Nom_estado = nombre;
                        daoestado.InsertEstado(nuevoestado);
                        EstadoMAX.Add(nuevoestado);
                        j = nuevoestado.Id_estado;
                    }
                    break;
            }
            return j;
        }



        #endregion

        //Lógica Cambios Maximo
        #region
        private void button9_Click(object sender, EventArgs e)
        {

            string Chosen_File = "";
            openFileDialog1.Title = "Ingresa la Solicitud";
            openFileDialog1.FileName = "";
            openFileDialog1.Filter = "Archivos Excel *.xls|*.xls*";
            openFileDialog1.ShowDialog();

            Chosen_File = openFileDialog1.FileName;

            if (Chosen_File == "")
            {
                MessageBox.Show("No ha Seleccionado ningun Archivo");
            }
            else
            {
                //Sentencias Excel
                label1.Text = "Cargando Tablas Auxiliares....";
                rellenarTablasMaximo();
                object misValue = System.Reflection.Missing.Value;
                xlApp = new Excel.Application();
                xlWorkBook = xlApp.Workbooks.Open(Chosen_File, 0, true, 5, "", "", true, Microsoft.Office.Interop.Excel.XlPlatform.xlWindows, "\t", false, false, 0, true, 1, 0);



                int lastRow = -1;
                foreach (Excel.Worksheet element in xlWorkBook.Worksheets)
                {

                    lastRow = element.get_Range("B" + element.Rows.Count).get_End(Excel.XlDirection.xlUp).Row;
                }
                progressBar1.Value = 0;
                label1.Text = "Leyendo las filas  0/" + lastRow;
                progressBar1.Maximum = lastRow;

                //Leyendo filas
                nuevosTickets = new ArrayList();
                oldTickets = new ArrayList();

                backgroundWorker3.RunWorkerAsync();
            }
        }

        private void backgroundWorker3_DoWork(object sender, DoWorkEventArgs e)
        {
            DateTime start = DateTime.Now;
            e.Result = "";
            object misValue = System.Reflection.Missing.Value;

            foreach (Excel.Worksheet element in xlWorkBook.Worksheets)
            {
                progressBar1.Value = 0;
                //Excel.Range rng = (Excel.Range)element.get_Range("A7", "A99999");
                int lastRow = -1;
                lastRow = element.get_Range("A" + element.Rows.Count).get_End(Excel.XlDirection.xlUp).Row;

                //Leyendo Filas
                if (lastRow != -1)
                {
                    //Definiendo Rango

                    Excel.Range rango = (Excel.Range)element.get_Range("A2", "BO" + lastRow);
                    CantidadActualizar = rango.Rows.Count;

                    //Recorriendo Rango de Datos
                    for (int row = 1; row <= rango.Rows.Count; row++)
                    {
                        DTOTicketCambio ticket = new DTOTicketCambio();
                        DAOTicketCambio daoTicket = new DAOTicketCambio();

                        if ((rango.Cells[row, 1] as Excel.Range).Value2 != null)
                        {
                            ticket.NumTicket_cambio = (rango.Cells[row, 1] as Excel.Range).Value2.ToString();
                        }
                        if ((rango.Cells[row, 9] as Excel.Range).Value2 != null)
                        {
                            ticket.FechaEstAct_cambio = corregirFechas((rango.Cells[row, 9] as Excel.Range).Value2.ToString());
                        }

                        int existencia = ActualizarTicketMaxCambio(ticket);

                        if (existencia != 0)
                        {

                            if ((rango.Cells[row, 4] as Excel.Range).Value2 != null)
                            {
                                ticket.Id_cliente_cambio = buscarTablasCambio("C", (rango.Cells[row, 4] as Excel.Range).Value2.ToString());
                            }
                            if ((rango.Cells[row, 5] as Excel.Range).Value2 != null)
                            {
                                ticket.Plataforma_cambio = (rango.Cells[row, 5] as Excel.Range).Value2.ToString();
                            }
                            if ((rango.Cells[row, 6] as Excel.Range).Value2 != null)
                            {
                                ticket.Id_servicio_cambio = buscarTablasCambio("S", (rango.Cells[row, 6] as Excel.Range).Value2.ToString());
                            }
                            if ((rango.Cells[row, 7] as Excel.Range).Value2 != null)
                            {
                                ticket.Asignado_cambio = (rango.Cells[row, 7] as Excel.Range).Value2.ToString();
                            }
                            if ((rango.Cells[row, 8] as Excel.Range).Value2 != null)
                            {
                                ticket.Id_estadoCambio = buscarTablasCambio("E", (rango.Cells[row, 8] as Excel.Range).Value2.ToString());
                            }

                            if ((rango.Cells[row, 10] as Excel.Range).Value2 != null)
                            {
                                ticket.Resumen_cambio = (rango.Cells[row, 10] as Excel.Range).Value2.ToString();
                            }
                            if ((rango.Cells[row, 11] as Excel.Range).Value2 != null)
                            {
                                ticket.Planificacion_cambio = (int)Int32.Parse((rango.Cells[row, 11] as Excel.Range).Value2.ToString());
                            }
                            if ((rango.Cells[row, 13] as Excel.Range).Value2 != null)
                            {
                                ticket.Categoria_cambio = (rango.Cells[row, 13] as Excel.Range).Value2.ToString();
                            }
                            if ((rango.Cells[row, 15] as Excel.Range).Value2 != null)
                            {
                                ticket.Riesgo_cambio = (int)Int32.Parse((rango.Cells[row, 15] as Excel.Range).Value2.ToString());
                            }
                            if ((rango.Cells[row, 17] as Excel.Range).Value2 != null)
                            {
                                ticket.FechaReportada_cambio = corregirFechas((rango.Cells[row, 17] as Excel.Range).Value2.ToString());
                            }
                            if ((rango.Cells[row, 19] as Excel.Range).Value2 != null)
                            {
                                ticket.InicioPlanificado_cambio = corregirFechas((rango.Cells[row, 19] as Excel.Range).Value2.ToString());
                            }
                            if ((rango.Cells[row, 20] as Excel.Range).Value2 != null)
                            {
                                ticket.FinPlanificado_cambio = corregirFechas((rango.Cells[row, 20] as Excel.Range).Value2.ToString());
                            }
                            if ((rango.Cells[row, 21] as Excel.Range).Value2 != null)
                            {
                                ticket.InicioVentana_cambio = corregirFechas((rango.Cells[row, 21] as Excel.Range).Value2.ToString());
                            }
                            if ((rango.Cells[row, 22] as Excel.Range).Value2 != null)
                            {
                                ticket.FinVentana_cambio = corregirFechas((rango.Cells[row, 22] as Excel.Range).Value2.ToString());
                            }
                            if ((rango.Cells[row, 28] as Excel.Range).Value2 != null)
                            {
                                ticket.ClaseTicketPadre_cambio = (rango.Cells[row, 28] as Excel.Range).Value2.ToString();
                            }
                            if ((rango.Cells[row, 32] as Excel.Range).Value2 != null)
                            {
                                ticket.ProbabilidadAnomalia_cambio = (int)Int32.Parse((rango.Cells[row, 32] as Excel.Range).Value2.ToString());
                            }
                            if ((rango.Cells[row, 33] as Excel.Range).Value2 != null)
                            {
                                ticket.Acelerado_cambio = (rango.Cells[row, 33] as Excel.Range).Value2.ToString();
                            }
                            if ((rango.Cells[row, 34] as Excel.Range).Value2 != null)
                            {
                                ticket.FechaCierre_cambio = corregirFechas((rango.Cells[row, 34] as Excel.Range).Value2.ToString());
                            }
                            if ((rango.Cells[row, 36] as Excel.Range).Value2 != null)
                            {
                                ticket.FechaImplementacion_cambio = corregirFechas((rango.Cells[row, 36] as Excel.Range).Value2.ToString());
                            }

                            if ((rango.Cells[row, 37] as Excel.Range).Value2 != null)
                            {
                                ticket.Observacion_cambio = (rango.Cells[row, 37] as Excel.Range).Value2.ToString();
                            }

                            if (existencia == -1)
                            {
                                nuevosTickets.Add(ticket);
                            }
                            else
                            {
                                oldTickets.Add(ticket);
                            }
                        }

                        backgroundWorker3.ReportProgress(row - 1, DateTime.Now);
                        if (backgroundWorker3.CancellationPending)
                        {
                            e.Cancel = true;
                            return;
                        }

                    }
                }
                releaseObject(element);
            }
            xlWorkBook.Close(false, misValue, misValue);
            xlApp.Quit();

            releaseObject(xlWorkBook);
            releaseObject(xlApp);

            TimeSpan duration = DateTime.Now - start;
            e.Result = "Duracion: " + duration.TotalMinutes.ToString() + "m.";
        }
        private void backgroundWorker3_ProgressChanged(object sender, ProgressChangedEventArgs e)
        {
            progressBar1.Value = e.ProgressPercentage;
            label1.Text = "Leyendo las filas  " + e.ProgressPercentage + "/" + CantidadActualizar;
            DateTime time = Convert.ToDateTime(e.UserState);
        }
        private void backgroundWorker3_RunWorkerCompleted(object sender, RunWorkerCompletedEventArgs e)
        {
            if (e.Cancelled)
            {
                MessageBox.Show("The task has been cancelled");
            }
            else if (e.Error != null)
            {
                MessageBox.Show("Error. Details: " + (e.Error as Exception).ToString());
            }
            else
            {
                MessageBox.Show("Lectura Completa!. Results: " + e.Result.ToString());
                progressBar1.Value = 0;

                label1.Text = "Actualizar las filas  " + 1 + "/" + oldTickets.Count;
                progressBar1.Maximum = oldTickets.Count;
                backgroundWorker6.RunWorkerAsync();
            }
        }

        private void backgroundWorker6_DoWork(object sender, DoWorkEventArgs e)
        {
            DateTime start = DateTime.Now;
            e.Result = "";

            DAOTicketCambio daoC = new DAOTicketCambio();
            int i = 0;
            foreach (DTOTicketCambio ticket in oldTickets)
            {
                int var = daoC.UpdateTicketMAX(ticket);

                if (var == -1)
                {
                    backgroundWorker6.CancelAsync();
                }

                backgroundWorker6.ReportProgress(i, DateTime.Now);

                if (backgroundWorker6.CancellationPending)
                {
                    e.Cancel = true;
                    return;
                }
                i++;
            }

            TimeSpan duration = DateTime.Now - start;
            e.Result = "Duracion: " + duration.TotalMinutes.ToString() + "m.";

        }
        private void backgroundWorker6_ProgressChanged(object sender, ProgressChangedEventArgs e)
        {

            progressBar1.Value = e.ProgressPercentage;
            label1.Text = "Actualizar las filas  " + e.ProgressPercentage + "/" + oldTickets.Count;
            DateTime time = Convert.ToDateTime(e.UserState);
        }
        private void backgroundWorker6_RunWorkerCompleted(object sender, RunWorkerCompletedEventArgs e)
        {
            if (e.Cancelled)
            {
                MessageBox.Show("The task has been cancelled");
            }
            else if (e.Error != null)
            {
                MessageBox.Show("Error. Details: " + (e.Error as Exception).ToString());
            }
            else
            {
                MessageBox.Show("Cargado!. Results: " + e.Result.ToString());

                DAOTicketCambio daoC = new DAOTicketCambio();
                label1.Text = "Insertando " + nuevosTickets.Count + " Tickets";


                progressBar1.Maximum = nuevosTickets.Count;
                progressBar1.Value = 0;

                foreach (DTOTicketCambio ticket in nuevosTickets)
                {
                    daoC.InsertTicketCambio(ticket);
                    progressBar1.Value += 1;
                }
            }
        }
     


        //Métodos Auxiliares - Cambios Maximo
        private int ActualizarTicketMaxCambio(DTOTicketCambio ticket)
        {
            DAOTicketCambio dao = new DAOTicketCambio();
            DataTable dt = dao.MostrarTicketCambio(ticket.NumTicket_cambio);

            if (dt.Rows.Count > 0)
            {
                if (dt.Rows[0]["actual"].ToString() != null)
                {
                    DateTime fechaEstadoActual = DateTime.ParseExact(dt.Rows[0]["actual"].ToString(), "yyyy-MM-dd HH:mm:ss", null);
                    if (fechaEstadoActual < ticket.FechaEstAct_cambio)
                    {
                        return 1;
                    }
                    else
                    {
                        return 0;
                    }
                }
                else
                {
                    return 0;
                }


            }
            else
            {
                return -1;
            }
        }
        private int? buscarTablasCambio(string tipo, string nombre)
        {
            int? j = -1;
            nombre = nombre.Trim().ToUpper();

            switch (tipo)
            {
                case "C":
                    foreach (DTOClienteCambio cliente in ClienteCambio)
                    {
                        if (cliente.Nom_cliente.Trim().ToUpper().Equals(nombre))
                        {
                            return cliente.Id_cliente;
                        }
                    }
                    if (j == -1)
                    {
                        DAOClienteCambio daocliente = new DAOClienteCambio();
                        DTOClienteCambio nuevoCliente = new DTOClienteCambio();
                        nuevoCliente.Id_cliente = (int)Int32.Parse(daocliente.selectClienteMayor().Tables[0].Rows[0][0].ToString()) + 1;
                        nuevoCliente.Nom_cliente = nombre;
                        daocliente.InsertCliente(nuevoCliente);
                        ClienteCambio.Add(nuevoCliente);
                        j = nuevoCliente.Id_cliente;
                    }
                    break;
                case "S":
                    foreach (DTOServicioCambio servicio in ServicioCambio)
                    {
                        if (servicio.Detal_servicio.Trim().ToUpper().Equals(nombre))
                        {
                            return servicio.Id_servicio;
                        }
                    }
                    if (j == -1)
                    {
                        DAOServicioCambio daoservicio = new DAOServicioCambio();
                        DTOServicioCambio nuevoservicio = new DTOServicioCambio();
                        nuevoservicio.Id_servicio = (int)Int32.Parse(daoservicio.selectServicioMayor().Tables[0].Rows[0][0].ToString()) + 1;
                        nuevoservicio.Detal_servicio = nombre;
                        daoservicio.InsertServicio(nuevoservicio);
                        ServicioCambio.Add(nuevoservicio);
                        j = nuevoservicio.Id_servicio;
                    }
                    break;
                case "E":
                    foreach (DTOEstadoCambio estado in EstadoCambio)
                    {
                        if (estado.Nom_abreviadoCambio.Trim().ToUpper().Equals(nombre))
                        {
                            return estado.Id_estado;
                        }
                    }
                    if (j == -1)
                    {
                        DAOEstadoCambio daoestado = new DAOEstadoCambio();
                        DTOEstadoCambio nuevoestado = new DTOEstadoCambio();
                        nuevoestado.Id_estado = (int)Int32.Parse(daoestado.selectEstadoMayor().Tables[0].Rows[0][0].ToString()) + 1;
                        nuevoestado.Nom_abreviadoCambio = nombre;
                        daoestado.InsertEstado(nuevoestado);
                        EstadoCambio.Add(nuevoestado);
                        j = nuevoestado.Id_estado;
                    }
                    break;
            }
            return j;
        }

        #endregion

    }
}



