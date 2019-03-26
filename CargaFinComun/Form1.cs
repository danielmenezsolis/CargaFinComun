using CsvHelper;
using RestSharp;
using RestSharp.Authenticators;
using System;
using System.Collections.Generic;
using System.Data;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Net;
using System.Text;
using System.Windows.Forms;
using Newtonsoft.Json;

namespace CargaFinComun
{
    public partial class Form1 : Form
    {

        public DataTable dataTable { get; set; }
        public int iterations { get; set; }
        public List<string> Resources { get; set; }
        private Results.RootObject RootObjResults { get; set; }
        public int actividadesCorrectas;
        public int actividadesFallidas;

        public Form1()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            OpenFileDialog theDialog = new OpenFileDialog();
            theDialog.Title = "Open Text File";
            theDialog.Filter = "CSV files|*.csv";
            theDialog.InitialDirectory = @"C:\";
            if (theDialog.ShowDialog() == DialogResult.OK)
            {
                string filename = theDialog.FileName;
                string[] filelines = File.ReadAllLines(filename, Encoding.UTF8);
                lblFile.Text = filename;
                lblRecords.Text = (filelines.Count()-1).ToString();
            }

            if (lblFile.Text.Contains("csv"))
            {
                button2.Enabled = true;
                CsvReader csv = new CsvReader(File.OpenText(lblFile.Text));
                csv.Read();
                csv.ReadHeader();

                List<string> headers = csv.Context.HeaderRecord.ToList();

                // Read csv into datatable
                dataTable = new DataTable();
                foreach (string header in headers)
                {
                    dataTable.Columns.Add(new System.Data.DataColumn(header));
                }

                // Check all required columns are present
                if (!headers.Exists(x => x == "idSucursal"))
                {
                    button2.Enabled = false;
                    throw new ArgumentException("idSucursal field not present in input file.");
                }
                else if (!headers.Exists(x => x == "Prioridad"))
                {
                    button2.Enabled = false;
                    throw new ArgumentException("Prioridad field not present in input file.");
                }
                else if (!headers.Exists(x => x == "nombre"))
                {
                    button2.Enabled = false;
                    throw new ArgumentException("nombre field not present in input file.");
                }
                if (!headers.Exists(x => x == "tipoActividad"))
                {
                    button2.Enabled = false;
                    throw new ArgumentException("tipoActividad field not present in input file.");
                }
                if (!headers.Exists(x => x == "date"))
                {
                    button2.Enabled = false;
                    throw new ArgumentException("date field not present in input file.");
                }

                while (csv.Read())
                {
                    DataRow row = dataTable.NewRow();

                    foreach (DataColumn column in dataTable.Columns)
                    {
                        row[column.ColumnName] = csv.GetField(column.DataType, column.ColumnName);
                    }

                    dataTable.Rows.Add(row);
                }

                if (dataTable != null)
                {
                    dataGrid.DataSource = dataTable;
                    Resources = (dataTable.AsEnumerable().Select(x => x["idK"].ToString()).Distinct().ToList());
                }
            }
        }

        private void button2_Click(object sender, EventArgs e)
        {

            string ambi = BtnTest.Checked == true ? "Test" : "Productivo";

            DialogResult dr = MessageBox.Show("¿Desea cargar la información en el ambiente " + ambi + "?",
                      "Confirmar", MessageBoxButtons.YesNo);
            switch (dr)
            {
                case DialogResult.Yes:

                    if (dataGrid.Rows.Count != 0)
                    {
                        actividadesCorrectas = 0;
                        actividadesFallidas = 0;
                        int dubedos = dataGrid.Rows.Count >= 5000 ? Convert.ToInt32(dataGrid.Rows.Count / 4) : dataGrid.Rows.Count;
                        Stopwatch stopwatch = new Stopwatch();
                        stopwatch.Start();
                        iterations = 0;
                        txtBody.Text += "Iniciando... \r\n";
                        ActivarRuta();
                        txtBody.Update();
                        txtBody.SelectionStart = txtBody.Text.Length;
                        txtBody.ScrollToCaret();
                        InsertActivities(dubedos);
                        stopwatch.Stop();
                        txtBody.Text += "Proceso finalizado... \r\n";
                        txtBody.Update();
                        txtBody.SelectionStart = txtBody.Text.Length;
                        txtBody.ScrollToCaret();
                        txtBody.Text += "Tiempo de proceso: " + Math.Round(stopwatch.Elapsed.TotalMinutes, 2) + " minutos, Total iteraciones: " + iterations + "\r\n";
                        txtBody.Text += "\r\nActividades cargadas correctamente: " + actividadesCorrectas + "\r\nActividades fallidas: " + actividadesFallidas + "\r\n";
                        txtBody.Update();
                        txtBody.SelectionStart = txtBody.Text.Length;
                        txtBody.ScrollToCaret();
                    }
                    break;
            }
        }

        private void InsertActivities(int dubedos)
        {
            try
            {
                if (dubedos != 0)
                {
                    txtBody.Text += "\r\nIteración " + iterations + " \r\n";
                    txtBody.Update();
                    txtBody.SelectionStart = txtBody.Text.Length;
                    txtBody.ScrollToCaret();
                    iterations++;
                    bool created = false;
                    int Max = dubedos;
                    PBar.Maximum = Max;
                    PBar.Value = 1;
                    PBar.Step = 1;
                    txtBody.Text += "Generando codigo JSON... \r\n";
                    txtBody.Update();
                    txtBody.SelectionStart = txtBody.Text.Length;
                    txtBody.ScrollToCaret();
                    string body = "{ \n";
                    body += "\"updateParameters\" : { \n ";
                    body += "\"identifyActivityBy\" : \"apptNumber\" , \n";
                    body += "\"ifInFinalStatusThen\" : \"doNothing\" \n ";
                    body += "}, \n ";
                    body += "\"activities\" : [ \n ";
                    int i = 0;
                    foreach (DataGridViewRow row in dataGrid.Rows)
                    {

                        //TipoActividad,id,idK,idSucursal,Prioridad,apellidoPaterno,apellidoMaterno,nombre,fechaNacimiento,sexo,Domicilio,lat,lng,telefonoFijo,telefonoCelular,capitalInicial,,interesesOrdinariosPendiente,interesesMoratoriosPendiente,iva,totalPendiente,numeroPagosPendientes,numeroPagosRealizados,montoPagoFrecuente,fechaUltimoPago,diasMoratorios,importeUltimoPago,fechaCreditoOtorgamiento,saldoExigible,date
                        //loader += "\"capitalAbonado\":\"" + row["capitalInicial"].ToString() + "\", \n";
                        //loader += "\"interesesOrdinariosAbonado\":\"" + row["capitalInicial"].ToString() + "\", \n";
                        //loader += "\"interesesMoratoriosAbonado\":\"" + row["capitalInicial"].ToString() + "\", \n";
                        //loader += "\"totalAbonado\":\"" + row["capitalInicial"].ToString() + "\", \n";
                        //body += "\"deliveryWindowStart\":\"" + row["capitalInicial"].ToString() + "\", \n";
                        //body += "\"deliveryWindowEnd\":\"" + row["capitalInicial"].ToString() + "\", \n";
                        //body += "\"startTime\":\"" + row["capitalInicial"].ToString() + "\", \n";
                        //body += "\"endTime\":\"" + row["capitalInicial"].ToString() + "\", \n";
                        //body += "\"city\":\"" + row["capitalInicial"].ToString() + "\", \n";
                        //body += "\"estado\":\"" + row["capitalInicial"].ToString() + "\", \n";
                        //body += "\"postalCode\":\"" + row["capitalInicial"].ToString() + "\" \n";
                        //body += "\"numeroInt\":\"" + row["capitalInicial"].ToString() + "\", \n";

                        i++;

                        string loader = "";
                        loader += "{ \n";
                        loader += "\"deliveryWindowStart\": \"16:45:00\", \n";
                        loader += "\"deliveryWindowEnd\": \"17:45:00\", \n";
                        loader += "\"startTime\": \"\", \n";
                        loader += "\"endTime\": \"\", \n";
                        loader += "\"city\": \".\", \n";
                        loader += "\"postalCode\": \".\", \n";
                        loader += "\"timeSlot\": \"Todo el d\u00eda\", \n";
                        loader += "\"resourceId\":\"" + row.Cells["idK"].Value.ToString() + "\", \n";
                        loader += "\"activityType\":\"" + row.Cells["tipoActividad"].Value.ToString() + "\", \n";
                        loader += "\"apptNumber\":\"" + row.Cells["id"].Value.ToString() + "\", \n";
                        loader += "\"capitalInicial\":\"" + row.Cells["capitalInicial"].Value.ToString() + "\", \n";
                        loader += "\"capitalPendiente\":\"" + row.Cells["capitalPendiente"].Value.ToString() + "\", \n";
                        loader += "\"interesesOrdinariosPendiente\":\"" + row.Cells["interesesOrdinariosPendiente"].Value.ToString() + "\", \n";
                        loader += "\"interesesMoratoriosPendiente\":\"" + row.Cells["interesesMoratoriosPendiente"].Value.ToString() + "\", \n";
                        loader += "\"iva\":\"" + row.Cells["iva"].Value.ToString() + "\", \n";
                        loader += "\"totalPendiente\":\"" + row.Cells["totalPendiente"].Value.ToString() + "\", \n";
                        loader += "\"numeroPagosPendientes\":\"" + row.Cells["numeroPagosPendientes"].Value.ToString() + "\", \n";
                        loader += "\"numeroPagosRealizados\":\"" + row.Cells["numeroPagosRealizados"].Value.ToString() + "\", \n";
                        loader += "\"montoPagoFrecuente\":\"" + row.Cells["montoPagoFrecuente"].Value.ToString() + "\", \n";
                        loader += "\"fechaUltimoPago\":\"" + row.Cells["fechaUltimoPago"].Value.ToString() + "\", \n";
                        loader += "\"diasMoratorios\":\"" + row.Cells["diasMoratorios"].Value.ToString() + "\", \n";
                        loader += "\"importeUltimoPago\":\"" + row.Cells["importeUltimoPago"].Value.ToString() + "\", \n";
                        loader += "\"saldoExigible\":\"" + row.Cells["saldoExigible"].Value.ToString() + "\", \n";
                        loader += "\"aPaterno\":\"" + row.Cells["apellidoPaterno"].Value.ToString() + "\", \n";
                        loader += "\"aMaterno\":\"" + row.Cells["apellidoMaterno"].Value.ToString() + "\", \n";
                        loader += "\"idSucursal\":\"" + row.Cells["idSucursal"].Value.ToString() + "\", \n";
                        loader += "\"fechaNacimiento\":\"" + row.Cells["fechaNacimiento"].Value.ToString() + "\", \n";
                        loader += "\"customerName\":\"" + row.Cells["nombre"].Value.ToString() + "\", \n";
                        // Cambio Esteban - Latitud / Longitud
                        if (!string.IsNullOrEmpty(row.Cells["lat"].Value.ToString()))
                        {
                            loader += "\"latitude\":" + row.Cells["lat"].Value + ", \n";
                        }
                        if (!string.IsNullOrEmpty(row.Cells["lng"].Value.ToString()))
                        {
                            loader += "\"longitude\":" + row.Cells["lng"].Value + ", \n";
                        }
                        // FIN Cambio
                        loader += "\"customerPhone\":\"" + row.Cells["telefonoFijo"].Value.ToString() + "\", \n";
                        loader += "\"customerCell\":\"" + row.Cells["telefonoCelular"].Value.ToString() + "\", \n";
                        loader += "\"Prioridad\":\"" + row.Cells["Prioridad"].Value.ToString() + "\", \n";
                        if (string.IsNullOrEmpty(row.Cells["date"].Value.ToString()))
                        {
                            loader += "\"streetAddress\":\"" + row.Cells["Domicilio"].Value.ToString() + "\" \n";
                        }
                        else
                        {
                            loader += "\"streetAddress\":\"" + row.Cells["Domicilio"].Value.ToString() + "\", \n";
                            loader += "\"date\":\"" + row.Cells["date"].Value.ToString() + "\" \n";
                        }
                        if (i == Max)
                        {
                            loader += "} \n ";
                            body += loader;
                            PBar.PerformStep();
                            break;
                        }
                        else
                        {
                            loader += "},\n ";
                            body += loader;
                            PBar.PerformStep();
                        }
                    }
                    body += "] \n";
                    body += "} \n";

                    txtBody.Text += "Realizando petición a WebServices... \r\n";
                    txtBody.Update();
                    txtBody.SelectionStart = txtBody.Text.Length;
                    txtBody.ScrollToCaret();

                    string urlcliente = BtnTest.Checked == true ? "https://fincomunfieldserv1.test.etadirect.com/" : "https://fincomunfieldserv.etadirect.com/";
                    string user = BtnTest.Checked == true ? "mid@fincomunfieldserv1.test" : "mid@fincomunfieldserv";
                    string pass = BtnTest.Checked == true ? "3a2aeb2ced436bcdddd8791bba5b279801076e04c51cf7243807d6563607e936" : "ce52a9a3eff4f55070760291f173a27e3fe80c91dc871ccadf13505c20d7f333";

                    var client = new RestClient(urlcliente);
                    client.Authenticator = new HttpBasicAuthenticator(user, pass);
                    var request = new RestRequest("rest/ofscCore/v1/activities/custom-actions/bulkUpdate", Method.POST)
                    {
                        RequestFormat = DataFormat.Json,
                    };
                    request.AddParameter("application/json", body, ParameterType.RequestBody);
                    request.Timeout = 12000000;
                    IRestResponse response = client.Execute(request);
                    var content = response.Content;

                    Results.RootObject resultados = JsonConvert.DeserializeObject<Results.RootObject>(content);
                    if (resultados != null & resultados.results.Count > 0)
                    {
                        RootObjResults = resultados;
                        foreach(Results.Result resultado in RootObjResults.results)
                        {
                            if (resultado.operationsFailed != null)
                            {
                                actividadesFallidas += 1;
                                txtBody.Text += "\r\nError al subir appt: \r\n";
                                txtBody.Text += "apptNumber: " + resultado.activityKeys.apptNumber.ToString() + "\r\n";
                                foreach (Results.Error error in resultado.errors)
                                {
                                    txtBody.Text += "error: " + error.errorDetail.ToString() + "\r\n";
                                }
                                txtBody.Update();
                                txtBody.SelectionStart = txtBody.Text.Length;
                                txtBody.ScrollToCaret();
                            }
                            else
                            {
                                actividadesCorrectas += 1;
                            }
                        }
                    }

                    if (!content.Contains("operationsFailed"))
                    {
                        txtBody.Text += "\r\nPetición correcta... \r\n";
                        txtBody.Update();
                        txtBody.SelectionStart = txtBody.Text.Length;
                        txtBody.ScrollToCaret();
                        created = true;
                    }
                    else
                    {
                        txtBody.Text += "\r\nPetición no correcta...  \r\n";
                        txtBody.Text += "Detalle: algunas actividades no fueron cargadas.  \r\n";// + response.Content;
                        txtBody.Update();
                        txtBody.SelectionStart = txtBody.Text.Length;
                        txtBody.ScrollToCaret();
                        created = true;
                        //MessageBox.Show("Error:");
                        //return;
                    }
                    /*
                    else
                    {
                        txtBody.Text += "Petición no correcta...  \r\n";
                        txtBody.Text += "Detalle: " + response.Content;
                        txtBody.Update();
                        //MessageBox.Show("Error:");
                        return;

                    }
                    */
                    if (created)
                    {
                        i = 0;
                        List<DataRow> rowsWantToDelete = new List<DataRow>();
                        foreach (DataRow row in dataTable.Rows)
                        {
                            i++;
                            rowsWantToDelete.Add(row);
                            if (i == Max)
                            {
                                break;
                            }
                        }

                        foreach (DataRow dr in rowsWantToDelete)
                        {
                            dataTable.Rows.Remove(dr);
                        }
                        dataTable.AcceptChanges();

                        dataGrid.DataSource = null;
                        dataGrid.DataSource = dataTable;
                        dataGrid.Refresh();
                        dataGrid.Update();
                        if (dataGrid.Rows.Count >= dubedos)
                        {
                            InsertActivities(dubedos);
                        }
                        else
                        {
                            InsertActivities(dataGrid.Rows.Count);
                        }
                    }
                }
                else
                {
                    return;
                }
            }
            catch (Exception ex)
            {
                txtBody.ReadOnly = false;
                txtBody.Clear();
                txtBody.Text = "Surgió un error: \r\n";
                txtBody.Text += "Mensaje de error: " + ex.Message + " \r\n";
                txtBody.Text += "Detalle: " + ex.StackTrace + " \r\n";
            }
        }

        private void ActivarRuta()
        {
            try
            {
                PBar.Maximum = Resources.Count;
                PBar.Value = 0;
                PBar.Step = 1;
                txtBody.Text += "Activando ruta para " + Resources.Count.ToString() + " recursos... \r\n";
                txtBody.Update();
                txtBody.SelectionStart = txtBody.Text.Length;
                txtBody.ScrollToCaret();
                foreach (string resource in Resources)
                {
                    string urlcliente = BtnTest.Checked == true ? "https://fincomunfieldserv1.test.etadirect.com/" : "https://fincomunfieldserv.etadirect.com/";
                    string user = BtnTest.Checked == true ? "mid@fincomunfieldserv1.test" : "mid@fincomunfieldserv";
                    string pass = BtnTest.Checked == true ? "3a2aeb2ced436bcdddd8791bba5b279801076e04c51cf7243807d6563607e936" : "ce52a9a3eff4f55070760291f173a27e3fe80c91dc871ccadf13505c20d7f333";

                    string dateRoute = DateTime.Now.ToString("yyyy-MM-dd");
                    var client = new RestClient(urlcliente);
                    client.Authenticator = new HttpBasicAuthenticator(user, pass);
                    var request = new RestRequest("rest/ofscCore/v1/resources/" + resource + "/routes/" + dateRoute + "/custom-actions/activate", Method.POST)
                    {
                        RequestFormat = DataFormat.Json,
                    };
                    IRestResponse response = client.Execute(request);
                    PBar.PerformStep();
                }
                txtBody.Text += "Rutas activas\r\n";
                txtBody.Update();
                txtBody.SelectionStart = txtBody.Text.Length;
                txtBody.ScrollToCaret();
            }
            catch (Exception ex)
            {
                txtBody.ReadOnly = false;
                txtBody.Clear();
                txtBody.Text = "Surgió un error: \r\n";
                txtBody.Text += "Mensaje de error: " + ex.Message + " \r\n";
                txtBody.Text += "Detalle: " + ex.StackTrace + " \r\n";
            }
        }

        private void BtnTest_CheckedChanged(object sender, EventArgs e)
        {
            if (BtnTest.Checked == true)
            {
                BtnProd.Checked = false;
            }
        }

        private void BtnProd_CheckedChanged(object sender, EventArgs e)
        {
            if (BtnProd.Checked == true)
            {
                BtnTest.Checked = false;
            }
        }
    }
}


