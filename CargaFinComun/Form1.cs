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

namespace CargaFinComun
{
    public partial class Form1 : Form
    {

        public DataTable dataTable { get; set; }
        public int iterations { get; set; }
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
                lblRecords.Text = filelines.Count().ToString();
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
                }
            }
        }

        private void button2_Click(object sender, EventArgs e)
        {
            if (dataGrid.Rows.Count != 0)
            {
                int dubedos = dataGrid.Rows.Count >= 5000 ? Convert.ToInt32(dataGrid.Rows.Count / 4) : dataGrid.Rows.Count;
                Stopwatch stopwatch = new Stopwatch();
                stopwatch.Start();
                iterations = 0;
                txtBody.Text += "Iniciando... \r\n";
                txtBody.Update();
                InsertActivities(dubedos);
                stopwatch.Stop();
                txtBody.Text += "Proceso finalizado... \r\n";
                txtBody.Update();
                txtBody.Text += "Tiempo de proceso: " + Math.Round(stopwatch.Elapsed.TotalMinutes, 2) + " minutos, Total iteraciones: " + iterations + "\r\n";
                txtBody.Update();
            }
        }

        private void InsertActivities(int dubedos)
        {
            if (dubedos != 0)
            {
                txtBody.Text += "Iteración " + iterations + " \r\n";
                txtBody.Update();
                iterations++;
                bool created = false;
                int Max = dubedos;
                PBar.Maximum = Max;
                PBar.Value = 1;
                PBar.Step = 1;
                txtBody.Text += "Generando codigo JSON... \r\n";
                txtBody.Update();
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
                    loader += "\"timeSlot\": \"Todo el d\u00eda\", \n";
                    loader += "\"resourceId\":\"" + row.Cells["idK"].Value.ToString() + "\", \n";
                    loader += "\"activityType\":\"" + row.Cells["tipoActividad"].Value.ToString() + "\", \n";
                    loader += "\"apptNumber\":\"" + row.Cells["id"].Value.ToString() + DateTime.Now.ToString("mmss") + "\", \n";
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
                    loader += "\"customerPhone\":\"" + row.Cells["telefonoCelular"].Value.ToString() + "\", \n";
                    loader += "\"customerNumber\":\"" + row.Cells["telefonoFijo"].Value.ToString() + "\", \n";
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
                var client = new RestClient("https://fincomunfieldserv1.test.etadirect.com/");
                client.Authenticator = new HttpBasicAuthenticator("mid@fincomunfieldserv1.test", "3a2aeb2ced436bcdddd8791bba5b279801076e04c51cf7243807d6563607e936");
                var request = new RestRequest("rest/ofscCore/v1/activities/custom-actions/bulkUpdate", Method.POST)
                {
                    RequestFormat = DataFormat.Json,
                };
                request.AddParameter("application/json", body, ParameterType.RequestBody);
                request.Timeout = 12000000;
                IRestResponse response = client.Execute(request);

                var content = response.Content;
                if (response.StatusCode == HttpStatusCode.OK)
                {
                    txtBody.Text += "Petición correcta... \r\n";
                    txtBody.Update();
                    created = true;
                }
                else
                {

                    txtBody.Text += "Petición no correcta...  \r\n";
                    txtBody.Update();
                    //MessageBox.Show("Error:");
                    return;

                }

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
    }
}


