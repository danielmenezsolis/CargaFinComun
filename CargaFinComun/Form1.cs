using CsvHelper;
using RestSharp;
using RestSharp.Authenticators;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.IO;
using System.Linq;
using System.Net;
using System.Text;
using System.Windows.Forms;

namespace CargaFinComun
{
    public partial class Form1 : Form
    {
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
                string[] filelines = File.ReadAllLines(filename);
                lblFile.Text = filename;
                lblRecords.Text = filelines.Count().ToString();
            }

        }

        private void button2_Click(object sender, EventArgs e)
        {

            string[] lines = File.ReadAllLines(lblFile.Text, Encoding.UTF8);
            CsvReader csv = new CsvReader(File.OpenText(lblFile.Text));
            csv.Read();
            csv.ReadHeader();

            List<string> headers = csv.Context.HeaderRecord.ToList();

            // Read csv into datatable
            DataTable dataTable = new DataTable();

            foreach (string header in headers)
            {
                dataTable.Columns.Add(new System.Data.DataColumn(header));
            }

            // Check all required columns are present
            if (!headers.Exists(x => x == "idSucursal"))
            {
                throw new ArgumentException("idSucursal field not present in input file.");
            }
            else if (!headers.Exists(x => x == "Prioridad"))
            {
                throw new ArgumentException("Prioridad field not present in input file.");
            }
            else if (!headers.Exists(x => x == "nombre"))
            {
                throw new ArgumentException("nombre field not present in input file.");
            }

            while (csv.Read())
            {
                System.Data.DataRow row = dataTable.NewRow();

                foreach (System.Data.DataColumn column in dataTable.Columns)
                {
                    row[column.ColumnName] = csv.GetField(column.DataType, column.ColumnName);
                }

                dataTable.Rows.Add(row);
            }

            if (dataTable != null)
            {
                dataGrid.DataSource = dataTable;
                if (dataTable.Rows.Count != 0)
                {
                    InsertActivities(dataTable);
                }
            }


        }

        private void InsertActivities(DataTable dataTable)
        {

            PBar.Minimum = 1;
            PBar.Maximum = 1000;
            PBar.Value = 1;
            // Set the Step property to a value of 1 to represent each file being copied.
            PBar.Step = 1;
            string body = "{ \n";
            body += "\"updateParameters\" : { \n ";

            body += "\"identifyActivityBy\" : \"apptNumber\" , \n";
            body += "\"ifInFinalStatusThen\" : \"doNothing\" \n ";
            body += "}, \n ";
            body += "\"activities\" : [ \n ";

            int i = 0;
            foreach (DataRow row in dataTable.Rows)
            {
                i++;

                body += "{ \n";
                body += "\"timeSlot\": \"Todo el d\u00eda \", \n";
                body += "\"resourceId\":34, \n";

                //TipoActividad,id,idK,idSucursal,Prioridad,apellidoPaterno,apellidoMaterno,nombre,fechaNacimiento,sexo,Domicilio,lat,lng,telefonoFijo,telefonoCelular,capitalInicial,,interesesOrdinariosPendiente,interesesMoratoriosPendiente,iva,totalPendiente,numeroPagosPendientes,numeroPagosRealizados,montoPagoFrecuente,fechaUltimoPago,diasMoratorios,importeUltimoPago,fechaCreditoOtorgamiento,saldoExigible,date
                body += "\"activityType\":\"" + row["TipoActividad"].ToString() + "\", \n";
                body += "\"apptNumber\":\"" + row["id"].ToString() + "\", \n";
                body += "\"capitalInicial\":\"" + row["capitalInicial"].ToString() + "\", \n";
                body += "\"capitalPendiente\":\"" + row["capitalPendiente"].ToString() + "\", \n";
                body += "\"interesesOrdinariosPendiente\":\"" + row["interesesOrdinariosPendiente"].ToString() + "\", \n";
                body += "\"interesesMoratoriosPendiente\":\"" + row["interesesMoratoriosPendiente"].ToString() + "\", \n";
                body += "\"iva\":\"" + row["iva"].ToString() + "\", \n";
                body += "\"totalPendiente\":\"" + row["totalPendiente"].ToString() + "\", \n";
                body += "\"capitalAbonado\":\"" + row["capitalInicial"].ToString() + "\", \n";
                body += "\"interesesOrdinariosAbonado\":\"" + row["capitalInicial"].ToString() + "\", \n";
                body += "\"interesesMoratoriosAbonado\":\"" + row["capitalInicial"].ToString() + "\", \n";
                body += "\"totalAbonado\":\"" + row["capitalInicial"].ToString() + "\", \n";
                body += "\"numeroPagosPendientes\":\"" + row["numeroPagosPendientes"].ToString() + "\", \n";
                body += "\"numeroPagosRealizados\":\"" + row["numeroPagosRealizados"].ToString() + "\", \n";
                body += "\"montoPagoFrecuente\":\"" + row["montoPagoFrecuente"].ToString() + "\", \n";
                body += "\"fechaUltimoPago\":\"" + row["fechaUltimoPago"].ToString() + "\", \n";
                body += "\"diasMoratorios\":\"" + row["diasMoratorios"].ToString() + "\", \n";
                body += "\"importeUltimoPago\":\"" + row["importeUltimoPago"].ToString() + "\", \n";
                body += "\"importeUltimoPago\":\"" + row["capitalInicial"].ToString() + "\", \n";
                body += "\"saldoExigible\":\"" + row["saldoExigible"].ToString() + "\", \n";
                body += "\"aPaterno\":\"" + row["apellidoPaterno"].ToString() + "\", \n";
                body += "\"aMaterno\":\"" + row["apellidoMaterno"].ToString() + "\", \n";
                body += "\"idSucursal\":\"" + row["idSucursal"].ToString() + "\", \n";
                body += "\"fechaNacimiento\":\"" + row["fechaNacimiento"].ToString() + "\", \n";
                body += "\"customerName\":\"" + row["nombre"].ToString() + "\", \n";
                body += "\"customerPhone\":\"" + row["telefonoCelular"].ToString() + "\", \n";
                body += "\"customerNumber\":\"" + row["telefonoFijo"].ToString() + "\", \n";
                body += "\"Prioridad\":\"" + row["Prioridad"].ToString() + "\", \n";
                body += "\"deliveryWindowStart\":\"" + row["capitalInicial"].ToString() + "\", \n";
                body += "\"deliveryWindowEnd\":\"" + row["capitalInicial"].ToString() + "\", \n";
                body += "\"startTime\":\"" + row["capitalInicial"].ToString() + "\", \n";
                body += "\"endTime\":\"" + row["capitalInicial"].ToString() + "\", \n";
                body += "\"streetAddress\":\"" + row["Domicilio"].ToString() + "\", \n";
                body += "\"numeroInt\":\"" + row["capitalInicial"].ToString() + "\", \n";
                body += "\"date\":\"" + row["date"].ToString() + "\", \n";
                body += "\"city\":\"" + row["capitalInicial"].ToString() + "\", \n";
                body += "\"estado\":\"" + row["capitalInicial"].ToString() + "\", \n";
                body += "\"postalCode\":\"" + row["capitalInicial"].ToString() + "\" \n";

                if (i == 1000)
                {
                    body += "} \n ";
                    break;
                }
                else
                {
                    body += "},\n ";
                }
                PBar.PerformStep();
            }
            body += "] \n";
            body += "} \n";
            txtBody.Text = body;
            dataGrid.DataSource = null;
            dataGrid.DataSource = dataTable;


            /*
                     var client = new RestClient("https://fincomunfieldserv1.test.etadirect.com/");

                     client.Authenticator = new HttpBasicAuthenticator("mid@fincomunfieldserv1.test", "3a2aeb2ced436bcdddd8791bba5b279801076e04c51cf7243807d6563607e936");
                     var request = new RestRequest("rest/ofscCore/v1/activities/custom-actions/bulkUpdate", Method.POST)
                     {
                         RequestFormat = DataFormat.Json,
                     };
                     request.AddParameter("application/json", body, ParameterType.RequestBody);
                     IRestResponse response = client.Execute(request);
                     var content = response.Content;
                     if (response.StatusCode == HttpStatusCode.Created)
                     {
                     }
                     else
                     {
                         MessageBox.Show("Error" + content);
                     }
                     */
        }
    }

}


