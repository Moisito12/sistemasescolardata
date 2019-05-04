using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using MySql.Data.MySqlClient;
using System.IO;
using ClosedXML.Excel;
using iTextSharp.text.pdf;
using iTextSharp.text;

namespace sistemasescolaresdata
{

    public partial class Form1 : Form
    {
        MySqlConnection mySqlConnection;
        MySqlCommand mySqlCommand;
        public Form1()
        {
            InitializeComponent();
        }

        private void Form1_Load(object sender, EventArgs e)
        {
            cargadatos();
        }
        private void cargadatos()
        {

            try
            {
                MySqlConnection mySqlConnection = new MySqlConnection("host=localhost;user=root;password=moisito12;database=sistemas_escolar");
                mySqlConnection.Open();
                //Aquí se empieza con la ejecución del adaptadater
                MySqlDataAdapter mySqlDataAdapter =
                new MySqlDataAdapter("SELECT matricula as 'Expediente',UPPER(ap1)" + " as 'Apellido Paterno',UPPER(ap2)" + " as 'Apellido Materno'," +
                "UPPER(nombre)" + " as 'Nombre',Fnacimiento as 'Fecha de Nacimiento',Correo as 'E-Mail', TELEFONO as 'Telefóno' FROM alumnos", mySqlConnection);
                DataTable dataTable = new DataTable();
                mySqlDataAdapter.Fill(dataTable);
                dataGridView1.DataSource = dataTable;
            }
            catch (Exception err)
            {

                MessageBox.Show(err.ToString());
            }

        }

        private void button1_Click(object sender, EventArgs e)
        {
            String query = "INSERT INTO alumnos (matricula,ap1,ap2,nombre,Fnacimiento,Correo,TELEFONO) VALUES (" + textBox1.Text + ",'" + textBox2.Text + "','" + textBox3.Text + "','" +
                textBox4.Text + "','" + dateTimePicker1.Value.Year + "-" + dateTimePicker1.Value.Month + "-" + dateTimePicker1.Value.Day + "', '"
                + textBox6.Text + "','" + textBox7.Text + "')";
            try
            {
                mySqlConnection = new MySqlConnection("host=localhost;user=root;password=moisito12;database=sistemas_escolar");
                mySqlConnection.Open();
                mySqlCommand = new MySqlCommand(query, mySqlConnection);
                mySqlCommand.ExecuteNonQuery();
                MessageBox.Show("Agregando Alumno", "Alumno Ingresado Exitosa");
                cargadatos();
                mySqlConnection.Close();
            }
            catch (Exception err)
            {

                MessageBox.Show(err.ToString(), "Titulo", MessageBoxButtons.RetryCancel, MessageBoxIcon.Exclamation);
            }
        }

        private void button2_Click(object sender, EventArgs e)
        {
            String query = "DELETE FROM alumnos WHERE matricula=" + textBox1.Text;
            try
            {
                mySqlConnection = new MySqlConnection("host=localhost;user=root;password=moisito12;database=sistemas_escolar");
                mySqlConnection.Open();
                mySqlCommand = new MySqlCommand(query, mySqlConnection);
                mySqlCommand.ExecuteNonQuery();
                MessageBox.Show("Eliminando Alumno", "Borrado Realizado");
                cargadatos();
                mySqlConnection.Close();
            }
            catch (Exception err)
            {
                MessageBox.Show(err.ToString(), "Título", MessageBoxButtons.RetryCancel, MessageBoxIcon.Hand);
            }

        }

        private void dataGridView1_CellClick(object sender, DataGridViewCellEventArgs e)
        {
            textBox1.Text = dataGridView1[0, e.RowIndex].Value.ToString();
            textBox2.Text = dataGridView1[1, e.RowIndex].Value.ToString();
            textBox3.Text = dataGridView1[2, e.RowIndex].Value.ToString();
            textBox4.Text = dataGridView1[3, e.RowIndex].Value.ToString();
            dateTimePicker1.Value = DateTime.Parse(dataGridView1[4, e.RowIndex].Value.ToString());
            textBox6.Text = dataGridView1[5, e.RowIndex].Value.ToString();
            textBox7.Text = dataGridView1[6, e.RowIndex].Value.ToString();
        }

        private void button3_Click(object sender, EventArgs e)
        {
            saveFileDialog1.Title = "guardado1";
            saveFileDialog1.FileName = "*.sql";
            saveFileDialog1.InitialDirectory = @"C:\Escritorio";
            saveFileDialog1.Filter = "archivo sql |*.sql";
            saveFileDialog1.ShowDialog();
            string archivo;
            archivo = saveFileDialog1.FileName;
            MessageBox.Show(archivo);
            StreamWriter writer = new StreamWriter(archivo);
            for (int i = 0; i < dataGridView1.Rows.Count; i++)
            {
                writer.WriteLine("INSERT INTO alumnos (matricula,ap1,ap2,nombre,Fnacimiento,Correo,TELEFONO) VALUES (" +
                dataGridView1[0, i].Value.ToString() + ",'" +
                dataGridView1[1, i].Value.ToString() + "','" +
                dataGridView1[2, i].Value.ToString() + "','" +
                dataGridView1[3, i].Value.ToString() + "','" +
                Convert.ToDateTime(dataGridView1[4, i].Value.ToString()).Year + "-" +
                Convert.ToDateTime(dataGridView1[4, i].Value.ToString()).Month + "-" +
                Convert.ToDateTime(dataGridView1[4, i].Value.ToString()).Day + "','" +
                dataGridView1[5, i].Value.ToString() + "','" +
                dataGridView1[6, i].Value.ToString() + "');");
            }
            writer.Close();
        }

        private void button5_Click(object sender, EventArgs e)
        {
            saveFileDialog1.Title = "guardado2";
            saveFileDialog1.FileName = "*.csv";
            saveFileDialog1.InitialDirectory = @"C:\Escritorio";
            saveFileDialog1.Filter = "archivo csv |*.csv";
            saveFileDialog1.ShowDialog();
            string archivo;
            archivo = saveFileDialog1.FileName;
            MessageBox.Show(archivo);
            StreamWriter writer = new StreamWriter(archivo);
            for (int i = 0; i < dataGridView1.Columns.Count; i++)
            {
                if ((+i + 1) == dataGridView1.ColumnCount)
                {
                    writer.Write(dataGridView1.Columns[i].HeaderText + '\n');
                }
                else
                {
                    writer.Write(dataGridView1.Columns[i].HeaderText + ";");
                }
            }

            for (int i = 0; i < dataGridView1.Rows.Count; i++)
            {
                writer.WriteLine(dataGridView1[0, i].Value.ToString() + ";" +
                dataGridView1[1, i].Value.ToString() + ";" +
                dataGridView1[2, i].Value.ToString() + ";" +
                dataGridView1[3, i].Value.ToString() + ";" +
                Convert.ToDateTime(dataGridView1[4, i].Value.ToString()).Year + "" +
                Convert.ToDateTime(dataGridView1[4, i].Value.ToString()).Month + "-" +
                Convert.ToDateTime(dataGridView1[4, i].Value.ToString()).Day + ";" +
                dataGridView1[5, i].Value.ToString() + ";" +
                dataGridView1[6, i].Value.ToString());
            }
            writer.Close();
        }

        private void button9_Click(object sender, EventArgs e)
        {
            saveFileDialog1.Title = "guardado3";
            saveFileDialog1.FileName = "*.json";
            saveFileDialog1.InitialDirectory = @"C:\Escritorio";
            saveFileDialog1.Filter = "archivo json |*.json";
            saveFileDialog1.ShowDialog();
            string archivo;
            archivo = saveFileDialog1.FileName;
            MessageBox.Show(archivo);
            StreamWriter writer = new StreamWriter(archivo);
            writer.WriteLine("{\"sistemas_escolar\" :");
            writer.WriteLine("\t\t\talumnos : [");
            for (int i = 0; i < dataGridView1.Rows.Count; i++)
            {
                writer.WriteLine("\t\t{");
                writer.WriteLine(("\t\t \"Matricula\":") + "\"" + dataGridView1[0, i].Value.ToString() + "\"," + ",");
                writer.WriteLine(("\t\t \"Apellido Paterno\":") + "\"" + dataGridView1[1, i].Value.ToString() + "\"" + ",");
                writer.WriteLine(("\t\t \"Apellido Materno\":") + "\"" + dataGridView1[2, i].Value.ToString() + "\"" + ",");
                writer.WriteLine(("\t\t \"Nombre\":") + "\"" + dataGridView1[3, i].Value.ToString() + "\"" + ",");
                writer.WriteLine(("\t\t \"Fecha de Nacimiento\":") + "\"" + Convert.ToDateTime(dataGridView1[4, i].Value.ToString()).Year + "-" +
                Convert.ToDateTime(dataGridView1[4, i].Value.ToString()).Month + "-" + Convert.ToDateTime(dataGridView1[4, i].Value.ToString()).Day + "\"" + ",");
                writer.WriteLine(("\t\t \"Correo\":") + dataGridView1[5, i].Value.ToString() + ",");
                writer.WriteLine(("\t\t \"Telefóno\":") + dataGridView1[6, i].Value.ToString() + "\n}");
            }
            writer.WriteLine("\t\t\t\t ]");
            writer.WriteLine("\t}");
            writer.Close();

        }

        private void button4_Click(object sender, EventArgs e)
        {
            saveFileDialog1.Title = "guardado4";
            saveFileDialog1.FileName = "*.xlsx";
            saveFileDialog1.InitialDirectory = @"C:\Escritorio";
            saveFileDialog1.Filter = "archivo xlsx |*.xlsx";
            saveFileDialog1.ShowDialog();

            string archivo;
            archivo = saveFileDialog1.FileName;
            MessageBox.Show(archivo);

            var wordbook = new XLWorkbook();
            var hoja = wordbook.Worksheets.Add("Alumnos");
            hoja.Cell(1, 1).Value = "Matricula";
            hoja.Cell(1, 1).Style.Font.Bold = true;
            hoja.Cell(1, 2).Value = "Apellido P.";
            hoja.Cell(1, 2).Style.Font.Bold = true;
            hoja.Cell(1, 3).Value = "Apellido M.";
            hoja.Cell(1, 3).Style.Font.Bold = true;
            hoja.Cell(1, 4).Value = "Nombre";
            hoja.Cell(1, 4).Style.Font.Bold = true;
            hoja.Cell(1, 5).Value = "Fecha de Nacimiento";
            hoja.Cell(1, 5).Style.Font.Bold = true;
            hoja.Cell(1, 6).Value = "E-mail";
            hoja.Cell(1, 6).Style.Font.Bold = true;
            hoja.Cell(1, 7).Value = "Teléfono";
            hoja.Cell(1, 7).Style.Font.Bold = true;

            for (int i = 0; i < dataGridView1.Rows.Count; i++)
            {
                for (int k = 0; k < dataGridView1.Columns.Count; k++)
                    hoja.Cell((i + 2), (k + 1)).Value = dataGridView1.Rows[i].Cells[k].Value.ToString();
            }
            wordbook.SaveAs(archivo);
        }

        private void button7_Click(object sender, EventArgs e)
        {
            try
            {

                saveFileDialog1.Title = "guardado5";
                saveFileDialog1.FileName = "prueba.pdf";
                saveFileDialog1.InitialDirectory = @"C:\Escritorio";
                saveFileDialog1.Filter = "archivo pdf |*.pdf";
                saveFileDialog1.ShowDialog();

                string archivo;
                archivo = saveFileDialog1.FileName;
                MessageBox.Show(archivo);

                Document pdf = new Document(iTextSharp.text.PageSize.LETTER.Rotate());
                PdfWriter.GetInstance(pdf, new FileStream(saveFileDialog1.FileName, FileMode.Create));
                pdf.Open();
                PdfPTable tablepdf = new PdfPTable(7);
                PdfPCell titulo = new PdfPCell(new Phrase("Información de alumnos"));
                titulo.Colspan = 7;
                tablepdf.AddCell(titulo);
                tablepdf.AddCell("Matricula");
                tablepdf.AddCell("Apellido Paterno");
                tablepdf.AddCell("Apellido Materno");
                tablepdf.AddCell("Nombre");
                tablepdf.AddCell("Fecha de nacimiento");
                tablepdf.AddCell("Correo Electrónico");
                tablepdf.AddCell("Telefóno");

                for (int i = 0; i < dataGridView1.Rows.Count; i++)
                {
                    tablepdf.AddCell(dataGridView1[0, i].Value.ToString());
                    tablepdf.AddCell(dataGridView1[1, i].Value.ToString());
                    tablepdf.AddCell(dataGridView1[2, i].Value.ToString());
                    tablepdf.AddCell(dataGridView1[3, i].Value.ToString());
                    tablepdf.AddCell(dataGridView1[4, i].Value.ToString());
                    tablepdf.AddCell(dataGridView1[5, i].Value.ToString());
                    tablepdf.AddCell(dataGridView1[6, i].Value.ToString());
                }
                pdf.Add(tablepdf);
                pdf.Close();
            }
            catch (Exception)
            {

                throw;
            }
        }
    }
}
