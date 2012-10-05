using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using System.Data.OleDb;

namespace Controle_de_Cartuchos
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private void textBox1_TextChanged(object sender, EventArgs e)
        {

        }

        private void label1_Click(object sender, EventArgs e)
        {

        }

        private void comboBox2_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (comboBox_Produto1.Text == "Cartucho" && comboBox_Servico1.Text == "Recarga")
                textBox_Valor1.Text = "15,00";
            
        }

        private void cartuchosBindingNavigatorSaveItem_Click(object sender, EventArgs e)
        {
        }

        private void Form1_Load(object sender, EventArgs e)
        {

        }

        private void cartuchosDataGridView_CellContentClick(object sender, DataGridViewCellEventArgs e)
        {

        }

        private void button_Processar_Click(object sender, EventArgs e)
        {
            string strConnection = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=I:\\PJS\\Controle-de-Cartuchos\\Cartuchos.mdb";
            string strSQL = "INSERT INTO Clientes(nome,endereco,cidade,estado,cep,telefone)"
                                    + " VALUES ('" + textBox_Valor1.Text + "'";
 
            //cria a conexão com o banco de dados
            OleDbConnection dbConnection = new OleDbConnection(strConnection);
 
            //Cria o comando que inicia a query
            OleDbCommand cmdQry = new OleDbCommand(strSQL, dbConnection);
            try
            {
            // abre o banco
            dbConnection.Open();
            // executa a query
            cmdQry.ExecuteNonQuery();
            //
            MessageBox.Show("Dados Salvos com sucesso.");
            }
            //Trata a exceção
            catch (OleDbException ex)
            {
            MessageBox.Show("Error: " + ex.Message);
            }
            finally
            {
            //fecha a conexao
            dbConnection.Close();
            }
        }
    }
}
