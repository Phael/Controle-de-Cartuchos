using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using System.Data;
using System.Data.OleDb;

namespace Controle_de_Cartuchos
{
    public partial class Form_Principal : Form
    {
        private OleDbConnection Conn;
        private OleDbDataAdapter Da;
        private DataSet Ds;

        public Form_Principal()
        {
            InitializeComponent();
        }
        public void IniciaAcesso()
        {
            Conn = new OleDbConnection("Provider=Microsoft.Jet.OLEDB.4.0;Data Source=D:/Repositorio/Controle-de-Cartuchos/Banco/Cartuchos.mdb");
            Ds = new DataSet();

            try
            {
                Conn.Open();
            }
            catch (System.Exception e)
            {
                MessageBox.Show(e.Message.ToString());
            }

            if (Conn.State == ConnectionState.Open)
            {
                Da = new OleDbDataAdapter("SELECT Código,Nome,Data  from Cartuchos", Conn);
                Da.Fill(Ds, "Cartuchos");
                dataGridView_Cartuchos.DataSource = Ds;
                dataGridView_Cartuchos.DataMember = "Cartuchos";
            }
        }

        private void Form1_Load(object sender, EventArgs e)
        {
            IniciaAcesso();
        }
        
        private void button_Processar_Click(object sender, EventArgs e)
        {
            string Conexao = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=D:/Repositorio/Controle-de-Cartuchos/Banco/Cartuchos.mdb";
            string Insert = "INSERT INTO Clientes(Nome, Telefone, Produto1, Produto2, Produto3, Produto4, Produto5, Produto6, Produto7, Produto8, Servico1 ,Servico2, Servico3, Servico4, Servico5, Servico6, Servico7, Servico8, Identificacao1, Identificacao2, Identificacao3, Identificacao4, Identificacao5, Identificacao6, Identificacao7, Identificacao8, PSEntrada1, PSEntrada2, PSEntrada3, PSEntrada4, PSEntrada5, PSEntrada6, PSEntrada7, PSEntrada8, PSSaida1, PSSaida2, PSSaida3, PSSaida4, PSSaida5, PSSaida6, PSSaida7, PSSaida8, Resultado1, Resultado2, Resultado3, Resultado4, Resultado5, Resultado6, Resultado7, Resultado8, Valor1, Valor2, Valor3, Valor4, Valor5, Valor6, Valor7, Valor8, Baia1, Baia2, Baia3, Baia4, Baia5, Baia6, Baia7, Baia8, Observacao, Data)" + " VALUES('" + textBox_Nome.Text + "','" + textBox_Telefone.Text + "','" + comboBox_Produto1.Text + "','" + comboBox_Servico1.Text + "','" + textBox_Identificacao.Text + "','" + textBox_PsEntrada1.Text + "','" + textBox_PsSaida1.Text + "','" + textBox_Resultado1.Text + "','" + textBox_Baia1.Text + "')";

            //cria a conexão com o banco de dados
            OleDbConnection dbConnection = new OleDbConnection(Conexao);

            //Cria o comando que inicia a query
            OleDbCommand cmdQry = new OleDbCommand(Insert, dbConnection);
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
