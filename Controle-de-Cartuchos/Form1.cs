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
                Da = new OleDbDataAdapter("SELECT OS,Nome,Data FROM Cartuchos", Conn);
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
            string Nome = textBox_Nome.Text;
            string Telefone = textBox_Telefone.Text;

            string Produto1 = comboBox_Produto1.Text;
            string Produto2 = comboBox_Produto2.Text;
            string Produto3 = comboBox_Produto3.Text;
            string Produto4 = comboBox_Produto4.Text;
            string Produto5 = comboBox_Produto5.Text;
            string Produto6 = comboBox_Produto6.Text;
            string Produto7 = comboBox_Produto7.Text;
            string Produto8 = comboBox_Produto8.Text;

            string Servico1 = comboBox_Servico1.Text;
            string Servico2 = comboBox_Servico2.Text;
            string Servico3 = comboBox_Servico3.Text;
            string Servico4 = comboBox_Servico4.Text;
            string Servico5 = comboBox_Servico5.Text;
            string Servico6 = comboBox_Servico6.Text;
            string Servico7 = comboBox_Servico7.Text;
            string Servico8 = comboBox_Servico8.Text;

            string Identificacao1 = textBox_Identificacao1.Text;
            string Identificacao2 = textBox_Identificacao2.Text;
            string Identificacao3 = textBox_Identificacao3.Text;
            string Identificacao4 = textBox_Identificacao4.Text;
            string Identificacao5 = textBox_Identificacao5.Text;
            string Identificacao6 = textBox_Identificacao6.Text;
            string Identificacao7 = textBox_Identificacao7.Text;
            string Identificacao8 = textBox_Identificacao8.Text;

            string PSEntrada1 = textBox_PsEntrada1.Text;
            string PSEntrada2 = textBox_PsEntrada2.Text;
            string PSEntrada3 = textBox_PsEntrada3.Text;
            string PSEntrada4 = textBox_PsEntrada4.Text;
            string PSEntrada5 = textBox_PsEntrada5.Text;
            string PSEntrada6 = textBox_PsEntrada6.Text;
            string PSEntrada7 = textBox_PsEntrada7.Text;
            string PSEntrada8 = textBox_PsEntrada8.Text;

            string PSSaida1 = textBox_PsSaida1.Text;
            string PSSaida2 = textBox_PsSaida2.Text;
            string PSSaida3 = textBox_PsSaida3.Text;
            string PSSaida4 = textBox_PsSaida4.Text;
            string PSSaida5 = textBox_PsSaida5.Text;
            string PSSaida6 = textBox_PsSaida6.Text;
            string PSSaida7 = textBox_PsSaida7.Text;
            string PSSaida8 = textBox_PsSaida8.Text;

            string Resultado1 = textBox_Resultado1.Text;
            string Resultado2 = textBox_Resultado2.Text;
            string Resultado3 = textBox_Resultado3.Text;
            string Resultado4 = textBox_Resultado4.Text;
            string Resultado5 = textBox_Resultado5.Text;
            string Resultado6 = textBox_Resultado6.Text;
            string Resultado7 = textBox_Resultado7.Text;
            string Resultado8 = textBox_Resultado8.Text;

            float Valor1 = float.Parse(textBox_Valor1.Text);
            float Valor2 = float.Parse(textBox_Valor2.Text);
            float Valor3 = float.Parse(textBox_Valor3.Text);
            float Valor4 = float.Parse(textBox_Valor4.Text);
            float Valor5 = float.Parse(textBox_Valor5.Text);
            float Valor6 = float.Parse(textBox_Valor6.Text);
            float Valor7 = float.Parse(textBox_Valor7.Text);
            float Valor8 = float.Parse(textBox_Valor8.Text);

            string Baia1 = textBox_Baia1.Text;
            string Baia2 = textBox_Baia2.Text;
            string Baia3 = textBox_Baia3.Text;
            string Baia4 = textBox_Baia4.Text;
            string Baia5 = textBox_Baia5.Text;
            string Baia6 = textBox_Baia6.Text;
            string Baia7 = textBox_Baia7.Text;
            string Baia8 = textBox_Baia8.Text;

            string Observacao = textBox_Observacao.Text;

            string Data = dateTimePicker_Data.Value.ToShortDateString();

            string Conexao = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=D:/Repositorio/Controle-de-Cartuchos/Banco/Cartuchos.mdb";
            string Insert = "INSERT INTO Cartuchos(Nome, Telefone, Produto1, Produto2, Produto3, Produto4, Produto5, Produto6, Produto7, Produto8, Servico1 ,Servico2, Servico3, Servico4, Servico5, Servico6, Servico7, Servico8, Identificacao1, Identificacao2, Identificacao3, Identificacao4, Identificacao5, Identificacao6, Identificacao7, Identificacao8, PSEntrada1, PSEntrada2, PSEntrada3, PSEntrada4, PSEntrada5, PSEntrada6, PSEntrada7, PSEntrada8, PSSaida1, PSSaida2, PSSaida3, PSSaida4, PSSaida5, PSSaida6, PSSaida7, PSSaida8, Resultado1, Resultado2, Resultado3, Resultado4, Resultado5, Resultado6, Resultado7, Resultado8, Valor1, Valor2, Valor3, Valor4, Valor5, Valor6, Valor7, Valor8, Baia1, Baia2, Baia3, Baia4, Baia5, Baia6, Baia7, Baia8, Observacao, Data)" + " VALUES('" + Nome + "' , '" + Telefone + "','" + Produto1 + "','" + Produto2 + "','" + Produto3 + "','" + Produto4 + "','" + Produto5 + "','" + Produto6 + "','" + Produto7 + "','" + Produto8 + "','" + Servico1 + "','" + Servico2 + "','" + Servico3 + "','" + Servico4 + "','" + Servico5 + "','" + Servico6 + "','" + Servico7 + "','" + Servico8 + "','" + Identificacao1 + "','" + Identificacao2 + "','" + Identificacao3 + "','" + Identificacao4 + "','" + Identificacao5 + "','" + Identificacao6 + "','" + Identificacao7 + "','" + Identificacao8 + "','" + PSEntrada1 + "','" + PSEntrada2 + "','" + PSEntrada3 + "','" + PSEntrada4 + "','" + PSEntrada5 + "','" + PSEntrada6 + "','" + PSEntrada7 + "','" + PSEntrada8 + "','" + PSSaida1 + "','" + PSSaida2 + "','" + PSSaida3 + "','" + PSSaida4 + "','" + PSSaida5 + "','" + PSSaida6 + "','" + PSSaida7 + "','" + PSSaida8 + "','" + Resultado1 + "','" + Resultado2 + "','" + Resultado3 + "','" + Resultado4 + "','" + Resultado5 + "','" + Resultado6 + "','" + Resultado7 + "','" + Resultado8 + "','" + Valor1 + "','" + Valor2 + "','" + Valor3 + "','" + Valor4 + "','" + Valor5 + "','" + Valor6 + "','" + Valor7 + "','" + Valor8 + "','" + Baia1 + "','" + Baia2 + "','" + Baia3 + "','" + Baia4 + "','" + Baia5 + "','" + Baia6 + "','" + Baia7 + "','" + Baia8 + "','" + Observacao + "','" + Data + "')";

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
