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
        string CaminhoBancoDados = "C:/Documents and Settings/Administrador/Desktop/Dropbox/Repositorio/Controle-de-Cartuchos/Banco/Cartuchos.mdb";
        int LinhaAtual;
        string CodigoID;

        public Form_Principal()
        {
            InitializeComponent();
        }

        public void Limpar()
        {
            textBox_Nome.Text = string.Empty;
            textBox_Telefone.Text = string.Empty;

            comboBox_Produto1.Text = string.Empty;
            comboBox_Produto2.Text = string.Empty;
            comboBox_Produto3.Text = string.Empty;
            comboBox_Produto4.Text = string.Empty;
            comboBox_Produto5.Text = string.Empty;
            comboBox_Produto6.Text = string.Empty;
            comboBox_Produto7.Text = string.Empty;
            comboBox_Produto8.Text = string.Empty;

            comboBox_Servico1.Text = string.Empty;
            comboBox_Servico2.Text = string.Empty;
            comboBox_Servico3.Text = string.Empty;
            comboBox_Servico4.Text = string.Empty;
            comboBox_Servico5.Text = string.Empty;
            comboBox_Servico6.Text = string.Empty;
            comboBox_Servico7.Text = string.Empty;
            comboBox_Servico8.Text = string.Empty;

            textBox_Identificacao1.Text = string.Empty;
            textBox_Identificacao2.Text = string.Empty;
            textBox_Identificacao3.Text = string.Empty;
            textBox_Identificacao4.Text = string.Empty;
            textBox_Identificacao5.Text = string.Empty;
            textBox_Identificacao6.Text = string.Empty;
            textBox_Identificacao7.Text = string.Empty;
            textBox_Identificacao8.Text = string.Empty;

            textBox_PsEntrada1.Text = string.Empty;
            textBox_PsEntrada2.Text = string.Empty;
            textBox_PsEntrada3.Text = string.Empty;
            textBox_PsEntrada4.Text = string.Empty;
            textBox_PsEntrada5.Text = string.Empty;
            textBox_PsEntrada6.Text = string.Empty;
            textBox_PsEntrada7.Text = string.Empty;
            textBox_PsEntrada8.Text = string.Empty;

            textBox_PsSaida1.Text = string.Empty;
            textBox_PsSaida2.Text = string.Empty;
            textBox_PsSaida3.Text = string.Empty;
            textBox_PsSaida4.Text = string.Empty;
            textBox_PsSaida5.Text = string.Empty;
            textBox_PsSaida6.Text = string.Empty;
            textBox_PsSaida7.Text = string.Empty;
            textBox_PsSaida8.Text = string.Empty;

            textBox_Resultado1.Text = string.Empty;
            textBox_Resultado2.Text = string.Empty;
            textBox_Resultado3.Text = string.Empty;
            textBox_Resultado4.Text = string.Empty;
            textBox_Resultado5.Text = string.Empty;
            textBox_Resultado6.Text = string.Empty;
            textBox_Resultado7.Text = string.Empty;
            textBox_Resultado8.Text = string.Empty;

            textBox_Valor1.Text = "0";
            textBox_Valor2.Text = "0";
            textBox_Valor3.Text = "0";
            textBox_Valor4.Text = "0";
            textBox_Valor5.Text = "0";
            textBox_Valor6.Text = "0";
            textBox_Valor7.Text = "0";
            textBox_Valor8.Text = "0";

            textBox_Baia1.Text = string.Empty;
            textBox_Baia2.Text = string.Empty;
            textBox_Baia3.Text = string.Empty;
            textBox_Baia4.Text = string.Empty;
            textBox_Baia5.Text = string.Empty;
            textBox_Baia6.Text = string.Empty;
            textBox_Baia7.Text = string.Empty;
            textBox_Baia8.Text = string.Empty;

            textBox_Observacao.Text = string.Empty;

            label_Os.Text = "Os";

            button_Processar.Text = "PROCESSAR";

            textBox_Nome.Focus();
        }
        public void BancoDeDados()
        {
            

            OleDbConnection Conexao = new OleDbConnection("Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" + CaminhoBancoDados);
            DataSet Ds = new DataSet();
            DataSet Da = new DataSet();

            try
            {
                Conexao.Open();
            }
            catch (System.Exception e)
            {
                MessageBox.Show(e.Message.ToString());
            }

            if (Conexao.State == ConnectionState.Open)
            {
                OleDbDataAdapter Historico = new OleDbDataAdapter("SELECT * FROM Cartuchos", Conexao);

                Historico.Fill(Ds, "Cartuchos");

                dataGridView_Cartuchos.DataSource = Ds;
                dataGridView_Cartuchos.DataMember = "Cartuchos";

                OleDbDataAdapter Historico2 = new OleDbDataAdapter("SELECT OS, Nome, Data  FROM Cartuchos", Conexao);

                Historico2.Fill(Da, "Cartuchos");

                dataGridView_Visao.DataSource = Da;
                dataGridView_Visao.DataMember = "Cartuchos";
            }
        }

        private void Form1_Load(object sender, EventArgs e)
        {
            BancoDeDados();
            textBox_Nome.Focus();
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

            string Valor1 = textBox_Valor1.Text;
            string Valor2 = textBox_Valor2.Text;
            string Valor3 = textBox_Valor3.Text;
            string Valor4 = textBox_Valor4.Text;
            string Valor5 = textBox_Valor5.Text;
            string Valor6 = textBox_Valor6.Text;
            string Valor7 = textBox_Valor7.Text;
            string Valor8 = textBox_Valor8.Text;

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

            string ValorTotal = label_Valor_Total.Text;/*float.Parse(Valor1) + float.Parse(Valor2) + float.Parse(Valor3) + float.Parse(Valor4) + float.Parse(Valor5) + float.Parse(Valor6) + float.Parse(Valor7) + float.Parse(Valor8);*/

            //label_Valor_Total.Text = Convert.ToString(ValorTotal);

            //string ValorTotalBd = Convert.ToString(ValorTotal);
            if (button_Processar.Text == "PROCESSAR")
            {
                string Conexao = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" + CaminhoBancoDados;
                string Inserir = "INSERT INTO Cartuchos(Nome, Telefone, Produto1, Produto2, Produto3, Produto4, Produto5, Produto6, Produto7, Produto8, Servico1 ,Servico2, Servico3, Servico4, Servico5, Servico6, Servico7, Servico8, Identificacao1, Identificacao2, Identificacao3, Identificacao4, Identificacao5, Identificacao6, Identificacao7, Identificacao8, PSEntrada1, PSEntrada2, PSEntrada3, PSEntrada4, PSEntrada5, PSEntrada6, PSEntrada7, PSEntrada8, PSSaida1, PSSaida2, PSSaida3, PSSaida4, PSSaida5, PSSaida6, PSSaida7, PSSaida8, Resultado1, Resultado2, Resultado3, Resultado4, Resultado5, Resultado6, Resultado7, Resultado8, Valor1, Valor2, Valor3, Valor4, Valor5, Valor6, Valor7, Valor8, Baia1, Baia2, Baia3, Baia4, Baia5, Baia6, Baia7, Baia8, Observacao, Data)" + " VALUES('" + Nome + "' , '" + Telefone + "','" + Produto1 + "','" + Produto2 + "','" + Produto3 + "','" + Produto4 + "','" + Produto5 + "','" + Produto6 + "','" + Produto7 + "','" + Produto8 + "','" + Servico1 + "','" + Servico2 + "','" + Servico3 + "','" + Servico4 + "','" + Servico5 + "','" + Servico6 + "','" + Servico7 + "','" + Servico8 + "','" + Identificacao1 + "','" + Identificacao2 + "','" + Identificacao3 + "','" + Identificacao4 + "','" + Identificacao5 + "','" + Identificacao6 + "','" + Identificacao7 + "','" + Identificacao8 + "','" + PSEntrada1 + "','" + PSEntrada2 + "','" + PSEntrada3 + "','" + PSEntrada4 + "','" + PSEntrada5 + "','" + PSEntrada6 + "','" + PSEntrada7 + "','" + PSEntrada8 + "','" + PSSaida1 + "','" + PSSaida2 + "','" + PSSaida3 + "','" + PSSaida4 + "','" + PSSaida5 + "','" + PSSaida6 + "','" + PSSaida7 + "','" + PSSaida8 + "','" + Resultado1 + "','" + Resultado2 + "','" + Resultado3 + "','" + Resultado4 + "','" + Resultado5 + "','" + Resultado6 + "','" + Resultado7 + "','" + Resultado8 + "','" + Valor1 + "','" + Valor2 + "','" + Valor3 + "','" + Valor4 + "','" + Valor5 + "','" + Valor6 + "','" + Valor7 + "','" + Valor8 + "','" + Baia1 + "','" + Baia2 + "','" + Baia3 + "','" + Baia4 + "','" + Baia5 + "','" + Baia6 + "','" + Baia7 + "','" + Baia8 + "','" + Observacao + "','" + Data + "')";


                //cria a conexão com o banco de dados
                OleDbConnection ConexaoBD = new OleDbConnection(Conexao);

                //Cria o comando que inicia a query
                OleDbCommand cmdInserir = new OleDbCommand(Inserir, ConexaoBD);
                try
                {
                    // abre o banco
                    ConexaoBD.Open();
                    // executa a query
                    cmdInserir.ExecuteNonQuery();
                
                }
                //Trata a exceção
                catch (OleDbException Erro)
                {
                    MessageBox.Show("Error: " + Erro.Message);
                }
                finally
                {
                    //fecha a conexao
                    ConexaoBD.Close();
                    dataGridView_Cartuchos.Update();
                    BancoDeDados();
                }
                Limpar();
            }
            else
            {
                string Conexao = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" + CaminhoBancoDados;
                CodigoID = dataGridView_Cartuchos[0, LinhaAtual].Value.ToString(); 

                string Editar = "UPDATE Cartuchos SET Nome = '" + Nome + "' , Telefone = '" + Telefone + "', Produto1 = '" + Produto1 + "' , Produto2 = '" + Produto2 + "' , Produto3 = '" + Produto3 + "' , Produto4 = '" + Produto4 + "' , Produto5 = '" + Produto5 + "' , Produto6 = '" + Produto6 + "' , Produto7 = '" + Produto7 + "' , Produto8 = '" + Produto8 + "' , Servico1 = '" + Servico1 + "',Servico2 = '" + Servico2 + "',Servico3 = '" + Servico3 + "',Servico4 = '" + Servico4 + "',Servico5 = '" + Servico5 + "' , Servico6 = '" + Servico6 + "' , Servico7 = '" + Servico7 + "' , Servico8 = '" + Servico8 + "' , Identificacao1 = '" + Identificacao1 + "' , Identificacao2 = '" + Identificacao2 + "' , Identificacao3 = '" + Identificacao3 + "' , Identificacao4 = '" + Identificacao4 + "' , Identificacao5 = '" + Identificacao5 + "' , Identificacao6 = '" + Identificacao6 + "' , Identificacao7 = '" + Identificacao7 + "',Identificacao8 = '" + Identificacao8 + "',PSEntrada1 = '" + PSEntrada1 + "',PSEntrada2 = '" + PSEntrada2 + "' , PSEntrada3 = '" + PSEntrada3 + "' , PSEntrada4 = '" + PSEntrada4 + "' , PSEntrada5 = '" + PSEntrada5 + "' , PSEntrada6 = '" + PSEntrada6 + "' , PSEntrada7 = '" + PSEntrada7 + "' , PSEntrada8 = '" + PSEntrada8 + "' , PSSaida1 = '" + PSSaida1 + "' , PSSaida2 = '" + PSSaida2 + "' , PSSaida3 = '" + PSSaida3 + "' , PSSaida4 = '" + PSSaida4 + "' , PSSaida5 = '" + PSSaida5 + "' , PSSaida6 = '" + PSSaida6 + "' , PSSaida7 = '" + PSSaida7 + "',PSSaida8 = '" + PSSaida8 + "' , Resultado1 = '" + Resultado1 + "', Resultado2 = '" + Resultado2 + "' , Resultado3 = '" + Resultado3 + "', Resultado4 = '" + Resultado4 + "', Resultado5 = '" + Resultado5 + "' , Resultado6 = '" + Resultado6 + "' , Resultado7 = '" + Resultado7 + "' , Resultado8 = '" + Resultado8 + "' , Valor1 = '" + Valor1 + "' , Valor2 = '" + Valor2 + "' , Valor3 = '" + Valor3 + "' , Valor4 = '" + Valor4 + "' , Valor5 = '" + Valor5 + "' , Valor6 = '" + Valor6 + "',Valor7 = '" + Valor7 + "',Valor8 = '" + Valor8 + "' ,Baia1 = '" + Baia1 + "' , Baia2 = '" + Baia2 + "' , Baia3 = '" + Baia3 + "' , Baia4 = '" + Baia4 + "' , Baia5 = '" + Baia5 + "', Baia6 = '" + Baia6 + "' , Baia7 = '" + Baia7 + "' , Baia8 = '" + Baia8 + "' , Observacao = '" + Observacao + "' , Data = '" + Data + "' WHERE OS= " + int.Parse(CodigoID) + "";

                //cria a conexão com o banco de dados
                OleDbConnection ConexaoBD = new OleDbConnection(Conexao);

                //Cria o comando que inicia a query
                OleDbCommand cmdEditar = new OleDbCommand(Editar, ConexaoBD);

                try
                {
                    ConexaoBD.Open();
                    cmdEditar.ExecuteNonQuery();
                }
                catch (OleDbException ex)
                {
                    MessageBox.Show("Error : " + ex.Message);
                }
                finally
                {
                    ConexaoBD.Close();
                    BancoDeDados();
                    button_Processar.Text = "PROCESSAR";
                }
                Limpar();
            }
        }

        private void dataGridView_Cartuchos_Click(object sender, DataGridViewCellEventArgs e)
        {
            LinhaAtual = int.Parse(e.RowIndex.ToString());

            if (LinhaAtual >= 0)
            {
                label_Os.Text = dataGridView_Cartuchos[0, LinhaAtual].Value.ToString();
                textBox_Nome.Text = dataGridView_Cartuchos[1, LinhaAtual].Value.ToString();
                textBox_Telefone.Text = dataGridView_Cartuchos[2, LinhaAtual].Value.ToString();

                comboBox_Produto1.Text = dataGridView_Cartuchos[3, LinhaAtual].Value.ToString();
                comboBox_Produto2.Text = dataGridView_Cartuchos[4, LinhaAtual].Value.ToString();
                comboBox_Produto3.Text = dataGridView_Cartuchos[5, LinhaAtual].Value.ToString();
                comboBox_Produto4.Text = dataGridView_Cartuchos[6, LinhaAtual].Value.ToString();
                comboBox_Produto5.Text = dataGridView_Cartuchos[7, LinhaAtual].Value.ToString();
                comboBox_Produto6.Text = dataGridView_Cartuchos[8, LinhaAtual].Value.ToString();
                comboBox_Produto7.Text = dataGridView_Cartuchos[9, LinhaAtual].Value.ToString();
                comboBox_Produto8.Text = dataGridView_Cartuchos[10, LinhaAtual].Value.ToString();

                comboBox_Servico1.Text = dataGridView_Cartuchos[11, LinhaAtual].Value.ToString();
                comboBox_Servico2.Text = dataGridView_Cartuchos[12, LinhaAtual].Value.ToString();
                comboBox_Servico3.Text = dataGridView_Cartuchos[13, LinhaAtual].Value.ToString();
                comboBox_Servico4.Text = dataGridView_Cartuchos[14, LinhaAtual].Value.ToString();
                comboBox_Servico5.Text = dataGridView_Cartuchos[15, LinhaAtual].Value.ToString();
                comboBox_Servico6.Text = dataGridView_Cartuchos[16, LinhaAtual].Value.ToString();
                comboBox_Servico7.Text = dataGridView_Cartuchos[17, LinhaAtual].Value.ToString();
                comboBox_Servico8.Text = dataGridView_Cartuchos[18, LinhaAtual].Value.ToString();

                textBox_Identificacao1.Text = dataGridView_Cartuchos[19, LinhaAtual].Value.ToString();
                textBox_Identificacao2.Text = dataGridView_Cartuchos[20, LinhaAtual].Value.ToString();
                textBox_Identificacao3.Text = dataGridView_Cartuchos[21, LinhaAtual].Value.ToString();
                textBox_Identificacao4.Text = dataGridView_Cartuchos[22, LinhaAtual].Value.ToString();
                textBox_Identificacao5.Text = dataGridView_Cartuchos[23, LinhaAtual].Value.ToString();
                textBox_Identificacao6.Text = dataGridView_Cartuchos[24, LinhaAtual].Value.ToString();
                textBox_Identificacao7.Text = dataGridView_Cartuchos[25, LinhaAtual].Value.ToString();
                textBox_Identificacao8.Text = dataGridView_Cartuchos[26, LinhaAtual].Value.ToString();

                textBox_PsEntrada1.Text = dataGridView_Cartuchos[27, LinhaAtual].Value.ToString();
                textBox_PsEntrada2.Text = dataGridView_Cartuchos[28, LinhaAtual].Value.ToString();
                textBox_PsEntrada3.Text = dataGridView_Cartuchos[29, LinhaAtual].Value.ToString();
                textBox_PsEntrada4.Text = dataGridView_Cartuchos[30, LinhaAtual].Value.ToString();
                textBox_PsEntrada5.Text = dataGridView_Cartuchos[31, LinhaAtual].Value.ToString();
                textBox_PsEntrada6.Text = dataGridView_Cartuchos[32, LinhaAtual].Value.ToString();
                textBox_PsEntrada7.Text = dataGridView_Cartuchos[33, LinhaAtual].Value.ToString();
                textBox_PsEntrada8.Text = dataGridView_Cartuchos[34, LinhaAtual].Value.ToString();

                textBox_PsSaida1.Text = dataGridView_Cartuchos[35, LinhaAtual].Value.ToString();
                textBox_PsSaida2.Text = dataGridView_Cartuchos[36, LinhaAtual].Value.ToString();
                textBox_PsSaida3.Text = dataGridView_Cartuchos[37, LinhaAtual].Value.ToString();
                textBox_PsSaida4.Text = dataGridView_Cartuchos[38, LinhaAtual].Value.ToString();
                textBox_PsSaida5.Text = dataGridView_Cartuchos[39, LinhaAtual].Value.ToString();
                textBox_PsSaida6.Text = dataGridView_Cartuchos[40, LinhaAtual].Value.ToString();
                textBox_PsSaida7.Text = dataGridView_Cartuchos[41, LinhaAtual].Value.ToString();
                textBox_PsSaida8.Text = dataGridView_Cartuchos[42, LinhaAtual].Value.ToString();

                textBox_Resultado1.Text = dataGridView_Cartuchos[43, LinhaAtual].Value.ToString();
                textBox_Resultado2.Text = dataGridView_Cartuchos[44, LinhaAtual].Value.ToString();
                textBox_Resultado3.Text = dataGridView_Cartuchos[45, LinhaAtual].Value.ToString();
                textBox_Resultado4.Text = dataGridView_Cartuchos[46, LinhaAtual].Value.ToString();
                textBox_Resultado5.Text = dataGridView_Cartuchos[47, LinhaAtual].Value.ToString();
                textBox_Resultado6.Text = dataGridView_Cartuchos[48, LinhaAtual].Value.ToString();
                textBox_Resultado7.Text = dataGridView_Cartuchos[49, LinhaAtual].Value.ToString();
                textBox_Resultado8.Text = dataGridView_Cartuchos[50, LinhaAtual].Value.ToString();

                textBox_Valor1.Text = dataGridView_Cartuchos[51, LinhaAtual].Value.ToString();
                textBox_Valor2.Text = dataGridView_Cartuchos[52, LinhaAtual].Value.ToString();
                textBox_Valor3.Text = dataGridView_Cartuchos[53, LinhaAtual].Value.ToString();
                textBox_Valor4.Text = dataGridView_Cartuchos[54, LinhaAtual].Value.ToString();
                textBox_Valor5.Text = dataGridView_Cartuchos[55, LinhaAtual].Value.ToString();
                textBox_Valor6.Text = dataGridView_Cartuchos[56, LinhaAtual].Value.ToString();
                textBox_Valor7.Text = dataGridView_Cartuchos[57, LinhaAtual].Value.ToString();
                textBox_Valor8.Text = dataGridView_Cartuchos[58, LinhaAtual].Value.ToString();

                textBox_Baia1.Text = dataGridView_Cartuchos[59, LinhaAtual].Value.ToString();
                textBox_Baia2.Text = dataGridView_Cartuchos[60, LinhaAtual].Value.ToString();
                textBox_Baia3.Text = dataGridView_Cartuchos[61, LinhaAtual].Value.ToString();
                textBox_Baia4.Text = dataGridView_Cartuchos[62, LinhaAtual].Value.ToString();
                textBox_Baia5.Text = dataGridView_Cartuchos[63, LinhaAtual].Value.ToString();
                textBox_Baia6.Text = dataGridView_Cartuchos[64, LinhaAtual].Value.ToString();
                textBox_Baia7.Text = dataGridView_Cartuchos[65, LinhaAtual].Value.ToString();
                textBox_Baia8.Text = dataGridView_Cartuchos[66, LinhaAtual].Value.ToString();

                textBox_Observacao.Text = dataGridView_Cartuchos[67, LinhaAtual].Value.ToString();

                dateTimePicker_Data.Text = dataGridView_Cartuchos[68, LinhaAtual].Value.ToString();

                label_Valor_Total.Text = dataGridView_Cartuchos[69, LinhaAtual].Value.ToString();

                button_Processar.Text = "SALVAR";

            }
            
        }

        private void comboBox_Servico1_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (comboBox_Produto1.Text == "Cartucho" && comboBox_Servico1.Text == "Recarga")
                textBox_Valor1.Text = "15";
            else if(comboBox_Produto1.Text == "Toner" && comboBox_Servico1.Text == "Recarga Samsung")
                textBox_Valor1.Text = "85";
            else if (comboBox_Produto1.Text == "Toner" && comboBox_Servico1.Text == "Recarga HP")
                textBox_Valor1.Text = "60";
            else if (comboBox_Produto1.Text == "Toner" && comboBox_Servico1.Text == "Recarga Lexmark")
                textBox_Valor1.Text = "70";
            else if (comboBox_Produto1.Text == "Toner" && comboBox_Servico1.Text == "Recarga Brother")
                textBox_Valor1.Text = "60";
        }

        private void groupBox2_Enter(object sender, EventArgs e)
        {

        }

        private void textBox1_TextChanged(object sender, EventArgs e)
        {
            string SqlCmd="";

            if (textBox_Pesquisa.Text != "")
                SqlCmd = "SELECT * FROM Cartuchos WHERE Nome LIKE '" + textBox_Pesquisa.Text + "%'";

            OleDbConnection Conexao = new OleDbConnection("Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" + CaminhoBancoDados);

            DataSet Ds = new DataSet();
            try
            {
                Conexao.Open();
            }
            catch (System.Exception Erro)
            {
                MessageBox.Show(Erro.Message.ToString());
            }

            if (Conexao.State == ConnectionState.Open)
            {
                if (textBox_Pesquisa.Text != null && textBox_Pesquisa.Text != "")
                {
                    OleDbDataAdapter Historico = new OleDbDataAdapter(SqlCmd, Conexao);

                    Historico.Fill(Ds, "Cartuchos");

                    dataGridView_Visao.DataSource = Ds;
                    dataGridView_Visao.DataMember = "Cartuchos";
                }
                else
                    BancoDeDados();

            }

        }

        private void dataGridView_Cartuchos_CellContentClick(object sender, DataGridViewCellEventArgs e)
        {

        }

        private void button_Novo_Click(object sender, EventArgs e)
        {

            Limpar();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            dataGridView_Visao.Update();
            dataGridView_Cartuchos.Update();
        }


    }
}
