﻿using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;

using System.Drawing.Printing;

using System.Data;
using System.Data.OleDb;

using System.IO;

namespace Controle_de_Cartuchos
{
    public partial class Form_Principal : Form
    {
        int LinhaAtual;
        string CodigoID;

        public Form_Principal()
        {
            InitializeComponent();
        }

        public void Cartucho01()
        {
            if (comboBox_Produto1.Text == "Cartucho" && comboBox_Servico1.Text == "Recarga")
                textBox_Valor1.Text = "15,00";
            else if (comboBox_Produto1.Text == "Toner" && comboBox_Servico1.Text == "Rcrg Samsung")
                textBox_Valor1.Text = "85,00";
            else if (comboBox_Produto1.Text == "Toner" && comboBox_Servico1.Text == "Rcrg HP")
                textBox_Valor1.Text = "60,00";
            else if (comboBox_Produto1.Text == "Toner" && comboBox_Servico1.Text == "Rcrg Lexmark")
                textBox_Valor1.Text = "70,00";
            else if (comboBox_Produto1.Text == "Toner" && comboBox_Servico1.Text == "Rcrg Brother")
                textBox_Valor1.Text = "60,00";
            else
                textBox_Valor1.Text = "0,00";

        }

        public void Cartucho02()
        {
            if (comboBox_Produto2.Text == "Cartucho" && comboBox_Servico2.Text == "Recarga")
                textBox_Valor2.Text = "15,00";
            else if (comboBox_Produto2.Text == "Toner" && comboBox_Servico2.Text == "Rcrg Samsung")
                textBox_Valor2.Text = "85,00";
            else if (comboBox_Produto2.Text == "Toner" && comboBox_Servico2.Text == "Rcrg HP")
                textBox_Valor2.Text = "60,00";
            else if (comboBox_Produto2.Text == "Toner" && comboBox_Servico2.Text == "Rcrg Lexmark")
                textBox_Valor2.Text = "70,00";
            else if (comboBox_Produto2.Text == "Toner" && comboBox_Servico2.Text == "Rcrg Brother")
                textBox_Valor2.Text = "60,00";
            else
                textBox_Valor2.Text = "0,00";
        }

        public void Cartucho03()
        {
            if (comboBox_Produto3.Text == "Cartucho" && comboBox_Servico3.Text == "Recarga")
                textBox_Valor3.Text = "15,00";
            else if (comboBox_Produto3.Text == "Toner" && comboBox_Servico3.Text == "Rcrg Samsung")
                textBox_Valor3.Text = "85,00";
            else if (comboBox_Produto3.Text == "Toner" && comboBox_Servico3.Text == "Rcrg HP")
                textBox_Valor3.Text = "60,00";
            else if (comboBox_Produto3.Text == "Toner" && comboBox_Servico3.Text == "Rcrg Lexmark")
                textBox_Valor3.Text = "70,00";
            else if (comboBox_Produto3.Text == "Toner" && comboBox_Servico3.Text == "Rcrg Brother")
                textBox_Valor3.Text = "60,00";
            else
                textBox_Valor3.Text = "0,00";
        }

        public void Cartucho04()
        {
            if (comboBox_Produto4.Text == "Cartucho" && comboBox_Servico4.Text == "Recarga")
                textBox_Valor4.Text = "15,00";
            else if (comboBox_Produto4.Text == "Toner" && comboBox_Servico4.Text == "Rcrg Samsung")
                textBox_Valor4.Text = "85,00";
            else if (comboBox_Produto4.Text == "Toner" && comboBox_Servico4.Text == "Rcrg HP")
                textBox_Valor4.Text = "60,00";
            else if (comboBox_Produto4.Text == "Toner" && comboBox_Servico4.Text == "Rcrg Lexmark")
                textBox_Valor4.Text = "70,00";
            else if (comboBox_Produto4.Text == "Toner" && comboBox_Servico4.Text == "Rcrg Brother")
                textBox_Valor4.Text = "60,00";
            else
                textBox_Valor4.Text = "0,00";
        }

        public void Cartucho05()
        {
            if (comboBox_Produto5.Text == "Cartucho" && comboBox_Servico5.Text == "Recarga")
                textBox_Valor5.Text = "15,00";
            else if (comboBox_Produto5.Text == "Toner" && comboBox_Servico5.Text == "Rcrg Samsung")
                textBox_Valor5.Text = "85,00";
            else if (comboBox_Produto5.Text == "Toner" && comboBox_Servico5.Text == "Rcrg HP")
                textBox_Valor5.Text = "60,00";
            else if (comboBox_Produto5.Text == "Toner" && comboBox_Servico5.Text == "Rcrg Lexmark")
                textBox_Valor5.Text = "70,00";
            else if (comboBox_Produto5.Text == "Toner" && comboBox_Servico5.Text == "Rcrg Brother")
                textBox_Valor5.Text = "60,00";
            else
                textBox_Valor5.Text = "0,00";
        }

        public void Cartucho06()
        {
            if (comboBox_Produto6.Text == "Cartucho" && comboBox_Servico6.Text == "Recarga")
                textBox_Valor6.Text = "15,00";
            else if (comboBox_Produto6.Text == "Toner" && comboBox_Servico6.Text == "Rcrg Samsung")
                textBox_Valor6.Text = "85,00";
            else if (comboBox_Produto6.Text == "Toner" && comboBox_Servico6.Text == "Rcrg HP")
                textBox_Valor6.Text = "60,00";
            else if (comboBox_Produto6.Text == "Toner" && comboBox_Servico6.Text == "Rcrg Lexmark")
                textBox_Valor6.Text = "70,00";
            else if (comboBox_Produto6.Text == "Toner" && comboBox_Servico6.Text == "Rcrg Brother")
                textBox_Valor6.Text = "60,00";
            else
                textBox_Valor6.Text = "0,00";
        }

        public void Cartucho07()
        {
            if (comboBox_Produto7.Text == "Cartucho" && comboBox_Servico7.Text == "Recarga")
                textBox_Valor7.Text = "15,00";
            else if (comboBox_Produto1.Text == "Toner" && comboBox_Servico7.Text == "Rcrg Samsung")
                textBox_Valor7.Text = "85,00";
            else if (comboBox_Produto1.Text == "Toner" && comboBox_Servico7.Text == "Rcrg HP")
                textBox_Valor7.Text = "60,00";
            else if (comboBox_Produto1.Text == "Toner" && comboBox_Servico7.Text == "Rcrg Lexmark")
                textBox_Valor7.Text = "70,00";
            else if (comboBox_Produto1.Text == "Toner" && comboBox_Servico7.Text == "Rcrg Brother")
                textBox_Valor7.Text = "60,00";
            else
                textBox_Valor7.Text = "0,00";
        }

        public void Cartucho08()
        {
            if (comboBox_Produto8.Text == "Cartucho" && comboBox_Servico8.Text == "Recarga")
                textBox_Valor8.Text = "15,00";
            else if (comboBox_Produto8.Text == "Toner" && comboBox_Servico8.Text == "Rcrg Samsung")
                textBox_Valor8.Text = "85,00";
            else if (comboBox_Produto8.Text == "Toner" && comboBox_Servico8.Text == "Rcrg HP")
                textBox_Valor8.Text = "60,00";
            else if (comboBox_Produto8.Text == "Toner" && comboBox_Servico8.Text == "Rcrg Lexmark")
                textBox_Valor8.Text = "70,00";
            else if (comboBox_Produto8.Text == "Toner" && comboBox_Servico8.Text == "Rcrg Brother")
                textBox_Valor8.Text = "60,00";
            else
                textBox_Valor8.Text = "0,00";
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

            textBox_Valor1.Enabled = false;
            textBox_Valor2.Enabled = false;
            textBox_Valor3.Enabled = false;
            textBox_Valor4.Enabled = false;
            textBox_Valor5.Enabled = false;
            textBox_Valor6.Enabled = false;
            textBox_Valor7.Enabled = false;
            textBox_Valor8.Enabled = false;
            

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

            textBox_Valor1.Enabled = true;
            textBox_Valor2.Enabled = true;
            textBox_Valor3.Enabled = true;
            textBox_Valor4.Enabled = true;
            textBox_Valor5.Enabled = true;
            textBox_Valor6.Enabled = true;
            textBox_Valor7.Enabled = true;
            textBox_Valor8.Enabled = true;

            textBox_Observacao.Text = string.Empty;

            label_Os.Text = "Os";

            textBox_ValorTotal.Text = string.Empty;

            comboBox_Encerrada.Text = "Nao";

            button_Processar.Text = "PROCESSAR";

            textBox_Nome.Focus();
        }
        public string CaminhoBancoDados()
        {
            //Abre o Arquivo que contem o Banco de Dados
            TextReader Arquivo = File.OpenText(@"C:\Controle-de-Cartuchos\Controle-de-Cartuchos\ArqID.txt");
            string Caminho = Arquivo.ReadLine();
            string CaminhoBanco;

            //Remove as aspas duplas da string 
            int indice = Caminho.Count();
            string Palavra = "";

            for (int I = 0; I < indice; ++I)
            {
                if (Caminho[I] != '\"')
                {
                    Palavra = Palavra + Caminho[I];
                }
            }
            textBox_Caminho.Text = Palavra;
            CaminhoBanco = textBox_Caminho.Text;
            Arquivo.Close();

            return CaminhoBanco;
        }
        public void BancoDeDados()
        {
            //Cria conexao com o banco de dados
            OleDbConnection Conexao = new OleDbConnection("Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" + CaminhoBancoDados());
            DataSet Ds = new DataSet();
            DataSet Da = new DataSet();
     
            //Verifica se o banco de dados foi aberto
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

                OleDbDataAdapter Historico_Visao = new OleDbDataAdapter("SELECT OS, Nome, Data, Encerramento  FROM Cartuchos", Conexao);

                Historico_Visao.Fill(Da, "Cartuchos_Visao");

                dataGridView_Visao.DataSource = Da;
                dataGridView_Visao.DataMember = "Cartuchos_Visao";

            }
        }

        private void Form1_Load(object sender, EventArgs e)
        {
            BancoDeDados();
            textBox_Nome.Focus();
            int GridMax = dataGridView_Cartuchos.RowCount;
            label_Os.Text = Convert.ToString(GridMax);

            if (comboBox_Encerrada.Text == "Sim")
                dateTimePicker_Encerramento.Visible = true;
            else
                dateTimePicker_Encerramento.Visible = false;
            

        }

        private void button_Processar_Click(object sender, EventArgs e)
        {
            if ((textBox_Valor1.Text != "") && (textBox_Valor2.Text != "") && (textBox_Valor3.Text != "") && (textBox_Valor4.Text) != "" && (textBox_Valor5.Text != "") && (textBox_Valor6.Text != "") && (textBox_Valor7.Text != "") && (textBox_Valor8.Text != ""))
            {
                //Passa todos os dados do layaout para sua determinadas strings
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

                string Encerrada = comboBox_Encerrada.Text;

                string Observacao = textBox_Observacao.Text;

                string Data = dateTimePicker_Data.Value.ToShortDateString();

                string Encerramento = ". . .";

                //Verifica se a OS Ja foi encerrada
                if (comboBox_Encerrada.Text == "Sim")
                {
                    Encerramento = dateTimePicker_Encerramento.Value.ToShortDateString();
                }

                else
                {
                    Encerramento = ". . .";
                }

                //Calcula O Total de Valores
                float ValorTotal =  Valor1+Valor2+Valor3+Valor4+Valor5+Valor6+Valor7+Valor8;
                textBox_ValorTotal.Text = Convert.ToString(ValorTotal) + ",00";

                if (button_Processar.Text == "SALVAR")
                {
                    string Conexao = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" + CaminhoBancoDados();
                    string Inserir = "INSERT INTO Cartuchos(Nome, Telefone, Produto1, Produto2, Produto3, Produto4, Produto5, Produto6, Produto7, Produto8, Servico1 ,Servico2, Servico3, Servico4, Servico5, Servico6, Servico7, Servico8, Identificacao1, Identificacao2, Identificacao3, Identificacao4, Identificacao5, Identificacao6, Identificacao7, Identificacao8, PSEntrada1, PSEntrada2, PSEntrada3, PSEntrada4, PSEntrada5, PSEntrada6, PSEntrada7, PSEntrada8, PSSaida1, PSSaida2, PSSaida3, PSSaida4, PSSaida5, PSSaida6, PSSaida7, PSSaida8, Resultado1, Resultado2, Resultado3, Resultado4, Resultado5, Resultado6, Resultado7, Resultado8, Valor1, Valor2, Valor3, Valor4, Valor5, Valor6, Valor7, Valor8, Baia1, Baia2, Baia3, Baia4, Baia5, Baia6, Baia7, Baia8, Observacao, Data, ValorTotal, Encerrada, Encerramento)" + " VALUES('" + Nome + "' , '" + Telefone + "','" + Produto1 + "','" + Produto2 + "','" + Produto3 + "','" + Produto4 + "','" + Produto5 + "','" + Produto6 + "','" + Produto7 + "','" + Produto8 + "','" + Servico1 + "','" + Servico2 + "','" + Servico3 + "','" + Servico4 + "','" + Servico5 + "','" + Servico6 + "','" + Servico7 + "','" + Servico8 + "','" + Identificacao1 + "','" + Identificacao2 + "','" + Identificacao3 + "','" + Identificacao4 + "','" + Identificacao5 + "','" + Identificacao6 + "','" + Identificacao7 + "','" + Identificacao8 + "','" + PSEntrada1 + "','" + PSEntrada2 + "','" + PSEntrada3 + "','" + PSEntrada4 + "','" + PSEntrada5 + "','" + PSEntrada6 + "','" + PSEntrada7 + "','" + PSEntrada8 + "','" + PSSaida1 + "','" + PSSaida2 + "','" + PSSaida3 + "','" + PSSaida4 + "','" + PSSaida5 + "','" + PSSaida6 + "','" + PSSaida7 + "','" + PSSaida8 + "','" + Resultado1 + "','" + Resultado2 + "','" + Resultado3 + "','" + Resultado4 + "','" + Resultado5 + "','" + Resultado6 + "','" + Resultado7 + "','" + Resultado8 + "','" + Valor1 + "','" + Valor2 + "','" + Valor3 + "','" + Valor4 + "','" + Valor5 + "','" + Valor6 + "','" + Valor7 + "','" + Valor8 + "','" + Baia1 + "','" + Baia2 + "','" + Baia3 + "','" + Baia4 + "','" + Baia5 + "','" + Baia6 + "','" + Baia7 + "','" + Baia8 + "','" + Observacao + "','" + Data + "' , '" + ValorTotal + "','" + Encerrada + "','"+Encerramento+"')";


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
                        BancoDeDados();
                        dataGridView_Cartuchos.Focus();
                        button_Processar.Enabled = false;
                        dateTimePicker_Encerramento.Visible = false;
                        
                    }
                }

                 //Else que controla a edicao
                else
                {
                    //Abre a Conexao com o Banco de Dados
                    string Conexao = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" + CaminhoBancoDados();

                    CodigoID = dataGridView_Cartuchos[0, LinhaAtual].Value.ToString();

                    if (comboBox_Encerrada.Text == "Sim")
                        Encerramento = dateTimePicker_Encerramento.Value.ToShortDateString();
                    else
                        Encerramento = ". . .";

                    string Editar = "UPDATE Cartuchos SET Nome = '" + Nome + "' , Telefone = '" + Telefone + "', Produto1 = '" + Produto1 + "' , Produto2 = '" + Produto2 + "' , Produto3 = '" + Produto3 + "' , Produto4 = '" + Produto4 + "' , Produto5 = '" + Produto5 + "' , Produto6 = '" + Produto6 + "' , Produto7 = '" + Produto7 + "' , Produto8 = '" + Produto8 + "' , Servico1 = '" + Servico1 + "',Servico2 = '" + Servico2 + "',Servico3 = '" + Servico3 + "',Servico4 = '" + Servico4 + "',Servico5 = '" + Servico5 + "' , Servico6 = '" + Servico6 + "' , Servico7 = '" + Servico7 + "' , Servico8 = '" + Servico8 + "' , Identificacao1 = '" + Identificacao1 + "' , Identificacao2 = '" + Identificacao2 + "' , Identificacao3 = '" + Identificacao3 + "' , Identificacao4 = '" + Identificacao4 + "' , Identificacao5 = '" + Identificacao5 + "' , Identificacao6 = '" + Identificacao6 + "' , Identificacao7 = '" + Identificacao7 + "',Identificacao8 = '" + Identificacao8 + "',PSEntrada1 = '" + PSEntrada1 + "',PSEntrada2 = '" + PSEntrada2 + "' , PSEntrada3 = '" + PSEntrada3 + "' , PSEntrada4 = '" + PSEntrada4 + "' , PSEntrada5 = '" + PSEntrada5 + "' , PSEntrada6 = '" + PSEntrada6 + "' , PSEntrada7 = '" + PSEntrada7 + "' , PSEntrada8 = '" + PSEntrada8 + "' , PSSaida1 = '" + PSSaida1 + "' , PSSaida2 = '" + PSSaida2 + "' , PSSaida3 = '" + PSSaida3 + "' , PSSaida4 = '" + PSSaida4 + "' , PSSaida5 = '" + PSSaida5 + "' , PSSaida6 = '" + PSSaida6 + "' , PSSaida7 = '" + PSSaida7 + "',PSSaida8 = '" + PSSaida8 + "' , Resultado1 = '" + Resultado1 + "', Resultado2 = '" + Resultado2 + "' , Resultado3 = '" + Resultado3 + "', Resultado4 = '" + Resultado4 + "', Resultado5 = '" + Resultado5 + "' , Resultado6 = '" + Resultado6 + "' , Resultado7 = '" + Resultado7 + "' , Resultado8 = '" + Resultado8 + "' , Valor1 = '" + Valor1 + "' , Valor2 = '" + Valor2 + "' , Valor3 = '" + Valor3 + "' , Valor4 = '" + Valor4 + "' , Valor5 = '" + Valor5 + "' , Valor6 = '" + Valor6 + "',Valor7 = '" + Valor7 + "',Valor8 = '" + Valor8 + "' ,Baia1 = '" + Baia1 + "' , Baia2 = '" + Baia2 + "' , Baia3 = '" + Baia3 + "' , Baia4 = '" + Baia4 + "' , Baia5 = '" + Baia5 + "', Baia6 = '" + Baia6 + "' , Baia7 = '" + Baia7 + "' , Baia8 = '" + Baia8 + "' , Observacao = '" + Observacao + "' , Data= '" + Data + "', ValorTotal = '" + ValorTotal + "', Encerrada = '" + Encerrada + "', Encerramento = '" + Encerramento + "' WHERE OS= " + int.Parse(CodigoID) + "";

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
                        dateTimePicker_Encerramento.Visible = false;
                    }
                }
            }
            else
                label_Valor.ForeColor = Color.Red;
        }

        private void dataGridView_Cartuchos_Click(object sender, DataGridViewCellEventArgs e)
        {
            //salva a posiçao da linha do data grid 
            LinhaAtual = int.Parse(e.RowIndex.ToString());
            int GridMax = dataGridView_Cartuchos.RowCount - 1;

            if (LinhaAtual >= 0 && LinhaAtual < GridMax )
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

                textBox_ValorTotal.Text = dataGridView_Cartuchos[69, LinhaAtual].Value.ToString()+",00";

                comboBox_Encerrada.Text = dataGridView_Cartuchos[70, LinhaAtual].Value.ToString();

                button_Processar.Enabled = true;

                button_Processar.Text = "EDITAR";
                dateTimePicker_Encerramento.Visible = false;

            }
            
        }

        private void textBox1_TextChanged(object sender, EventArgs e)
        {
            string SqlCms="";
            string SqlCmd_Visao = "";
            OleDbConnection Conexao = new OleDbConnection("Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" + CaminhoBancoDados());

            DataSet Ds = new DataSet();
            DataSet Ds_Visao = new DataSet();

            try
            {
                Conexao.Open();
            }
            catch (System.Exception Erro)
            {
                MessageBox.Show(Erro.Message.ToString());
            }

            if (textBox_Pesquisa.Text != "")
            {
                SqlCmd_Visao = "SELECT OS,Nome,Data,Telefone FROM Cartuchos WHERE Nome LIKE '" + textBox_Pesquisa.Text + "%'";
                SqlCms = "SELECT * FROM Cartuchos WHERE Nome LIKE '" + textBox_Pesquisa.Text + "%'";
            }   

            if (Conexao.State == ConnectionState.Open)
            {
                if (textBox_Pesquisa.Text != null && textBox_Pesquisa.Text != "")
                {
                    OleDbDataAdapter Historico_Visao = new OleDbDataAdapter(SqlCmd_Visao, Conexao);

                    OleDbDataAdapter Historico = new OleDbDataAdapter(SqlCms, Conexao);

                    Historico.Fill(Ds, "Cartuchos");

                    dataGridView_Cartuchos.DataSource = Ds;
                    dataGridView_Cartuchos.DataMember = "Cartuchos";

                    Historico_Visao.Fill(Ds_Visao, "Cartuchos_Visao");

                    dataGridView_Visao.DataSource = Ds_Visao;
                    dataGridView_Visao.DataMember = "Cartuchos_Visao";

                }
                else
                    BancoDeDados();

            }

        }

        private void button_Novo_Click(object sender, EventArgs e)
        {
            int GridMax = dataGridView_Cartuchos.RowCount;
            Limpar();
            label_Os.Text = Convert.ToString(GridMax);
            button_Processar.Enabled = true;
            button_Processar.Text = "SALVAR";
        }

        private void button1_Click(object sender, EventArgs e)
        {
            dataGridView_Visao.Update();
            dataGridView_Cartuchos.Update();
            BancoDeDados();
            textBox_Pesquisa.Text = string.Empty;
        }

        private void button_AlterarCaminho_Click(object sender, EventArgs e)
        {
            if (MessageBox.Show("Altera Caminho ?", "Aviso", MessageBoxButtons.YesNo, MessageBoxIcon.Question) == DialogResult.Yes)
            {
                try
                {
                    System.Diagnostics.Process.Start(@"C:\Controle-de-Cartuchos\Controle-de-Cartuchos\ArqID.txt");
                }
                catch (Exception Erro)
                {
                    MessageBox.Show("Erro" + Erro);
                }
                CaminhoBancoDados();
                BancoDeDados();
                
            }
        }

        private void dataGridView_Visao_CellDoubleClick(object sender, DataGridViewCellEventArgs e)
        {

        }

        private void dataGridView_Visao_CellContentDoubleClick(object sender, DataGridViewCellEventArgs e)
        {

        }

        private void button_Excluir_Click(object sender, EventArgs e)
        {
            int GridMax = dataGridView_Cartuchos.RowCount - 1;

            if(MessageBox.Show("Deseja Realmente Apagar o Registro ?","Alerta",MessageBoxButtons.YesNo,MessageBoxIcon.Question) == DialogResult.Yes)
            {
            
                if (LinhaAtual >= 0 && LinhaAtual < GridMax)
                {
                    CodigoID = dataGridView_Cartuchos[0, LinhaAtual].Value.ToString();

                    string Conexao = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" + CaminhoBancoDados();

                    string Excluir = "DELETE FROM Cartuchos WHERE OS=" + int.Parse(CodigoID) + "";

                    OleDbConnection dbConexao = new OleDbConnection(Conexao);

                    OleDbCommand cmdExcluir = new OleDbCommand(Excluir, dbConexao);

                    try
                    {
                        dbConexao.Open();
                        cmdExcluir.ExecuteNonQuery();
                    }
                    catch (Exception Erro)
                    {
                        MessageBox.Show("Erro " + Erro);
                    }
                    finally
                    {
                        dbConexao.Close();
                        BancoDeDados();
                        Limpar();
                    }

                }
            }
        }

        private void comboBox_Produto1_SelectedIndexChanged(object sender, EventArgs e)
        {
            Cartucho01();
        }

        private void comboBox_Produto2_SelectedIndexChanged(object sender, EventArgs e)
        {
            Cartucho02();
        }

        private void comboBox_Produto3_SelectedIndexChanged(object sender, EventArgs e)
        {
            Cartucho03();
        }

        private void comboBox_Produto4_SelectedIndexChanged(object sender, EventArgs e)
        {
            Cartucho04();
        }

        private void comboBox_Produto5_SelectedIndexChanged(object sender, EventArgs e)
        {
            Cartucho05();
        }

        private void comboBox_Produto6_SelectedIndexChanged(object sender, EventArgs e)
        {
            Cartucho06();
        }

        private void comboBox_Produto7_SelectedIndexChanged(object sender, EventArgs e)
        {
            Cartucho07();
        }

        private void comboBox_Produto8_SelectedIndexChanged(object sender, EventArgs e)
        {
            Cartucho08();
        }

        private void comboBox_Servico1_SelectedIndexChanged(object sender, EventArgs e)
        {
            Cartucho01();
        }

        private void comboBox_Servico2_SelectedIndexChanged(object sender, EventArgs e)
        {
            Cartucho02();
        }

        private void comboBox_Servico3_SelectedIndexChanged(object sender, EventArgs e)
        {
            Cartucho03();
        }

        private void comboBox_Servico4_SelectedIndexChanged(object sender, EventArgs e)
        {
            Cartucho04();
        }

        private void comboBox_Servico5_SelectedIndexChanged(object sender, EventArgs e)
        {
            Cartucho05();
        }

        private void comboBox_Servico6_SelectedIndexChanged(object sender, EventArgs e)
        {
            Cartucho06();
        }

        private void comboBox_Servico7_SelectedIndexChanged(object sender, EventArgs e)
        {
            Cartucho07();
        }

        private void comboBox_Servico8_SelectedIndexChanged(object sender, EventArgs e)
        {
            Cartucho08();
        }

        private void button_Imprimir_Click(object sender, EventArgs e)
        {
            try
            {
                PrintDocument Imprimir = new PrintDocument();

                Imprimir.PrintPage += new PrintPageEventHandler(this.printDocument_Imprimir_PrintPage);

                Imprimir.Print();
            }
            catch
            {
            }
        }

        private void printDocument_Imprimir_PrintPage(object sender, PrintPageEventArgs e)
        {

            e.Graphics.DrawString("=============================================\n\n     ..::Central     Do     Cartucho::..    \n\n" + "=============================================\n" + "Cliente  : " + textBox_Nome.Text + "\nTelefone : " + textBox_Telefone.Text + "            OS : " + label_Os.Text + "\nData     : " + dateTimePicker_Data.Value.ToShortDateString() + "\n=============================================\nServiços : \n" + "Artigo      " + "Identificação   " + "PSEntrada    " + "Valor\n------------------------------------------------\n" + comboBox_Produto1.Text + "        " + textBox_Identificacao1.Text + "         " + textBox_PsEntrada1.Text + "           " + textBox_Valor1.Text + "\n" + comboBox_Produto2.Text + "        " + textBox_Identificacao2.Text + "         " + textBox_PsEntrada2.Text + "           " + textBox_Valor2.Text + "\n" + comboBox_Produto3.Text + "        " + textBox_Identificacao3.Text + "         " + textBox_PsEntrada3.Text + "           " + textBox_Valor3.Text + "\n" + comboBox_Produto4.Text + "        " + textBox_Identificacao4.Text + "         " + textBox_PsEntrada4.Text + "           " + textBox_Valor4.Text + "\n" + comboBox_Produto5.Text + "        " + textBox_Identificacao5.Text + "         " + textBox_PsEntrada5.Text + "           " + textBox_Valor5.Text + "\n" + comboBox_Produto6.Text + "        " + textBox_Identificacao6.Text + "         " + textBox_PsEntrada6.Text + "           " + textBox_Valor6.Text + "\n" + comboBox_Produto7.Text + "        " + textBox_Identificacao7.Text + "         " + textBox_PsEntrada7.Text + "           " + textBox_Valor7.Text + "\n" + comboBox_Produto8.Text + "        " + textBox_Identificacao8.Text + "         " + textBox_PsEntrada8.Text + "           " + textBox_Valor8.Text + "\n" + "================================= TOTAL : " + textBox_ValorTotal.Text + "\nInformacoes : \n" + "1 - Os cartuchos so poderao ser retirados mediante a apresentacao deste comprovante.\nConserve-o;\n" + "2 - Apos 15 dias se nao forem retirados, os Cartuchos poderao ser revendidos para\ncobrirem gastos e Mao de obra;\n" + "3 - Nao nos responsabilizamos por uso incorreto dos cartuchos, caso haja alguma\nduvida, peca auxilio a um de nossos tecnicos;\n\n" + "     RECARGA   NAO   DANIFICA   CARTUCHOS\n         A CENTRAL DO CARTUCHO AGRADECE \n\n" + "     RUA DORA LIGIA N° 25 - VILA ABERNESIA  " + "                   (12)3662-4150", new Font("Arial", 16), Brushes.Black, 0, 0);

            e.HasMorePages = false;


        }

        private void comboBox_Encerrada_SelectedIndexChanged(object sender, EventArgs e)
        {
            if ((comboBox_Encerrada.Text != "Nao") && (comboBox_Encerrada.ForeColor != Color.Red))
                dateTimePicker_Encerramento.Visible = true;
            else
                dateTimePicker_Encerramento.Visible = false;
        }
    }
}
