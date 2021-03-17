// mover mensagem apos desanexar
// 

using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using Pop3;
using System.Collections;
using System.Configuration;
using System.IO;

namespace EmailNFE
{
    public partial class Form1 : Form
    {
        public bool fecharForm = false;

        public Form1()
        {
            InitializeComponent();
            txtEmail.Text = ConfigurationSettings.AppSettings["email"];
            txtSenha.Text = ConfigurationSettings.AppSettings["senha"];
            txtServidorPop.Text = ConfigurationSettings.AppSettings["servidorPop"];
            txtDiretorio.Text = ConfigurationSettings.AppSettings["diretorio"];

            if (!Directory.Exists(txtDiretorio.Text))
                Directory.CreateDirectory(txtDiretorio.Text);
        }

        private void button1_Click(object sender, EventArgs e)
        {
            Receber();
        }

        public void MostrarTela()
        {
            Show();
            WindowState = FormWindowState.Normal;
        }

        private void btnFechar_Click(object sender, EventArgs e)
        {
            if (MessageBox.Show("Deseja realmente fechar o programa ?", "Atenção", MessageBoxButtons.YesNo, MessageBoxIcon.Question, MessageBoxDefaultButton.Button2) == DialogResult.Yes)
            {
                fecharForm = true;
                Close();
            }
        }

        private void Form1_Shown(object sender, EventArgs e)
        {
            if (WindowState == FormWindowState.Minimized)
            {
                Hide();
            }
        }

        private void Form1_Resize(object sender, EventArgs e)
        {
            notifyIcon1.Text = this.Text;
            notifyIcon1.BalloonTipTitle = this.Text;
            notifyIcon1.BalloonTipText = "Clique duas vezes no ícone para retornar ao programa de e-mail NF-e!";
            notifyIcon1.ShowBalloonTip(0);

            if (WindowState == FormWindowState.Minimized)
            {
                Hide();
            }
        }

        private void Form1_FormClosing(object sender, FormClosingEventArgs e)
        {
            if (!fecharForm)
            {
                e.Cancel = true;
                this.Visible = false;
                this.OnResize(e);
                notifyIcon1.ShowBalloonTip(6000);
            }
            else
            {
                timer1.Enabled = false;
            }
        }

        private void abrirToolStripMenuItem_Click(object sender, EventArgs e)
        {
            MostrarTela();
        }

        private void btnSalvar_Click(object sender, EventArgs e)
        {
            GravarConfiguracao("email", txtEmail.Text);
            GravarConfiguracao("senha", txtSenha.Text);
            GravarConfiguracao("servidorPop", txtServidorPop.Text);
            GravarConfiguracao("diretorio", txtDiretorio.Text);
        }

        public static void GravarConfiguracao(string chave, string valor)
        {
            Configuration config = ConfigurationManager.OpenExeConfiguration(ConfigurationUserLevel.None);
            config.AppSettings.Settings.Remove(chave);
            config.Save();
            config.AppSettings.Settings.Add(chave, valor);
            config.Save();
            ConfigurationManager.RefreshSection("appSettings");
        }

        private void notifyIcon1_DoubleClick(object sender, EventArgs e)
        {
            MostrarTela();
        }

        private void timer1_Tick(object sender, EventArgs e)
        {
            Receber();
        }


        public void Receber()
        {

            string erro = string.Empty;

            Refresh();
            timer1.Enabled = false;
            try
            {
                Pop3Client email = new Pop3Client(txtEmail.Text, txtSenha.Text, txtServidorPop.Text);
                email.OpenInbox();

                try
                {
                    while (email.NextEmail())
                    {
                        if (email.IsMultipart)
                        {
                            IEnumerator enumerator = email.MultipartEnumerator;
                            while (enumerator.MoveNext())
                            {
                                Pop3Component multipart = (Pop3Component)enumerator.Current;

                                try
                                {
                                    //if (email.DeleteEmail())
                                    //{
                                    //    Grava_Log(">> OK - EXCLUIDO (01)- " + email.Subject);
                                    //    txtExcluidos.Text += email.Subject + "\r\n";
                                    //}
                                }
                                catch(Exception ex)
                                {
                                    erro = ">> ERRO AO RECEBER O E-MAIL "+ex.Message.ToString();
                                    Grava_Log(erro);

                                    txtStatus.Text += erro + "\r\n";

                                    //if (email.DeleteEmail())
                                    //{
                                    //    Grava_Log(">> ERRO - EXCLUIDO (01)- " + email.Subject);
                                    //    txtExcluidos.Text += email.Subject + "\r\n";
                                    //}

                                }
                            }
                        }
                        else
                        {
                            //if (email.DeleteEmail())
                            //{
                            //    Grava_Log(">> OK - EXCLUIDO (02) - "+email.Subject);
                            //    txtExcluidos.Text += "" + email.Subject + "\r\n";
                            //}
                        }


                    }
                    
                }
                finally
                {
                    Grava_Log(email.Subject);

                    email.CloseConnection();
                    Receber();
                }


            }
            catch (Pop3LoginException)
            {
                erro = ">> ERRO: problemas ao efetuar login, verifique se o e-mail e a senha estão corretos!";
                Grava_Log(erro);

                txtStatus.Text += erro + "\r\n";
            }
            catch (Exception ex)
            {
                erro = ">> ERRO (03): " + ex.Message.ToString();
                Grava_Log(erro);

                txtStatus.Text += erro + "\r\n";
                
            }
            finally
            {
                timer1.Enabled = true;

            }
        }

        private void button2_Click(object sender, EventArgs e)
        {
            openFileDialog1.ShowDialog();
            string arquivo = openFileDialog1.FileName;
            try
            {
                if (File.Exists(arquivo))
                {
                    //create the Dataset object
                    using (DataSet ds = new DataSet())
                    {
                        //load the xml data to the dataset
                        ds.ReadXml(arquivo);
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("Erro ao ler o XML - " + ex.Message.ToString());
            }
        }

        public void Grava_Log(string conteudo)
        {
            string arquivoLog = Directory.GetCurrentDirectory()+ "\\log.txt";
            StreamWriter arquivoGerado = new StreamWriter(arquivoLog, true);

            arquivoGerado.WriteLine(conteudo);
            arquivoGerado.Close();

         }
    }
}
