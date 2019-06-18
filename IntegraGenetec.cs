using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Data.SqlClient;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Genetec.Sdk;
using Genetec.Sdk.Entities;
using Genetec.Sdk.Queries;
using Genetec.Sdk.Scripting;
using System.Configuration;
using System.IO;
using System.Reflection;

namespace IntegraGenetecPeopleSoft
{

    public partial class IntegraGenetec : Form
    {
        private static string caminhoExe = string.Empty;
        private static string caminhoArquivo = string.Empty;
        private List<CardholderGroup> grupos;
        private Engine sdk = new Engine();
        private string inicio = DateTime.Now.ToString();
        private string fim;
        public IntegraGenetec()
        {
            sdk.ClientCertificate = "KxsD11z743Hf5Gq9mv3+5ekxzemlCiUXkTFY5ba1NOGcLCmGstt2n0zYE9NsNimv";
            sdk.LogOn("172.16.190.108", "admin", "");
            InitializeComponent();
            Refresh();
            //ExportarGenetc();
            timer1.Enabled = true;
        }

        private void ExportarGenetc()
        {

            try
            {
                Cursor.Current = Cursors.WaitCursor;
                IniciarLog();

                progressBar1.Minimum = 1;
               
                grupos = RetornarGrupos();
                var colaboradoresGenetec = RetornarColaboradoresGenetec();
                var gruposPermanentes = new List<string>() { "DATA CENTER", "GESEC MONITORAMENTO", "GESEC COGIL", "DIRETORIA", "HALL DIRETORIA", "PRESIDÊNCIA", "15 ANDAR B", "SITE BACKUP", "SB DATA CENTER", "SB SALA", "VIGILANTES", "BOMBEIROS CIVIL", "SERVIÇOS GERAIS", "CATAVENTO", "HELP DESK", "GESTORES", "TEMPORÁRIO", "JANELA OPERACIONAL" };


                Configuration config = ConfigurationManager.OpenExeConfiguration(ConfigurationUserLevel.None);
                string connectionString = config.ConnectionStrings.ConnectionStrings["Conexao"].ToString();

                using (var connection = new SqlConnection(connectionString))
                {

                    string queryString = ConfigurationManager.AppSettings["Select"];                    

                    var command = new SqlCommand(queryString, connection);
                    connection.Open();
                    //sdk.TransactionManager.CreateTransaction();
                    DataTable dt = new DataTable();

                    dt.Load(command.ExecuteReader());
                    progressBar1.Maximum = dt.Rows.Count;
                    using (var reader = command.ExecuteReader())
                    {
                        while (reader.Read())
                        {
                            try
                            {
                                //aqui incluo o método necessário para continuar o trabalho 
                                sdk.TransactionManager.CreateTransaction();

                                if (!string.IsNullOrEmpty(reader[5].ToString()) && !string.IsNullOrEmpty(reader[3].ToString()) && !string.IsNullOrEmpty(reader[4].ToString()))
                                {
                                    string nome = reader[0].ToString();
                                    string sobrenome = reader[1].ToString() + " " + reader[2].ToString();
                                    string status = reader[5].ToString();
                                    string matricula = reader[3].ToString();

                                    Cardholder cardHolder = colaboradoresGenetec.Where(c => c.CustomFields["Matricula"].ToString() == matricula).FirstOrDefault();
                                    if (cardHolder == null || TeveAlteracao(cardHolder, reader))
                                    {
                                        if (status == Status.ATIVO.ToString() || status == Status.ATIVO_AUSENTE.ToString())
                                        {
                                            string nomeGrupo = reader[4].ToString();
                                            var gurpoNovo = grupos.Where(g => g.Name == nomeGrupo).FirstOrDefault();

                                            if (gurpoNovo == null)
                                            {
                                                gurpoNovo = sdk.CreateEntity(nomeGrupo, EntityType.CardholderGroup) as CardholderGroup;
                                                grupos.Add(gurpoNovo);
                                            }

                                            if (cardHolder == null)
                                            {
                                                AppendLog("Inserido CardHolder: " + nome + "" + sobrenome + " Matrícula: " + matricula);
                                                cardHolder = sdk.CreateEntity(nome, EntityType.Cardholder) as Cardholder;
                                                cardHolder.FirstName = nome;
                                                cardHolder.LastName = sobrenome;
                                                cardHolder.SetCustomFieldAsync("Matricula", matricula);
                                                cardHolder.Groups.Add(gurpoNovo.Guid);
                                            }
                                            else
                                            {
                                                if (cardHolder.Groups.Count > 0)
                                                {
                                                    var cardholdergrupo = new List<Guid>();
                                                    cardholdergrupo.AddRange(cardHolder.Groups.ToList<Guid>());

                                                    foreach (Guid guid in cardholdergrupo)
                                                    {
                                                        var cardHolderGroup = sdk.GetEntity(guid) as CardholderGroup;
                                                        if (!gruposPermanentes.Contains(cardHolderGroup.Name))
                                                        {
                                                            cardHolder.Groups.Remove(cardHolderGroup.Guid);
                                                        }
                                                    }
                                                    AppendLog("Alterado CardHolder: " + nome + "" + sobrenome + " Matrícula: " + matricula);
                                                }
                                                cardHolder.Groups.Add(gurpoNovo.Guid);
                                            }
                                            cardHolder.State = ObterStatus(status);
                                        }
                                        else if (status == Status.INATIVO.ToString())
                                        {
                                            sdk.DeleteEntity(cardHolder.Guid);
                                            AppendLog("Excluido CardHolder: " + nome + "" + sobrenome + " Matrícula: " + matricula);
                                        }
                                    }
                                    
                                }
                                progressBar1.Value += 1;
                                sdk.TransactionManager.CommitTransaction();
                            }
                            catch (Exception)
                            {
                                sdk.TransactionManager.RollbackTransaction();
                                AppendLog("Erro no CardHolder: " + reader[0].ToString() + "" + reader[1].ToString() + " " + reader[2].ToString() + " Matrícula: " + reader[3].ToString());
                            }

                        }
                    }
                }
                Cursor.Current = Cursors.Default;
                AppendLog("Fim de Integração..." + inicio + " Fim: " + DateTime.Now.ToString());
                AppendLog("Toral de Registros: " + progressBar1.Value);

                Close();
            }
            catch (SqlException sqlex){
                Cursor.Current = Cursors.Default;
                AppendLog(sqlex.Message);
                Close();
            }
            catch (Exception ex)
            {
                Cursor.Current = Cursors.Default;
                AppendLog(ex.Message);
                Close();
            }
        }
        /// <summary>
        /// Enum dos Status do CardHolder
        /// </summary>
        enum Status
        {
            ATIVO = 1,
            ATIVO_AUSENTE = 2,
            INATIVO = 3,
        }
        private bool TeveAlteracao(Cardholder cardHolder, SqlDataReader reader)
        {
           
            var grupo = cardHolder.Groups[0];
            var CardholderGrup = grupos.Where(g => g.Guid == grupo).FirstOrDefault();


            return (CardholderGrup == null || cardHolder.State != ObterStatus(reader[5].ToString()) || CardholderGrup.Name != reader[4].ToString());
        }

        private CardholderState ObterStatus(string status)
        {
            try
            {
                return (status == "ATIVO") ? CardholderState.Active : CardholderState.Inactive;
            }
            catch (Exception)
            {

                throw;
            }

        }
        /// <summary>
        /// Lista os Grupos de Titulares de Cartão existentes no SC
        /// </summary>
        /// <returns></returns>
        private List<CardholderGroup> RetornarGrupos()
        {
            EntityConfigurationQuery query;
            QueryCompletedEventArgs result;
            List<CardholderGroup> groupos = new List<CardholderGroup>();

            query = sdk.ReportManager.CreateReportQuery(ReportType.EntityConfiguration) as EntityConfigurationQuery;
            query.EntityTypeFilter.Add(EntityType.CardholderGroup);
            query.NameSearchMode = StringSearchMode.StartsWith;
            result = query.Query();
            SystemConfiguration systemConfiguration = sdk.GetEntity(SdkGuids.SystemConfiguration) as SystemConfiguration;
            var service = systemConfiguration.CustomFieldService;
            if (result.Success)
            {
                foreach (DataRow dr in result.Data.Rows)    //sempre remove todas as regras de um CardHolder
                {
                    CardholderGroup grupocradholder = sdk.GetEntity((Guid)dr[0]) as CardholderGroup;
                    groupos.Add(grupocradholder);
                }
            }
            return groupos;
        }
        /// <summary>
        /// Lista todoas os CardHolder existente no SC
        /// </summary>
        /// <returns></returns>
        private List<Cardholder> RetornarColaboradoresGenetec()
        {
            EntityConfigurationQuery query;
            QueryCompletedEventArgs result;
            List<Cardholder> colaboradores = new List<Cardholder>();

            query = sdk.ReportManager.CreateReportQuery(ReportType.EntityConfiguration) as EntityConfigurationQuery;
            query.EntityTypeFilter.Add(EntityType.Cardholder);
            query.NameSearchMode = StringSearchMode.StartsWith;
            result = query.Query();
            SystemConfiguration systemConfiguration = sdk.GetEntity(SdkGuids.SystemConfiguration) as SystemConfiguration;
            var service = systemConfiguration.CustomFieldService;
            if (result.Success)
            {
                foreach (DataRow dr in result.Data.Rows)    //sempre remove todas as regras de um CardHolder
                {
                    Cardholder cardholder = sdk.GetEntity((Guid)dr[0]) as Cardholder;
                    colaboradores.Add(cardholder);
                }
            }
            return colaboradores;
        }
        private static void IniciarLog()
        {
            caminhoExe = Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location);
            caminhoArquivo = Path.Combine(caminhoExe, "Log_Integracao.txt");
            AppendLog("_____________________________________________________________");
            AppendLog("Inicio da carga..." + DateTime.Now.ToString());
        }
        private static void AppendLog(string logMensagem)
        {
            try
            {
                using (StreamWriter txtWriter = System.IO.File.AppendText(caminhoArquivo))
                {
   
                    txtWriter.Write($"{DateTime.Now.ToLongTimeString()} {DateTime.Now.ToLongDateString()}");
                    txtWriter.WriteLine($"  :{logMensagem}");

                }

            }
            catch (Exception ex)
            {
                throw ex;
            }
        }

        private void Timer1_Tick(object sender, EventArgs e)
        {
            ExportarGenetc();
        }
    }
}


