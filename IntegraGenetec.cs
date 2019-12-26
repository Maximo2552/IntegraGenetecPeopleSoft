using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
//using System.Data.SqlClient;
using System.Data.OracleClient;
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
using System.Data.SqlClient;
using System.Globalization;

namespace IntegraGenetecPeopleSoft
{

    public partial class IntegraGenetec : Form
    {
        Configuration config = ConfigurationManager.OpenExeConfiguration(ConfigurationUserLevel.None);
        private static string certificado = string.Empty;
        private static string diretorioSC = string.Empty;
        private static string usuarioSC = string.Empty;
        private static string senhaSC = string.Empty;
        private static string caminhoExe = string.Empty;
        private static string caminhoArquivo = string.Empty;
        //private static string gruposPermanentes;
        private List<CardholderGroup> grupos;
        private Engine sdk = new Engine();
        private string inicio = DateTime.Now.ToString();
        private string fim;
        public IntegraGenetec()
        {
            try
            {
                Cursor.Current = Cursors.WaitCursor;

                IniciarLog();
                certificado = ConfigurationManager.AppSettings["Certificado"];
                diretorioSC = ConfigurationManager.AppSettings["Diretorio"].ToString();
                usuarioSC = ConfigurationManager.AppSettings["UsuarioSC"].ToString();
                senhaSC = ConfigurationManager.AppSettings["SenhaSC"].ToString();

                //diretorioSC = "172.16.190.162";
                //usuarioSC = "admin";
                //senhaSC = "G&n&t&c2019";
                //sdk.ClientCertificate = "KxsD11z743Hf5Gq9mv3+5ekxzemlCiUXkTFY5ba1NOGcLCmGstt2n0zYE9NsNimv";
                //sdk.ClientCertificate = "y+BiIiYO5VxBax6/HNi7/ZcXWuvlnEemfaMhoQS1RMkfOGvEBWdUV7zQN272yHVG";                


                sdk.ClientCertificate = certificado;
                try
                {
                    sdk.LogOn(diretorioSC, usuarioSC, senhaSC);
                }
                catch (Exception ex)
                {
                    AppendLog(ex.Message);

                }

                if (sdk.IsConnected)
                {
                    Cursor.Current = Cursors.Default;
                    AppendLog("Logado no SC...");
                    InitializeComponent();
                    Refresh();
                  
                    ExportacaoGenetecSQL();                  
                }
                else
                {
                    Cursor.Current = Cursors.Default;
                    AppendLog("Não foi possível Logar no SC...");                    
                }
                

            }
            catch (Exception ex)
            {
                Cursor.Current = Cursors.Default;
                AppendLog(ex.Message);                
            }

        }

        private void ExportarGenetc()
        {
            int numero_inserido = 0;
            int numero_alterado = 0;
            AppendLog("Inicio da Integração com o Genetc!");
            try
            {
                
                Cursor.Current = Cursors.WaitCursor;
                IniciarLog();

                grupos = RetornarGrupos();
                var colaboradoresGenetec = RetornarColaboradoresGenetec();


                Configuration config = ConfigurationManager.OpenExeConfiguration(ConfigurationUserLevel.None);
                string connectionString = config.ConnectionStrings.ConnectionStrings["Conexao"].ToString();
                
                using (OracleConnection connection = new OracleConnection(connectionString))
                {

                    string queryString = ConfigurationManager.AppSettings["Select"];

                    OracleCommand command = new OracleCommand(queryString, connection);
                    connection.Open();
                    DataTable dt = new DataTable();

                    dt.Load(command.ExecuteReader());
                    using (var reader = command.ExecuteReader())
                    {
                        while (reader.Read())
                        {
                            try
                            {
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
                                        if (status == Status.ATIVO.ToString() || status == Status.ATIVO_AUSENTE.ToString() || status == Status.INATIVO.ToString())
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
                                                AppendLog("Inserido CardHolder: " + nome + " " + sobrenome + " Matricula: " + matricula + " Status: " + status);
                                                cardHolder = sdk.CreateEntity(nome, EntityType.Cardholder) as Cardholder;
                                                cardHolder.FirstName = nome;
                                                cardHolder.LastName = sobrenome;
                                                cardHolder.SetCustomFieldAsync("Matricula", matricula);
                                                cardHolder.Groups.Add(gurpoNovo.Guid);
                                                numero_inserido += 1;
                                            }
                                            else
                                            {
                                                numero_inserido += 1;
                                                if (cardHolder.Groups.Count > 0)
                                                {
                                                    var old_Group = "";
                                                    var cardholdergrupo = new List<Guid>();
                                                    cardholdergrupo.AddRange(cardHolder.Groups.ToList<Guid>());

                                                    foreach (Guid guid in cardholdergrupo)
                                                    {
                                                        var cardHolderGroup = sdk.GetEntity(guid) as CardholderGroup;
                                                        if (!(Boolean)cardHolderGroup.CustomFields["Especial"])
                                                        {
                                                            cardHolder.Groups.Remove(cardHolderGroup.Guid);
                                                        }
                                                    }
                                                    string newSatus = "";
                                                    if (status.ToUpper() == "INATIVO" || status.ToUpper() == "ATIVO_AUSENTE")
                                                    {
                                                        newSatus = "Inactive";
                                                    }
                                                    else
                                                    {
                                                        newSatus = "Active";
                                                    }
                                                    if (old_Group != gurpoNovo.Name)
                                                    {
                                                        if (cardHolder.State.ToString() == newSatus)
                                                        {
                                                            AppendLog("Alterado CardHolder: " + nome + " " + sobrenome + " Matricula: " + matricula + " Grupo Anterior:" + old_Group + " Para Novo Grupo: " + gurpoNovo.Name);
                                                        }
                                                        else
                                                        {
                                                            AppendLog("Alterado CardHolder: " + nome + " " + sobrenome + " Matricula: " + matricula + " Grupo Anterior:" + old_Group + " Para Novo Grupo: " + gurpoNovo.Name + " Status: " + status);
                                                        }

                                                    }
                                                    else
                                                    {
                                                        if (cardHolder.State.ToString() == newSatus)
                                                        {
                                                            AppendLog("Alterado CardHolder: " + nome + " " + sobrenome + " Matricula: " + matricula + " Status: " + status);
                                                        }
                                                    }
                                                }
                                                cardHolder.Groups.Add(gurpoNovo.Guid);
                                            }
                                            cardHolder.State = ObterStatus(status);
                                        }
                                    }

                                }
                                sdk.TransactionManager.CommitTransaction();
                            }
                            catch (Exception ex)
                            {
                                AppendLog(ex.Message);
                                sdk.TransactionManager.RollbackTransaction();
                                AppendLog("Erro no CardHolder: Grupo " + reader[4].ToString() + " " + reader[0].ToString() + " " + reader[1].ToString() + " " + reader[2].ToString() + " Matrícula: " + reader[3].ToString());
                            }

                        }
                        
                    }
                }
                Cursor.Current = Cursors.Default;
                AppendLog("__________________________________________________________________");
                AppendLog("Fim de Integração..." + inicio + " Fim: " + DateTime.Now.ToString());
                AppendLog("Toral de Registros Inseridos( " + numero_inserido + ") Alerados ( " + numero_alterado + " )");

                Close();
            }
            catch (OracleException sqlex)
            {
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
        private void ExportacaoGenetecSQL()
        {
            int numero_inserido = 0;
            int numero_alterado = 0;
            AppendLog("Inicio da Integração com o Genetc!");
            try
            {
                Cursor.Current = Cursors.WaitCursor;
                IniciarLog();
                progressBar1.Minimum = 1;

                grupos = RetornarGrupos();                
                var colaboradoresGenetec = RetornarColaboradoresGenetec();
                
                string gurposEspeciais = ConfigurationManager.AppSettings["Grupos"];
                var gruposPermanentes = gurposEspeciais.Split(',').ToList<string>();

                Configuration config = ConfigurationManager.OpenExeConfiguration(ConfigurationUserLevel.None);
                string connectionString = config.ConnectionStrings.ConnectionStrings["Conexao"].ToString();
                
                //connectionString = "Data Source=GCTEC04;Initial Catalog=Integracao;User ID=imod;Password=imod;Min Pool Size=5;Max Pool Size=15;Connection Reset=True;Connection Lifetime=600;Trusted_Connection=no;MultipleActiveResultSets=True";
                //AppendLog(connectionString);

                using (SqlConnection connection = new SqlConnection(connectionString))
                {

                    string queryString = ConfigurationManager.AppSettings["Select"];

                    SqlCommand command = new SqlCommand(queryString, connection);
                    //AppendLog(queryString);
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

                                   
                                    if (cardHolder == null || TeveAlteracaoSQL(cardHolder, reader))
                                    {
                                        
                                        if (status == Status.ATIVO.ToString() || status == Status.ATIVO_AUSENTE.ToString() || status == Status.INATIVO.ToString())
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
                                                AppendLog("Inserido CardHolder: " + nome + " " + sobrenome + " Matricula: " + matricula + " Status: " + status);
                                                cardHolder = sdk.CreateEntity(nome, EntityType.Cardholder) as Cardholder;
                                                cardHolder.FirstName = nome;
                                                cardHolder.LastName = sobrenome;
                                                cardHolder.SetCustomFieldAsync("Matricula", matricula);
                                                cardHolder.Groups.Add(gurpoNovo.Guid);
                                                numero_inserido += 1;
                                            }
                                            else
                                            {
                                                numero_alterado += 1;
                                                if (cardHolder.Groups.Count > 0)
                                                {
                                                    var old_Group = "";
                                                    var cardholdergrupo = new List<Guid>();
                                                    cardholdergrupo.AddRange(cardHolder.Groups.ToList<Guid>());

                                                    foreach (Guid guid in cardholdergrupo)
                                                    {
                                                        var cardHolderGroup = sdk.GetEntity(guid) as CardholderGroup;
                                                        old_Group = cardHolderGroup.Name;
                                                        if (!(Boolean)cardHolderGroup.CustomFields["Especial"])
                                                        {
                                                            cardHolder.Groups.Remove(cardHolderGroup.Guid);
                                                        }
                                                    }
                                                    string newSatus = "";
                                                    if (status.ToUpper() == "INATIVO" || status.ToUpper() == "ATIVO_AUSENTE")
                                                    {
                                                        newSatus = "Inactive";
                                                    }
                                                    else 
                                                    {
                                                        newSatus = "Active";
                                                    }
                                                    if(old_Group != gurpoNovo.Name)
                                                    {
                                                        if(cardHolder.State.ToString() == newSatus)
                                                        {
                                                            AppendLog("Alterado CardHolder: " + nome + " " + sobrenome + " Matricula: " + matricula + " Grupo Anterior:" + old_Group + " Para Novo Grupo: " + gurpoNovo.Name);
                                                        }
                                                        else
                                                        {
                                                            AppendLog("Alterado CardHolder: " + nome + " " + sobrenome + " Matricula: " + matricula + " Grupo Anterior:" + old_Group + " Para Novo Grupo: " + gurpoNovo.Name + " Status: " + status);
                                                        }
                                                        
                                                    }
                                                    else
                                                    {
                                                        if (cardHolder.State.ToString() == newSatus)
                                                        {
                                                            AppendLog("Alterado CardHolder: " + nome + " " + sobrenome + " Matricula: " + matricula + " Status: " + status);
                                                        }
                                                    }
                                                }
                                                cardHolder.Groups.Add(gurpoNovo.Guid);
                                            }
                                            cardHolder.State = ObterStatus(status);
                                            
                                        }
                                        
                                    }

                                }
                                
                                sdk.TransactionManager.CommitTransaction();
                            }
                            catch (Exception ex)
                            {
                                AppendLog(ex.Message);
                                sdk.TransactionManager.RollbackTransaction();
                                AppendLog("Erro no CardHolder: Grupo " + reader[4].ToString() + " " + reader[0].ToString() + " " + reader[1].ToString() + " " + reader[2].ToString() + " Matrícula: " + reader[3].ToString());
                            }

                        }

                    }
                }
                Cursor.Current = Cursors.Default;
                AppendLog("__________________________________________________________________");
                AppendLog("Fim de Integração..." + inicio + " Fim: " + DateTime.Now.ToString());
                AppendLog("Toral de Registros Inseridos( " + numero_inserido + ") Alerados ( " + numero_alterado + " )");

                Close();
            }
            catch (OracleException sqlex)
            {
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
        private bool TeveAlteracao(Cardholder cardHolder, OracleDataReader reader)
        {
            if (cardHolder == null) return false;
            var grupo = cardHolder.Groups[0];
            var CardholderGrup = grupos.Where(g => g.Guid == grupo).FirstOrDefault();


            return (CardholderGrup == null || cardHolder.State != ObterStatus(reader[5].ToString()) || CardholderGrup.Name != reader[4].ToString());
        }

        private bool TeveAlteracaoSQL(Cardholder cardHolder, SqlDataReader reader)
        {
            try
            {
                if (cardHolder == null) return false;
                var grupo = cardHolder.Groups[0];
                var CardholderGrup = grupos.Where(g => g.Guid == grupo).FirstOrDefault();


                return (CardholderGrup == null || cardHolder.State != ObterStatus(reader[5].ToString()) || CardholderGrup.Name != reader[4].ToString());

            }
            catch (Exception ex)
            {

                throw;
            }
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
                    //String semacento = new string(grupocradholder.Name.Normalize(NormalizationForm.FormD).Where(ch => char.GetUnicodeCategory(ch) != UnicodeCategory.NonSpacingMark).ToArray());
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
                foreach (DataRow dr in result.Data.Rows)
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
        private string CriarGruposEspeciais(string meusgrupos)
        {
            string gurposEspeciais = ConfigurationManager.AppSettings["Grupos"];
            var gruposPermanentes = gurposEspeciais.Split(',').ToList<string>();
            //grupos = gurposEspeciais.Split(',').ToList<string>();
            foreach (string nomeGrupo in gruposPermanentes)
            {
                var gurpoNovo = grupos.Where(g => g.Name == nomeGrupo).FirstOrDefault();
                if (gurpoNovo == null)
                {
                    gurpoNovo = sdk.CreateEntity(nomeGrupo, EntityType.CardholderGroup) as CardholderGroup;
                    grupos.Add(gurpoNovo);
                }
            }
            return gruposPermanentes.ToString();
        }
        //private string CriarGruposEspeciais2<Strint>(List<String> EspeciaiGrupos)
        //{
        //    foreach (string nomeGrupo in EspeciaiGrupos)
        //    {
        //        var gurpoNovo = grupos.Where(g => g.Name == nomeGrupo).FirstOrDefault();
        //        if (gurpoNovo == null)
        //        {
        //            gurpoNovo = sdk.CreateEntity(nomeGrupo, EntityType.CardholderGroup) as CardholderGroup;
        //            grupos.Add(gurpoNovo);
        //        }
        //    }

        //    return EspeciaiGrupos.ToString();
        //}

        private void Timer1_Tick(object sender, EventArgs e)
        {
            AppendLog("Timer1_Tick");
            //ExportarGenetc();
        }
    }
}


