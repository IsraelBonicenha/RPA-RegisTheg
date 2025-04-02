using System.Configuration;
using System.Data;
using System.Data.SQLite;
using System.Text;

namespace FrmInsercao.Utils
{
    class SqliteManager
    {
        private static SQLiteConnection GetConnection()
        {
            return new SQLiteConnection(ConfigurationManager.ConnectionStrings["SQLiteConnection"].ConnectionString);
        }

        public static void CriarBancoDados()
        {
            try
            {
                if (!Directory.Exists("data"))
                    Directory.CreateDirectory("data");

                if (!File.Exists("data/base.sqlite"))
                    SQLiteConnection.CreateFile("data/base.sqlite");

                
            }
            catch (Exception ex) { MessageBox.Show($"Erro ao criar banco de dados: {ex.Message}"); }
        }

        public static void CriarTabela()
        {
            using (var connection = GetConnection())
            {
                try
                {
                    connection.Open();
                    using (var command = connection.CreateCommand())
                    {
                        command.CommandText = @"CREATE TABLE IF NOT EXISTS principal(
                                                p_id TEXT, 
                                                p_hash TEXT NOT NULL, 
                                                p_valores TEXT, 
                                                p_criterio TEXT NOT NULL,
                                                p_status TEXT NOT NULL,
                                                p_dataCadastro DATE
                                                );
                        ";

                        command.ExecuteNonQuery();
                    }
                }
                catch (Exception ex){ MessageBox.Show($"Erro ao criar tabela: {ex.Message}"); }
            }
        }

        public static bool SalvarAlteracoesTable()
        {
            ClearTableBancoDados();

            using (var connection = GetConnection())
            {
                try
                {
                    connection.Open();

                    bool primeiraLinha = true;
                    StringBuilder queryBuild = new StringBuilder();
                    queryBuild.AppendLine(@"INSERT INTO principal (p_id, p_hash, p_valores, p_criterio, p_status, p_dataCadastro) VALUES");
                    foreach (DataGridViewRow row in FrmInsercao.dtg.Rows)
                    {

                        string id = row.Cells[0].Value?.ToString().ToUpper() ?? "";
                        string hash = row.Cells[1].Value?.ToString().ToUpper() ?? "";
                        string valores = row.Cells[2].Value?.ToString().ToUpper() ?? "";
                        string criterio = row.Cells[3].Value?.ToString().ToUpper() ?? "";
                        string status = row.Cells[4].Value?.ToString().ToUpper() ?? "";
                        var data = DateTime.Now.Date;

                        if (row.IsNewRow || status == "E")
                            continue;

                        if (!Guid.TryParse(id, out Guid _))
                            id = Guid.NewGuid().ToString();

                        // Validação para o campo 'hash'
                        if (string.IsNullOrEmpty(hash))
                        {
                            MessageBox.Show("Erro: Hash não pode ser vazio", "Erro", MessageBoxButtons.OK);
                            return false;  // Interrompe o processo
                        }

                        // Validação para o campo 'valores'
                        if (string.IsNullOrEmpty(valores))
                        {
                            MessageBox.Show("Erro: Valores não podem ser vazios", "Erro", MessageBoxButtons.OK);
                            return false;  // Interrompe o processo
                        }

                        // Validação para o campo 'criterio'
                        if (string.IsNullOrEmpty(criterio))
                        {
                            MessageBox.Show("Erro: Critério não pode ser vazio", "Erro", MessageBoxButtons.OK);
                            return false;  // Interrompe o processo
                        }

                        // Validação para o campo 'status'
                        if (string.IsNullOrEmpty(status))
                        {
                            MessageBox.Show("Erro: Status não pode ser vazio", "Erro", MessageBoxButtons.OK);
                            return false;  // Interrompe o processo
                        }

                        if (!primeiraLinha)
                            queryBuild.AppendLine(",");

                        queryBuild.AppendLine($"('{id}', '{hash}', '{valores}', '{criterio}', '{status}', '{data:yyyy-MM-dd}')");

                        primeiraLinha = false;
                    }


                    string query = queryBuild.ToString();

                    using (var command = connection.CreateCommand())
                    {
                        command.CommandText = query;
                        command.ExecuteNonQuery();
                    }

                    return true;
                }
                catch (Exception ex) { MessageBox.Show($"Erro ao salvar alterações na tabela: {ex.Message}"); return false; }
            };
        }

        private static void ClearTableBancoDados()
        {
            using (var connection = GetConnection())
            {
                try
                {
                    connection.Open();
                    using (var command = connection.CreateCommand())
                    {
                        command.CommandText = @"DELETE FROM principal;";

                        command.ExecuteNonQuery();
                    }
                }
                catch (Exception ex) { MessageBox.Show($"Erro ao limpar tabela: {ex.Message}"); }
            }
        }

        public static void CarregarDadosGrid()
        {
            DataGridView dtg = FrmInsercao.dtg;
            using (var connection = GetConnection())
            {
                DataTable dt = new DataTable();
                try
                {
                    connection.Open();

                    using (var command = connection.CreateCommand())
                    {
                        command.CommandText = "SELECT * FROM principal";
                        SQLiteDataAdapter da = new SQLiteDataAdapter(command);
                        da.Fill(dt);

                        dtg.Columns["colID"].DataPropertyName = "p_id";
                        dtg.Columns["ColHashtag"].DataPropertyName = "p_hash";
                        dtg.Columns["ColValor"].DataPropertyName = "P_valores";
                        dtg.Columns["ColCriterio"].DataPropertyName = "p_criterio";
                        dtg.Columns["colStatus"].DataPropertyName = "p_status";

                        dtg.DataSource = dt;
                    }
                }
                catch (Exception ex)
                {
                    MessageBox.Show($"Erro ao carregar a tabela: {ex.Message}");
                }
            }
        }

    }
}
