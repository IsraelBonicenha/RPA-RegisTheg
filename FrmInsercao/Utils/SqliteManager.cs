using System.Configuration;
using System.Data.SQLite;

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
                                                P_valores TEXT, 
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

        // TERMINAR ESSE MÉTODO
        public static void SalvarAlteracoesTable()
        {
            ClearTableBancoDados();

            using (var connection = GetConnection())
            {
                try
                {
                    connection.Open();
                }
                catch (Exception ex) { MessageBox.Show($"Erro ao salvar alterações na tabela: {ex.Message}"); }
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
    }
}
