using FrmInsercao.Utils;

namespace FrmInsercao
{
    internal static class Program
    {
        static void Main()
        {
            ConfigrarAplicationSQLite();
            ApplicationConfiguration.Initialize();
            Application.Run(new FrmInsercao());
        }

        private static void ConfigrarAplicationSQLite()
        {
            SqliteManager.CriarBancoDados();
            SqliteManager.CriarTabela();
        }
    }
}