namespace FrmInsercao.Utils
{
    public class DtgManager
    {
        public bool VeriricarAlteracaoDtg(int _linesDtgInicial, int linesAtual)
        {
            bool sucesso = true;

            if (linesAtual != _linesDtgInicial)
            {
                var respostaUser = MessageBox.Show("Deseja salvar as alterações feitas na tabela?",
                    "CONFIRMAÇÃO", MessageBoxButtons.YesNo, MessageBoxIcon.Question);

                if (respostaUser == DialogResult.Yes)
                    sucesso = SqliteManager.SalvarAlteracoesTable();
            }

            return sucesso;
        }
    }
}
