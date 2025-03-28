namespace FrmInsercao.Utils
{
    public class DtgManager
    {
        public void VeriricarAlteracaoDtg(int _linesDtgInicial, int linesAtual)
        {
            if (linesAtual != _linesDtgInicial)
            {
                var respostaUser = MessageBox.Show("Deseja salvar as alterações feitas na tabela?",
                    "CONFIRMAÇÃO", MessageBoxButtons.YesNo, MessageBoxIcon.Question);

                if (respostaUser == DialogResult.Yes) { }
                //SalvarAlteracoesTable();
            }
        }
    }
}
