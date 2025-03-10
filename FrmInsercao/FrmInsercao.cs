namespace FrmInsercao
{
    public partial class FrmInsercao : Form
    {
        public FrmInsercao()
        {
            InitializeComponent();
            dtgPrincipal.AllowUserToAddRows = true;
            // O comando acima foi para a correção de um bug
        }

        private void FrmInsercao_Load(object sender, EventArgs e)
        {
            // Aviso de Termo de Uso (talvez eu personalize a mensagem com outra tela futuramente.......)
            DialogResult respostaUser = MessageBox.Show(
                "Este software é apenas uma demonstração baseada no sistema da empresa onde trabalho. " +
                "Não há intenção de vazar dados sensíveis, e qualquer uso indevido não será de minha r" +
                "esponsabilidade. Ao continuar, você concorda em utilizar este software apenas para fi" +
                "ns de teste.\n\nDeseja continuar?",
                "Termo de Uso",
                MessageBoxButtons.YesNo,
                MessageBoxIcon.Warning
            );

            if (respostaUser == DialogResult.No)
            {
                Application.Exit();
            }
        }
    }
}
