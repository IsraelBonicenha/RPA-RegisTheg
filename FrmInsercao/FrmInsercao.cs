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

        #region btnClose Events
        private void btnClose_MouseEnter(object sender, EventArgs e)
        {
            btnClose.BackColor = Color.Firebrick;
        }

        private void btnClose_MouseLeave(object sender, EventArgs e)
        {
            btnClose.BackColor = Color.White;
        }

        private void btnClose_MouseDown(object sender, MouseEventArgs e)
        {
            btnClose.BackColor = Color.IndianRed;
        }

        private void btnClose_MouseUp(object sender, MouseEventArgs e)
        {
            btnClose.BackColor = Color.Firebrick;
        }

        private void btnClose_Click(object sender, EventArgs e)
        {
            Application.Exit();
        }
        #endregion

        #region btnMaximize Events
        private void btnMaximize_MouseEnter(object sender, EventArgs e)
        {
            btnMaximize.BackColor = Color.SteelBlue;
        }

        private void btnMaximize_MouseLeave(object sender, EventArgs e)
        {
            btnMaximize.BackColor = Color.White;
        }

        private void btnMaximize_MouseDown(object sender, MouseEventArgs e)
        {
            btnMaximize.BackColor = Color.CornflowerBlue;
        }

        private void btnMaximize_MouseUp(object sender, MouseEventArgs e)
        {
            btnMaximize.BackColor = Color.WhiteSmoke;
        }

        // Preciso corrigir a responsividade do form agora... Está horrível
        private void btnMaximize_Click(object sender, EventArgs e)
        {
            bool stateWindow = this.WindowState == FormWindowState.Normal;

            if (stateWindow)
                this.WindowState = FormWindowState.Maximized;
            else
                this.WindowState = FormWindowState.Normal;
        }
        #endregion

        #region btnMinimize Events
        private void btnMinimize_MouseEnter(object sender, EventArgs e)
        {
            btnMinimize.BackColor = Color.SteelBlue;
        }

        private void btnMinimize_MouseLeave(object sender, EventArgs e)
        {
            btnMinimize.BackColor = Color.White;
        }

        private void btnMinimize_MouseDown(object sender, MouseEventArgs e)
        {
            btnMinimize.BackColor = Color.CornflowerBlue;
        }

        private void btnMinimize_MouseUp(object sender, MouseEventArgs e)
        {
            btnMinimize.BackColor = Color.WhiteSmoke;
        }

        private void btnMinimize_Click(object sender, EventArgs e)
        {
            this.WindowState = FormWindowState.Minimized;
        }
        #endregion
    }
}
