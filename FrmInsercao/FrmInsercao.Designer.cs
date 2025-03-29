namespace FrmInsercao
{
    partial class FrmInsercao
    {
        /// <summary>
        ///  Required designer variable.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        /// <summary>
        ///  Clean up any resources being used.
        /// </summary>
        /// <param name="disposing">true if managed resources should be disposed; otherwise, false.</param>
        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        #region Windows Form Designer generated code

        /// <summary>
        ///  Required method for Designer support - do not modify
        ///  the contents of this method with the code editor.
        /// </summary>
        private void InitializeComponent()
        {
            DataGridViewCellStyle dataGridViewCellStyle1 = new DataGridViewCellStyle();
            DataGridViewCellStyle dataGridViewCellStyle3 = new DataGridViewCellStyle();
            DataGridViewCellStyle dataGridViewCellStyle4 = new DataGridViewCellStyle();
            DataGridViewCellStyle dataGridViewCellStyle2 = new DataGridViewCellStyle();
            btnClose = new FontAwesome.Sharp.IconPictureBox();
            btnMaximize = new FontAwesome.Sharp.IconPictureBox();
            btnMinimize = new FontAwesome.Sharp.IconPictureBox();
            lblTituloForm = new Label();
            lblGrupo = new Label();
            txtPesquisar = new TextBox();
            lblPesquisar = new Label();
            lblDicas = new Label();
            dtgPrincipal = new DataGridView();
            colId = new DataGridViewTextBoxColumn();
            ColHashtag = new DataGridViewTextBoxColumn();
            ColValor = new DataGridViewTextBoxColumn();
            ColCriterio = new DataGridViewComboBoxColumn();
            colStatus = new DataGridViewComboBoxColumn();
            gpbStatus = new GroupBox();
            lblStatusN = new Label();
            lblStatusA = new Label();
            lblStatusF = new Label();
            lblStatusE = new Label();
            gpbCriterios = new GroupBox();
            lblNC = new Label();
            lblC = new Label();
            lblIgual = new Label();
            btnProcessar = new Button();
            btnFechar = new Button();
            ((System.ComponentModel.ISupportInitialize)btnClose).BeginInit();
            ((System.ComponentModel.ISupportInitialize)btnMaximize).BeginInit();
            ((System.ComponentModel.ISupportInitialize)btnMinimize).BeginInit();
            ((System.ComponentModel.ISupportInitialize)dtgPrincipal).BeginInit();
            gpbStatus.SuspendLayout();
            gpbCriterios.SuspendLayout();
            SuspendLayout();
            // 
            // btnClose
            // 
            btnClose.BackColor = Color.White;
            btnClose.BackgroundImageLayout = ImageLayout.None;
            btnClose.ForeColor = Color.Gainsboro;
            btnClose.IconChar = FontAwesome.Sharp.IconChar.Close;
            btnClose.IconColor = Color.Gainsboro;
            btnClose.IconFont = FontAwesome.Sharp.IconFont.Solid;
            btnClose.IconSize = 25;
            btnClose.Location = new Point(865, 0);
            btnClose.Name = "btnClose";
            btnClose.Size = new Size(35, 25);
            btnClose.SizeMode = PictureBoxSizeMode.CenterImage;
            btnClose.TabIndex = 0;
            btnClose.TabStop = false;
            btnClose.UseGdi = true;
            btnClose.Click += btnClose_Click;
            btnClose.MouseDown += btnClose_MouseDown;
            btnClose.MouseEnter += btnClose_MouseEnter;
            btnClose.MouseLeave += btnClose_MouseLeave;
            btnClose.MouseUp += btnClose_MouseUp;
            // 
            // btnMaximize
            // 
            btnMaximize.BackColor = Color.White;
            btnMaximize.BackgroundImageLayout = ImageLayout.None;
            btnMaximize.ForeColor = Color.Gainsboro;
            btnMaximize.IconChar = FontAwesome.Sharp.IconChar.WindowMaximize;
            btnMaximize.IconColor = Color.Gainsboro;
            btnMaximize.IconFont = FontAwesome.Sharp.IconFont.Solid;
            btnMaximize.IconSize = 25;
            btnMaximize.Location = new Point(835, 0);
            btnMaximize.Name = "btnMaximize";
            btnMaximize.Size = new Size(30, 25);
            btnMaximize.SizeMode = PictureBoxSizeMode.CenterImage;
            btnMaximize.TabIndex = 1;
            btnMaximize.TabStop = false;
            btnMaximize.UseGdi = true;
            btnMaximize.Click += btnMaximize_Click;
            btnMaximize.MouseDown += btnMaximize_MouseDown;
            btnMaximize.MouseEnter += btnMaximize_MouseEnter;
            btnMaximize.MouseLeave += btnMaximize_MouseLeave;
            btnMaximize.MouseUp += btnMaximize_MouseUp;
            // 
            // btnMinimize
            // 
            btnMinimize.BackColor = Color.White;
            btnMinimize.BackgroundImageLayout = ImageLayout.None;
            btnMinimize.ForeColor = Color.Gainsboro;
            btnMinimize.IconChar = FontAwesome.Sharp.IconChar.WindowMinimize;
            btnMinimize.IconColor = Color.Gainsboro;
            btnMinimize.IconFont = FontAwesome.Sharp.IconFont.Solid;
            btnMinimize.IconSize = 25;
            btnMinimize.Location = new Point(805, 0);
            btnMinimize.Name = "btnMinimize";
            btnMinimize.Size = new Size(30, 25);
            btnMinimize.SizeMode = PictureBoxSizeMode.CenterImage;
            btnMinimize.TabIndex = 2;
            btnMinimize.TabStop = false;
            btnMinimize.UseGdi = true;
            btnMinimize.Click += btnMinimize_Click;
            btnMinimize.MouseDown += btnMinimize_MouseDown;
            btnMinimize.MouseEnter += btnMinimize_MouseEnter;
            btnMinimize.MouseLeave += btnMinimize_MouseLeave;
            btnMinimize.MouseUp += btnMinimize_MouseUp;
            // 
            // lblTituloForm
            // 
            lblTituloForm.BackColor = Color.White;
            lblTituloForm.Font = new Font("Arial", 9F, FontStyle.Regular, GraphicsUnit.Point, 0);
            lblTituloForm.ForeColor = Color.DimGray;
            lblTituloForm.Location = new Point(0, 0);
            lblTituloForm.Name = "lblTituloForm";
            lblTituloForm.Size = new Size(805, 25);
            lblTituloForm.TabIndex = 3;
            lblTituloForm.Text = "  # Formulário Inserção";
            lblTituloForm.TextAlign = ContentAlignment.MiddleLeft;
            // 
            // lblGrupo
            // 
            lblGrupo.BackColor = Color.Transparent;
            lblGrupo.Font = new Font("Arial", 12F, FontStyle.Bold, GraphicsUnit.Point, 0);
            lblGrupo.ForeColor = Color.SteelBlue;
            lblGrupo.Location = new Point(0, 28);
            lblGrupo.Name = "lblGrupo";
            lblGrupo.Size = new Size(900, 30);
            lblGrupo.TabIndex = 4;
            lblGrupo.Text = "Grupo 333 - MEIA BESTA";
            lblGrupo.TextAlign = ContentAlignment.MiddleCenter;
            // 
            // txtPesquisar
            // 
            txtPesquisar.BackColor = Color.Gainsboro;
            txtPesquisar.BorderStyle = BorderStyle.FixedSingle;
            txtPesquisar.ForeColor = Color.DimGray;
            txtPesquisar.Location = new Point(738, 61);
            txtPesquisar.Name = "txtPesquisar";
            txtPesquisar.Size = new Size(150, 21);
            txtPesquisar.TabIndex = 5;
            // 
            // lblPesquisar
            // 
            lblPesquisar.ForeColor = Color.DimGray;
            lblPesquisar.Location = new Point(665, 61);
            lblPesquisar.Name = "lblPesquisar";
            lblPesquisar.Size = new Size(67, 21);
            lblPesquisar.TabIndex = 6;
            lblPesquisar.Text = "Pesquisar:";
            lblPesquisar.TextAlign = ContentAlignment.MiddleLeft;
            // 
            // lblDicas
            // 
            lblDicas.BackColor = Color.Transparent;
            lblDicas.Font = new Font("Arial", 9F, FontStyle.Bold, GraphicsUnit.Point, 0);
            lblDicas.ForeColor = Color.SteelBlue;
            lblDicas.ImageAlign = ContentAlignment.MiddleLeft;
            lblDicas.Location = new Point(0, 65);
            lblDicas.Name = "lblDicas";
            lblDicas.Size = new Size(100, 20);
            lblDicas.TabIndex = 7;
            lblDicas.Text = " Dicas:";
            lblDicas.TextAlign = ContentAlignment.MiddleLeft;
            // 
            // dtgPrincipal
            // 
            dtgPrincipal.AllowUserToAddRows = false;
            dtgPrincipal.BackgroundColor = Color.DarkGray;
            dtgPrincipal.BorderStyle = BorderStyle.Fixed3D;
            dtgPrincipal.ColumnHeadersBorderStyle = DataGridViewHeaderBorderStyle.Single;
            dataGridViewCellStyle1.Alignment = DataGridViewContentAlignment.MiddleLeft;
            dataGridViewCellStyle1.BackColor = Color.White;
            dataGridViewCellStyle1.Font = new Font("Arial", 9F, FontStyle.Regular, GraphicsUnit.Point, 0);
            dataGridViewCellStyle1.ForeColor = Color.DimGray;
            dataGridViewCellStyle1.SelectionBackColor = Color.White;
            dataGridViewCellStyle1.SelectionForeColor = Color.DimGray;
            dataGridViewCellStyle1.WrapMode = DataGridViewTriState.True;
            dtgPrincipal.ColumnHeadersDefaultCellStyle = dataGridViewCellStyle1;
            dtgPrincipal.ColumnHeadersHeightSizeMode = DataGridViewColumnHeadersHeightSizeMode.DisableResizing;
            dtgPrincipal.Columns.AddRange(new DataGridViewColumn[] { colId, ColHashtag, ColValor, ColCriterio, colStatus });
            dataGridViewCellStyle3.Alignment = DataGridViewContentAlignment.MiddleLeft;
            dataGridViewCellStyle3.BackColor = Color.LightGray;
            dataGridViewCellStyle3.Font = new Font("Arial", 9F, FontStyle.Regular, GraphicsUnit.Point, 0);
            dataGridViewCellStyle3.ForeColor = Color.DimGray;
            dataGridViewCellStyle3.SelectionBackColor = SystemColors.Highlight;
            dataGridViewCellStyle3.SelectionForeColor = SystemColors.HighlightText;
            dataGridViewCellStyle3.WrapMode = DataGridViewTriState.False;
            dtgPrincipal.DefaultCellStyle = dataGridViewCellStyle3;
            dtgPrincipal.EnableHeadersVisualStyles = false;
            dtgPrincipal.GridColor = Color.DimGray;
            dtgPrincipal.Location = new Point(0, 88);
            dtgPrincipal.Name = "dtgPrincipal";
            dtgPrincipal.RowHeadersBorderStyle = DataGridViewHeaderBorderStyle.Single;
            dataGridViewCellStyle4.Alignment = DataGridViewContentAlignment.MiddleLeft;
            dataGridViewCellStyle4.BackColor = Color.White;
            dataGridViewCellStyle4.Font = new Font("Arial", 9F, FontStyle.Regular, GraphicsUnit.Point, 0);
            dataGridViewCellStyle4.ForeColor = Color.DimGray;
            dataGridViewCellStyle4.SelectionBackColor = SystemColors.Highlight;
            dataGridViewCellStyle4.SelectionForeColor = SystemColors.HighlightText;
            dataGridViewCellStyle4.WrapMode = DataGridViewTriState.True;
            dtgPrincipal.RowHeadersDefaultCellStyle = dataGridViewCellStyle4;
            dtgPrincipal.RowHeadersVisible = false;
            dtgPrincipal.RowTemplate.Height = 21;
            dtgPrincipal.Size = new Size(900, 262);
            dtgPrincipal.TabIndex = 8;
            // 
            // colId
            // 
            colId.AutoSizeMode = DataGridViewAutoSizeColumnMode.None;
            dataGridViewCellStyle2.Alignment = DataGridViewContentAlignment.TopCenter;
            colId.DefaultCellStyle = dataGridViewCellStyle2;
            colId.FillWeight = 25.38071F;
            colId.HeaderText = "ID";
            colId.Name = "colId";
            colId.Width = 55;
            // 
            // ColHashtag
            // 
            ColHashtag.FillWeight = 118.654823F;
            ColHashtag.HeaderText = "# Hash";
            ColHashtag.Name = "ColHashtag";
            ColHashtag.Width = 160;
            // 
            // ColValor
            // 
            ColValor.FillWeight = 118.654823F;
            ColValor.HeaderText = "Valor(es) separados por \"ponto e vírgula\"";
            ColValor.Name = "ColValor";
            ColValor.Width = 450;
            // 
            // ColCriterio
            // 
            ColCriterio.FillWeight = 118.654823F;
            ColCriterio.HeaderText = "Critério";
            ColCriterio.Items.AddRange(new object[] { "C", "I", "NC" });
            ColCriterio.Name = "ColCriterio";
            ColCriterio.Resizable = DataGridViewTriState.True;
            ColCriterio.SortMode = DataGridViewColumnSortMode.Automatic;
            ColCriterio.Width = 132;
            // 
            // colStatus
            // 
            colStatus.FillWeight = 118.654823F;
            colStatus.HeaderText = "Status";
            colStatus.Items.AddRange(new object[] { "A", "E", "F", "N" });
            colStatus.Name = "colStatus";
            colStatus.Resizable = DataGridViewTriState.True;
            colStatus.SortMode = DataGridViewColumnSortMode.Automatic;
            // 
            // gpbStatus
            // 
            gpbStatus.BackColor = Color.LightGray;
            gpbStatus.Controls.Add(lblStatusN);
            gpbStatus.Controls.Add(lblStatusA);
            gpbStatus.Controls.Add(lblStatusF);
            gpbStatus.Controls.Add(lblStatusE);
            gpbStatus.FlatStyle = FlatStyle.Flat;
            gpbStatus.Font = new Font("Arial", 9F, FontStyle.Bold, GraphicsUnit.Point, 0);
            gpbStatus.ForeColor = Color.DimGray;
            gpbStatus.Location = new Point(12, 356);
            gpbStatus.Name = "gpbStatus";
            gpbStatus.Size = new Size(645, 65);
            gpbStatus.TabIndex = 9;
            gpbStatus.TabStop = false;
            gpbStatus.Text = "Status";
            // 
            // lblStatusN
            // 
            lblStatusN.BackColor = Color.Transparent;
            lblStatusN.Font = new Font("Arial", 9F, FontStyle.Bold, GraphicsUnit.Point, 0);
            lblStatusN.ForeColor = Color.SteelBlue;
            lblStatusN.ImageAlign = ContentAlignment.MiddleLeft;
            lblStatusN.Location = new Point(114, 37);
            lblStatusN.Name = "lblStatusN";
            lblStatusN.Size = new Size(88, 20);
            lblStatusN.TabIndex = 13;
            lblStatusN.Text = "N - Novo";
            lblStatusN.TextAlign = ContentAlignment.MiddleLeft;
            // 
            // lblStatusA
            // 
            lblStatusA.BackColor = Color.Transparent;
            lblStatusA.Font = new Font("Arial", 9F, FontStyle.Bold, GraphicsUnit.Point, 0);
            lblStatusA.ForeColor = Color.SteelBlue;
            lblStatusA.ImageAlign = ContentAlignment.MiddleLeft;
            lblStatusA.Location = new Point(20, 17);
            lblStatusA.Name = "lblStatusA";
            lblStatusA.Size = new Size(88, 20);
            lblStatusA.TabIndex = 10;
            lblStatusA.Text = "A - Alteração";
            lblStatusA.TextAlign = ContentAlignment.MiddleLeft;
            // 
            // lblStatusF
            // 
            lblStatusF.BackColor = Color.Transparent;
            lblStatusF.Font = new Font("Arial", 9F, FontStyle.Bold, GraphicsUnit.Point, 0);
            lblStatusF.ForeColor = Color.SteelBlue;
            lblStatusF.ImageAlign = ContentAlignment.MiddleLeft;
            lblStatusF.Location = new Point(20, 37);
            lblStatusF.Name = "lblStatusF";
            lblStatusF.Size = new Size(88, 20);
            lblStatusF.TabIndex = 12;
            lblStatusF.Text = "F - Finalizado";
            lblStatusF.TextAlign = ContentAlignment.MiddleLeft;
            // 
            // lblStatusE
            // 
            lblStatusE.BackColor = Color.Transparent;
            lblStatusE.Font = new Font("Arial", 9F, FontStyle.Bold, GraphicsUnit.Point, 0);
            lblStatusE.ForeColor = Color.SteelBlue;
            lblStatusE.ImageAlign = ContentAlignment.MiddleLeft;
            lblStatusE.Location = new Point(114, 17);
            lblStatusE.Name = "lblStatusE";
            lblStatusE.Size = new Size(88, 20);
            lblStatusE.TabIndex = 11;
            lblStatusE.Text = "E - Exclusão";
            lblStatusE.TextAlign = ContentAlignment.MiddleLeft;
            // 
            // gpbCriterios
            // 
            gpbCriterios.BackColor = Color.LightGray;
            gpbCriterios.Controls.Add(lblNC);
            gpbCriterios.Controls.Add(lblC);
            gpbCriterios.Controls.Add(lblIgual);
            gpbCriterios.FlatStyle = FlatStyle.Flat;
            gpbCriterios.Font = new Font("Arial", 9F, FontStyle.Bold, GraphicsUnit.Point, 0);
            gpbCriterios.ForeColor = Color.DimGray;
            gpbCriterios.Location = new Point(12, 427);
            gpbCriterios.Name = "gpbCriterios";
            gpbCriterios.Size = new Size(876, 86);
            gpbCriterios.TabIndex = 14;
            gpbCriterios.TabStop = false;
            gpbCriterios.Text = "Critérios";
            // 
            // lblNC
            // 
            lblNC.BackColor = Color.Transparent;
            lblNC.Font = new Font("Arial", 9F, FontStyle.Bold, GraphicsUnit.Point, 0);
            lblNC.ForeColor = Color.SteelBlue;
            lblNC.ImageAlign = ContentAlignment.MiddleLeft;
            lblNC.Location = new Point(20, 37);
            lblNC.Name = "lblNC";
            lblNC.Size = new Size(700, 20);
            lblNC.TabIndex = 13;
            lblNC.Text = "NC - Pode ser inserido mais de um valor separado por \";\". Caso [...] não será sugerido.";
            lblNC.TextAlign = ContentAlignment.MiddleLeft;
            // 
            // lblC
            // 
            lblC.BackColor = Color.Transparent;
            lblC.Font = new Font("Arial", 9F, FontStyle.Bold, GraphicsUnit.Point, 0);
            lblC.ForeColor = Color.SteelBlue;
            lblC.ImageAlign = ContentAlignment.MiddleLeft;
            lblC.Location = new Point(20, 59);
            lblC.Name = "lblC";
            lblC.Size = new Size(700, 20);
            lblC.TabIndex = 11;
            lblC.Text = "C - Será obrigatório nformar pelo menos um valor. Caso o usuário informe mais de um valor [...] OBRIGATORIAMENTE [...]";
            lblC.TextAlign = ContentAlignment.MiddleLeft;
            // 
            // lblIgual
            // 
            lblIgual.BackColor = Color.Transparent;
            lblIgual.Font = new Font("Arial", 9F, FontStyle.Bold, GraphicsUnit.Point, 0);
            lblIgual.ForeColor = Color.SteelBlue;
            lblIgual.ImageAlign = ContentAlignment.MiddleLeft;
            lblIgual.Location = new Point(20, 17);
            lblIgual.Name = "lblIgual";
            lblIgual.Size = new Size(700, 20);
            lblIgual.TabIndex = 10;
            lblIgual.Text = "I - Não é necessário informar um valor.";
            lblIgual.TextAlign = ContentAlignment.MiddleLeft;
            // 
            // btnProcessar
            // 
            btnProcessar.BackColor = Color.LightGray;
            btnProcessar.FlatStyle = FlatStyle.Flat;
            btnProcessar.ForeColor = Color.DimGray;
            btnProcessar.Location = new Point(665, 375);
            btnProcessar.Name = "btnProcessar";
            btnProcessar.Size = new Size(110, 25);
            btnProcessar.TabIndex = 14;
            btnProcessar.Text = "Processar";
            btnProcessar.UseVisualStyleBackColor = false;
            // 
            // btnFechar
            // 
            btnFechar.BackColor = Color.LightGray;
            btnFechar.FlatStyle = FlatStyle.Flat;
            btnFechar.ForeColor = Color.DimGray;
            btnFechar.Location = new Point(778, 375);
            btnFechar.Name = "btnFechar";
            btnFechar.Size = new Size(110, 25);
            btnFechar.TabIndex = 15;
            btnFechar.Text = "Fechar";
            btnFechar.UseVisualStyleBackColor = false;
            btnFechar.Click += btnFechar_Click;
            // 
            // FrmInsercao
            // 
            AutoScaleDimensions = new SizeF(7F, 15F);
            AutoScaleMode = AutoScaleMode.Font;
            BackColor = Color.Gainsboro;
            ClientSize = new Size(900, 520);
            Controls.Add(btnFechar);
            Controls.Add(btnProcessar);
            Controls.Add(gpbCriterios);
            Controls.Add(gpbStatus);
            Controls.Add(dtgPrincipal);
            Controls.Add(lblDicas);
            Controls.Add(lblPesquisar);
            Controls.Add(txtPesquisar);
            Controls.Add(lblGrupo);
            Controls.Add(lblTituloForm);
            Controls.Add(btnMinimize);
            Controls.Add(btnMaximize);
            Controls.Add(btnClose);
            Font = new Font("Arial", 9F, FontStyle.Regular, GraphicsUnit.Point, 0);
            ForeColor = Color.DimGray;
            FormBorderStyle = FormBorderStyle.None;
            Margin = new Padding(4, 3, 4, 3);
            Name = "FrmInsercao";
            StartPosition = FormStartPosition.CenterScreen;
            Text = "FrmInsercao";
            Load += FrmInsercao_Load;
            ((System.ComponentModel.ISupportInitialize)btnClose).EndInit();
            ((System.ComponentModel.ISupportInitialize)btnMaximize).EndInit();
            ((System.ComponentModel.ISupportInitialize)btnMinimize).EndInit();
            ((System.ComponentModel.ISupportInitialize)dtgPrincipal).EndInit();
            gpbStatus.ResumeLayout(false);
            gpbCriterios.ResumeLayout(false);
            ResumeLayout(false);
            PerformLayout();
        }

        #endregion

        private FontAwesome.Sharp.IconPictureBox btnClose;
        private FontAwesome.Sharp.IconPictureBox btnMaximize;
        private FontAwesome.Sharp.IconPictureBox btnMinimize;
        private Label lblTituloForm;
        private Label lblGrupo;
        private TextBox txtPesquisar;
        private Label lblPesquisar;
        private Label lblDicas;
        private DataGridView dtgPrincipal;
        private GroupBox gpbStatus;
        private Label lblStatusA;
        private Label lblStatusE;
        private Label lblStatusN;
        private Label lblStatusF;
        private GroupBox gpbCriterios;
        private Label lblNC;
        private Label lblC;
        private Label lblIgual;
        private Button btnProcessar;
        private Button btnFechar;
        private DataGridViewTextBoxColumn colId;
        private DataGridViewTextBoxColumn ColHashtag;
        private DataGridViewTextBoxColumn ColValor;
        private DataGridViewComboBoxColumn ColCriterio;
        private DataGridViewComboBoxColumn colStatus;
    }
}
