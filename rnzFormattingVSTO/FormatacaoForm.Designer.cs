namespace rnzFormattingVSTO
{
  partial class FormatacaoForm
  {
    /// <summary>
    /// Required designer variable.
    /// </summary>
    private System.ComponentModel.IContainer components = null;
    private System.Windows.Forms.Label labelColunaMin;
    private System.Windows.Forms.TextBox textBoxColunaMin;
    private System.Windows.Forms.Label labelColunaMax;
    private System.Windows.Forms.TextBox textBoxColunaMax;
    private System.Windows.Forms.Label labelIntervalo;
    private System.Windows.Forms.TextBox textBoxIntervalo;
    private System.Windows.Forms.Button buttonOK;
    private System.Windows.Forms.Button buttonCancelar;
    private System.Windows.Forms.GroupBox groupBoxConfiguracoes;
    private System.Windows.Forms.Label labelInstrucoes;
    private System.Windows.Forms.Label labelExemplo;

    /// <summary>
    /// Clean up any resources being used.
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
    /// Required method for Designer support - do not modify
    /// the contents of this method with the code editor.
    /// </summary>
    private void InitializeComponent()
    {
      this.labelColunaMin = new System.Windows.Forms.Label();
      this.textBoxColunaMin = new System.Windows.Forms.TextBox();
      this.labelColunaMax = new System.Windows.Forms.Label();
      this.textBoxColunaMax = new System.Windows.Forms.TextBox();
      this.labelIntervalo = new System.Windows.Forms.Label();
      this.textBoxIntervalo = new System.Windows.Forms.TextBox();
      this.buttonOK = new System.Windows.Forms.Button();
      this.buttonCancelar = new System.Windows.Forms.Button();
      this.groupBoxConfiguracoes = new System.Windows.Forms.GroupBox();
      this.labelExemplo = new System.Windows.Forms.Label();
      this.labelInstrucoes = new System.Windows.Forms.Label();
      this.groupBoxConfiguracoes.SuspendLayout();
      this.SuspendLayout();
      // 
      // labelColunaMin
      // 
      this.labelColunaMin.AutoSize = true;
      this.labelColunaMin.Location = new System.Drawing.Point(13, 26);
      this.labelColunaMin.Name = "labelColunaMin";
      this.labelColunaMin.Size = new System.Drawing.Size(124, 13);
      this.labelColunaMin.TabIndex = 0;
      this.labelColunaMin.Text = "Coluna Valores Mínimos:";
      // 
      // textBoxColunaMin
      // 
      this.textBoxColunaMin.Location = new System.Drawing.Point(146, 23);
      this.textBoxColunaMin.MaxLength = 3;
      this.textBoxColunaMin.Name = "textBoxColunaMin";
      this.textBoxColunaMin.Size = new System.Drawing.Size(52, 20);
      this.textBoxColunaMin.TabIndex = 1;
      this.textBoxColunaMin.TextAlign = System.Windows.Forms.HorizontalAlignment.Center;
      this.textBoxColunaMin.KeyPress += new System.Windows.Forms.KeyPressEventHandler(this.textBox_KeyPress);
      // 
      // labelColunaMax
      // 
      this.labelColunaMax.AutoSize = true;
      this.labelColunaMax.Location = new System.Drawing.Point(13, 56);
      this.labelColunaMax.Name = "labelColunaMax";
      this.labelColunaMax.Size = new System.Drawing.Size(125, 13);
      this.labelColunaMax.TabIndex = 2;
      this.labelColunaMax.Text = "Coluna Valores Máximos:";
      // 
      // textBoxColunaMax
      // 
      this.textBoxColunaMax.Location = new System.Drawing.Point(146, 54);
      this.textBoxColunaMax.MaxLength = 3;
      this.textBoxColunaMax.Name = "textBoxColunaMax";
      this.textBoxColunaMax.Size = new System.Drawing.Size(52, 20);
      this.textBoxColunaMax.TabIndex = 3;
      this.textBoxColunaMax.TextAlign = System.Windows.Forms.HorizontalAlignment.Center;
      this.textBoxColunaMax.KeyPress += new System.Windows.Forms.KeyPressEventHandler(this.textBox_KeyPress);
      // 
      // labelIntervalo
      // 
      this.labelIntervalo.AutoSize = true;
      this.labelIntervalo.Location = new System.Drawing.Point(13, 87);
      this.labelIntervalo.Name = "labelIntervalo";
      this.labelIntervalo.Size = new System.Drawing.Size(119, 13);
      this.labelIntervalo.TabIndex = 4;
      this.labelIntervalo.Text = "Intervalo para Formatar:";
      // 
      // textBoxIntervalo
      // 
      this.textBoxIntervalo.Location = new System.Drawing.Point(146, 84);
      this.textBoxIntervalo.MaxLength = 10;
      this.textBoxIntervalo.Name = "textBoxIntervalo";
      this.textBoxIntervalo.Size = new System.Drawing.Size(52, 20);
      this.textBoxIntervalo.TabIndex = 5;
      this.textBoxIntervalo.TextAlign = System.Windows.Forms.HorizontalAlignment.Center;
      this.textBoxIntervalo.KeyPress += new System.Windows.Forms.KeyPressEventHandler(this.textBox_KeyPress);
      // 
      // buttonOK
      // 
      this.buttonOK.Location = new System.Drawing.Point(12, 238);
      this.buttonOK.Name = "buttonOK";
      this.buttonOK.Size = new System.Drawing.Size(64, 26);
      this.buttonOK.TabIndex = 6;
      this.buttonOK.Text = "OK";
      this.buttonOK.UseVisualStyleBackColor = true;
      this.buttonOK.Click += new System.EventHandler(this.buttonOK_Click);
      // 
      // buttonCancelar
      // 
      this.buttonCancelar.Location = new System.Drawing.Point(90, 238);
      this.buttonCancelar.Name = "buttonCancelar";
      this.buttonCancelar.Size = new System.Drawing.Size(64, 26);
      this.buttonCancelar.TabIndex = 7;
      this.buttonCancelar.Text = "Cancelar";
      this.buttonCancelar.UseVisualStyleBackColor = true;
      this.buttonCancelar.Click += new System.EventHandler(this.buttonCancelar_Click);
      // 
      // groupBoxConfiguracoes
      // 
      this.groupBoxConfiguracoes.Controls.Add(this.labelExemplo);
      this.groupBoxConfiguracoes.Controls.Add(this.labelColunaMin);
      this.groupBoxConfiguracoes.Controls.Add(this.textBoxColunaMin);
      this.groupBoxConfiguracoes.Controls.Add(this.labelColunaMax);
      this.groupBoxConfiguracoes.Controls.Add(this.textBoxColunaMax);
      this.groupBoxConfiguracoes.Controls.Add(this.labelIntervalo);
      this.groupBoxConfiguracoes.Controls.Add(this.textBoxIntervalo);
      this.groupBoxConfiguracoes.Location = new System.Drawing.Point(17, 66);
      this.groupBoxConfiguracoes.Name = "groupBoxConfiguracoes";
      this.groupBoxConfiguracoes.Size = new System.Drawing.Size(296, 166);
      this.groupBoxConfiguracoes.TabIndex = 8;
      this.groupBoxConfiguracoes.TabStop = false;
      this.groupBoxConfiguracoes.Text = "Configurações de Formatação";
      // 
      // labelExemplo
      // 
      this.labelExemplo.Font = new System.Drawing.Font("Segoe UI", 8.25F, System.Drawing.FontStyle.Italic);
      this.labelExemplo.ForeColor = System.Drawing.Color.DarkBlue;
      this.labelExemplo.Location = new System.Drawing.Point(13, 113);
      this.labelExemplo.Name = "labelExemplo";
      this.labelExemplo.Size = new System.Drawing.Size(274, 43);
      this.labelExemplo.TabIndex = 6;
      this.labelExemplo.Text = "Exemplos:\r\n• Colunas: A, B, C, AB, AC, etc.\r\n• Intervalos: G-I, A-Z, H (uma única" +
    " coluna)";
      // 
      // labelInstrucoes
      // 
      this.labelInstrucoes.Font = new System.Drawing.Font("Segoe UI", 9F);
      this.labelInstrucoes.Location = new System.Drawing.Point(17, 13);
      this.labelInstrucoes.Name = "labelInstrucoes";
      this.labelInstrucoes.Size = new System.Drawing.Size(296, 50);
      this.labelInstrucoes.TabIndex = 9;
      this.labelInstrucoes.Text = "Configure as colunas para formatação condicional. Os valores serão coloridos em v" +
    "erde (dentro dos limites) ou vermelho (fora dos limites).";
      // 
      // FormatacaoForm
      // 
      this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
      this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
      this.ClientSize = new System.Drawing.Size(329, 271);
      this.Controls.Add(this.labelInstrucoes);
      this.Controls.Add(this.groupBoxConfiguracoes);
      this.Controls.Add(this.buttonCancelar);
      this.Controls.Add(this.buttonOK);
      this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedDialog;
      this.KeyPreview = true;
      this.MaximizeBox = false;
      this.MinimizeBox = false;
      this.Name = "FormatacaoForm";
      this.StartPosition = System.Windows.Forms.FormStartPosition.CenterParent;
      this.Text = "RNZ Formatting - Configurações";
      this.KeyDown += new System.Windows.Forms.KeyEventHandler(this.FormatacaoForm_KeyDown);
      this.groupBoxConfiguracoes.ResumeLayout(false);
      this.groupBoxConfiguracoes.PerformLayout();
      this.ResumeLayout(false);

    }

    #endregion
  }
}