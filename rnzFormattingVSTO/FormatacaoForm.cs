using System;
using System.Drawing;
using System.Windows.Forms;

namespace rnzFormattingVSTO
{
  public partial class FormatacaoForm : Form
  {
    public string ColunaMinima { get; private set; }
    public string ColunaMaxima { get; private set; }
    public string IntervaloFormatacao { get; private set; }

    public FormatacaoForm()
    {
      InitializeComponent();
      InitializeValues();
    }

    private void InitializeValues()
    {
      // Valores padrão
      textBoxColunaMin.Text = "E";
      textBoxColunaMax.Text = "F";
      textBoxIntervalo.Text = "G-I";
    }

    private void buttonOK_Click(object sender, EventArgs e)
    {
      // Validar entradas
      if (string.IsNullOrWhiteSpace(textBoxColunaMin.Text))
      {
        MessageBox.Show("Por favor, informe a coluna de valores mínimos.", "Validação",
            MessageBoxButtons.OK, MessageBoxIcon.Warning);
        textBoxColunaMin.Focus();
        return;
      }

      if (string.IsNullOrWhiteSpace(textBoxColunaMax.Text))
      {
        MessageBox.Show("Por favor, informe a coluna de valores máximos.", "Validação",
            MessageBoxButtons.OK, MessageBoxIcon.Warning);
        textBoxColunaMax.Focus();
        return;
      }

      if (string.IsNullOrWhiteSpace(textBoxIntervalo.Text))
      {
        MessageBox.Show("Por favor, informe o intervalo de colunas para formatação.", "Validação",
            MessageBoxButtons.OK, MessageBoxIcon.Warning);
        textBoxIntervalo.Focus();
        return;
      }

      // Validar formato das colunas
      if (!ValidarFormatoColuna(textBoxColunaMin.Text.Trim()))
      {
        MessageBox.Show("Formato inválido para coluna mínima. Use letras como A, B, AB, etc.", "Validação",
            MessageBoxButtons.OK, MessageBoxIcon.Warning);
        textBoxColunaMin.Focus();
        return;
      }

      if (!ValidarFormatoColuna(textBoxColunaMax.Text.Trim()))
      {
        MessageBox.Show("Formato inválido para coluna máxima. Use letras como A, B, AB, etc.", "Validação",
            MessageBoxButtons.OK, MessageBoxIcon.Warning);
        textBoxColunaMax.Focus();
        return;
      }

      if (!ValidarFormatoIntervalo(textBoxIntervalo.Text.Trim()))
      {
        MessageBox.Show("Formato inválido para intervalo. Use formatos como G-I, H, A-Z, etc.", "Validação",
            MessageBoxButtons.OK, MessageBoxIcon.Warning);
        textBoxIntervalo.Focus();
        return;
      }

      // Armazenar valores
      ColunaMinima = textBoxColunaMin.Text.Trim().ToUpper();
      ColunaMaxima = textBoxColunaMax.Text.Trim().ToUpper();
      IntervaloFormatacao = textBoxIntervalo.Text.Trim().ToUpper();

      DialogResult = DialogResult.OK;
      Close();
    }

    private void buttonCancelar_Click(object sender, EventArgs e)
    {
      DialogResult = DialogResult.Cancel;
      Close();
    }

    private bool ValidarFormatoColuna(string coluna)
    {
      if (string.IsNullOrEmpty(coluna)) return false;

      foreach (char c in coluna.ToUpper())
      {
        if (c < 'A' || c > 'Z') return false;
      }
      return true;
    }

    private bool ValidarFormatoIntervalo(string intervalo)
    {
      if (string.IsNullOrEmpty(intervalo)) return false;

      if (intervalo.Contains("-"))
      {
        var parts = intervalo.Split('-');
        if (parts.Length != 2) return false;
        return ValidarFormatoColuna(parts[0].Trim()) && ValidarFormatoColuna(parts[1].Trim());
      }
      else
      {
        return ValidarFormatoColuna(intervalo);
      }
    }

    private void textBox_KeyPress(object sender, KeyPressEventArgs e)
    {
      // Permitir apenas letras, backspace e hífen (para o intervalo)
      if (char.IsLetter(e.KeyChar) || e.KeyChar == (char)Keys.Back ||
          (sender == textBoxIntervalo && e.KeyChar == '-'))
      {
        e.KeyChar = char.ToUpper(e.KeyChar);
      }
      else
      {
        e.Handled = true;
      }
    }

    private void FormatacaoForm_KeyDown(object sender, KeyEventArgs e)
    {
      if (e.KeyCode == Keys.Enter)
      {
        buttonOK_Click(sender, e);
      }
      else if (e.KeyCode == Keys.Escape)
      {
        buttonCancelar_Click(sender, e);
      }
    }
  }
}