using System;
using System.Collections.Generic;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Text;
using System.Windows.Forms;
using Office = Microsoft.Office.Core;
using Excel = Microsoft.Office.Interop.Excel;

// TODO: Follow these steps to enable the Ribbon (XML) item:

// 1: Copy the following code block into the ThisAddin, ThisWorkbook, or ThisDocument class.

//  protected override Microsoft.Office.Core.IRibbonExtensibility CreateRibbonExtensibilityObject()
//  {
//      return new FormattingRibbon();
//  }

// 2. Create callback methods in the "Ribbon Callbacks" region of this class to handle user
//    actions, such as button clicks. Note: if you have exported this Ribbon from the Ribbon designer,
//    move your code from the event handlers to the callback methods and modify the code in the Ribbon designer.

// 3. Assign attributes to the control tags in the Ribbon XML file to identify the appropriate callback methods in your code.  

// For more information, see the Ribbon XML documentation in the Visual Studio Tools for Office Help.


namespace rnzFormattingVSTO
{
  [ComVisible(true)]
  public class FormattingRibbon : Office.IRibbonExtensibility
  {
    private Office.IRibbonUI ribbon;

    public FormattingRibbon()
    {
    }

    #region IRibbonExtensibility Members

    public string GetCustomUI(string ribbonID)
    {
      return GetResourceText("rnzFormattingVSTO.FormattingRibbon.xml");
    }

    #endregion

    #region Ribbon Callbacks
    //Create callback methods here. For more information about adding callback methods, visit https://go.microsoft.com/fwlink/?LinkID=271226

    public void Ribbon_Load(Office.IRibbonUI ribbonUI)
    {
      this.ribbon = ribbonUI;
    }

    public void OnFormatarLimites(Office.IRibbonControl control)
    {
      try
      {
        // Verificar se existe uma planilha ativa
        Excel.Application app = Globals.ThisAddIn.Application;
        Excel.Worksheet activeSheet = app.ActiveSheet as Excel.Worksheet;

        if (activeSheet == null)
        {
          MessageBox.Show("Nenhuma planilha ativa encontrada.", "Erro", MessageBoxButtons.OK, MessageBoxIcon.Warning);
          return;
        }

        // Abrir o form de configuração
        using (FormatacaoForm form = new FormatacaoForm())
        {
          if (form.ShowDialog() == DialogResult.OK)
          {
            // Processar formatação com os valores do form
            FormatarCelulas(activeSheet, form.ColunaMinima, form.ColunaMaxima, form.IntervaloFormatacao);

            MessageBox.Show("Formatação concluída com sucesso!", "Sucesso", MessageBoxButtons.OK, MessageBoxIcon.Information);
          }
        }
      }
      catch (Exception ex)
      {
        MessageBox.Show($"Erro durante a formatação:\n{ex.Message}", "Erro", MessageBoxButtons.OK, MessageBoxIcon.Error);
      }
    }

    private void FormatarCelulas(Excel.Worksheet worksheet, string colMin, string colMax, string rangeInput)
    {
      // Extrair colunas inicial e final
      string startCol, endCol;
      if (rangeInput.Contains("-"))
      {
        var parts = rangeInput.Split('-');
        startCol = parts[0].Trim();
        endCol = parts[1].Trim();
      }
      else
      {
        startCol = rangeInput;
        endCol = rangeInput;
      }

      // Encontrar a última linha com dados
      Excel.Range usedRange = worksheet.UsedRange;
      int lastRow = usedRange.Rows.Count + usedRange.Row - 1;

      // Obter índices das colunas
      int startColIndex = GetColumnIndex(startCol);
      int endColIndex = GetColumnIndex(endCol);
      int minColIndex = GetColumnIndex(colMin);
      int maxColIndex = GetColumnIndex(colMax);

      // Limpar formatação existente
      Excel.Range formatRange = worksheet.Range[startCol + "2:" + endCol + lastRow];
      formatRange.Font.ColorIndex = Excel.XlColorIndex.xlColorIndexAutomatic;

      // Processar cada linha
      for (int row = 2; row <= lastRow; row++)
      {
        // Obter valores mínimo e máximo para esta linha
        Excel.Range minCell = worksheet.Cells[row, minColIndex];
        Excel.Range maxCell = worksheet.Cells[row, maxColIndex];

        if (minCell.Value != null && maxCell.Value != null &&
            IsNumeric(minCell.Value) && IsNumeric(maxCell.Value))
        {
          double minValue = Convert.ToDouble(minCell.Value);
          double maxValue = Convert.ToDouble(maxCell.Value);

          // Processar cada coluna no intervalo para esta linha
          for (int col = startColIndex; col <= endColIndex; col++)
          {
            Excel.Range cell = worksheet.Cells[row, col];
            if (cell.Value != null && IsNumeric(cell.Value))
            {
              double cellValue = Convert.ToDouble(cell.Value);

              if (cellValue >= minValue && cellValue <= maxValue)
              {
                // Dentro dos limites - Texto VERDE
                cell.Font.Color = ColorTranslator.ToOle(Color.Green);
              }
              else
              {
                // Fora dos limites - Texto VERMELHO
                cell.Font.Color = ColorTranslator.ToOle(Color.Red);
              }
            }
          }
        }
      }
    }

    // Converter letra de coluna em índice numérico
    private int GetColumnIndex(string colLetter)
    {
      int index = 0;
      foreach (char c in colLetter.ToUpper())
      {
        index = index * 26 + (c - 'A' + 1);
      }
      return index;
    }

    // Verificar se um valor é numérico
    private bool IsNumeric(object value)
    {
      return value is double || value is int || value is decimal || value is float;
    }

    #endregion

    #region Helpers

    private static string GetResourceText(string resourceName)
    {
      Assembly asm = Assembly.GetExecutingAssembly();
      string[] resourceNames = asm.GetManifestResourceNames();
      for (int i = 0; i < resourceNames.Length; ++i)
      {
        if (string.Compare(resourceName, resourceNames[i], StringComparison.OrdinalIgnoreCase) == 0)
        {
          using (StreamReader resourceReader = new StreamReader(asm.GetManifestResourceStream(resourceNames[i])))
          {
            if (resourceReader != null)
            {
              return resourceReader.ReadToEnd();
            }
          }
        }
      }
      return null;
    }

    #endregion
  }
}