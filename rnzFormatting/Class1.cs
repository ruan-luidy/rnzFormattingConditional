using System;
using System.Drawing;
using System.Runtime.InteropServices;
using ExcelDna.Integration;
using ExcelDna.Integration.CustomUI;
using Microsoft.Office.Interop.Excel;
using ExcelRange = Microsoft.Office.Interop.Excel.Range;

namespace rnzFormatting
{
  public class FormatacaoCondicional
  {
    [ExcelFunction(Description = "Formata células com base em valores mínimos e máximos")]
    public static string FormatarPorLimites()
    {
      try
      {
        // Obter a aplicação Excel
        var app = (Microsoft.Office.Interop.Excel.Application)ExcelDnaUtil.Application;
        var activeSheet = app.ActiveSheet as Worksheet;

        if (activeSheet == null)
        {
          return "Erro: Nenhuma planilha ativa encontrada";
        }

        // Solicitar informações do usuário
        string colMin = SolicitarEntrada("Digite a letra da coluna que contém os valores MÍNIMOS (ex: E):", "Coluna de Mínimos");
        if (string.IsNullOrEmpty(colMin)) return "Operação cancelada";

        string colMax = SolicitarEntrada("Digite a letra da coluna que contém os valores MÁXIMOS (ex: F):", "Coluna de Máximos");
        if (string.IsNullOrEmpty(colMax)) return "Operação cancelada";

        string rangeInput = SolicitarEntrada("Digite o intervalo de colunas para formatar (ex: G-I):", "Intervalo de Colunas");
        if (string.IsNullOrEmpty(rangeInput)) return "Operação cancelada";

        // Extrair as colunas inicial e final
        string startCol, endCol;
        if (rangeInput.Contains("-"))
        {
          var parts = rangeInput.Split('-');
          startCol = parts[0].Trim().ToUpper();
          endCol = parts[1].Trim().ToUpper();
        }
        else
        {
          startCol = rangeInput.Trim().ToUpper();
          endCol = rangeInput.Trim().ToUpper();
        }

        // Encontrar a última linha com dados
        ExcelRange minColRange = (ExcelRange)activeSheet.Range[$"{colMin}:{colMin}"];
        int lastRow = activeSheet.Cells.SpecialCells(XlCellType.xlCellTypeLastCell).Row;

        // Definir o intervalo a ser formatado
        ExcelRange formatRange = (ExcelRange)activeSheet.Range[$"{startCol}2:{endCol}{lastRow}"];

        // Limpar formatação existente
        formatRange.Font.ColorIndex = XlColorIndex.xlColorIndexAutomatic;

        // Processar cada célula no intervalo
        int startColIndex = ObterIndiceColuna(startCol);
        int endColIndex = ObterIndiceColuna(endCol);
        int minColIndex = ObterIndiceColuna(colMin.ToUpper());
        int maxColIndex = ObterIndiceColuna(colMax.ToUpper());

        for (int row = 2; row <= lastRow; row++)
        {
          // Obter valores mínimo e máximo para esta linha
          ExcelRange minCell = (ExcelRange)activeSheet.Cells[row, minColIndex];
          ExcelRange maxCell = (ExcelRange)activeSheet.Cells[row, maxColIndex];

          if (minCell.Value != null && maxCell.Value != null &&
              EhNumerico(minCell.Value) && EhNumerico(maxCell.Value))
          {
            double minValue = Convert.ToDouble(minCell.Value);
            double maxValue = Convert.ToDouble(maxCell.Value);

            // Processar cada coluna no intervalo para esta linha
            for (int col = startColIndex; col <= endColIndex; col++)
            {
              ExcelRange cell = (ExcelRange)activeSheet.Cells[row, col];
              if (cell.Value != null && EhNumerico(cell.Value))
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

        return "Formatação concluída com sucesso!";
      }
      catch (Exception ex)
      {
        return $"Erro: {ex.Message}";
      }
    }

    // Método para solicitar entrada do usuário
    private static string SolicitarEntrada(string prompt, string title)
    {
      try
      {
        // Usar InputBox do Excel
        var app = (Microsoft.Office.Interop.Excel.Application)ExcelDnaUtil.Application;
        var result = app.InputBox(prompt, title, Type: 2); // Type 2 = text

        if (result is string texto)
          return texto;
        else
          return string.Empty;
      }
      catch
      {
        return string.Empty;
      }
    }

    // Converter letra de coluna em índice numérico
    private static int ObterIndiceColuna(string colLetter)
    {
      int index = 0;
      foreach (char c in colLetter.ToUpper())
      {
        index = index * 26 + (c - 'A' + 1);
      }
      return index;
    }

    // Verificar se um valor é numérico
    private static bool EhNumerico(object value)
    {
      return value is double || value is int || value is decimal || value is float;
    }
  }

  [ComVisible(true)]
  public class RibbonController : ExcelRibbon
  {
    public void OnFormatarLimites(IRibbonControl control)
    {
      var resultado = FormatacaoCondicional.FormatarPorLimites();

      // Mostrar resultado em uma MessageBox
      var app = (Microsoft.Office.Interop.Excel.Application)ExcelDnaUtil.Application;
      app.StatusBar = resultado;
    }

    public override string GetCustomUI(string RibbonID)
    {
      return @"
            <customUI xmlns='http://schemas.microsoft.com/office/2006/01/customui'>
              <ribbon>
                <tabs>
                  <tab id='tabRnzFormatting' label='RNZ Formatting'>
                    <group id='groupFormatar' label='Formatação'>
                      <button id='btnFormatarLimites' 
                              label='Formatar por Limites' 
                              size='large'
                              imageMso='ConditionalFormattingColorScales'
                              onAction='OnFormatarLimites' />
                    </group>
                  </tab>
                </tabs>
              </ribbon>
            </customUI>";
    }
  }
}