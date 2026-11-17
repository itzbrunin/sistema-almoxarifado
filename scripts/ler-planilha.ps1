
param(
  [Parameter(Mandatory = $true)]
  [string]$CaminhoPlanilha
)

$ErrorActionPreference = "Stop"

# ============================================================
# 1. FUNÇÃO PARA CONVERTER VALORES EM TEXTO
# ============================================================

function Convert-ParaTexto {
  param([object]$Valor)

  if ($null -eq $Valor) {
    return ""
  }

  # Mantém números inteiros sem casas decimais.
  if ($Valor -is [double] -and [math]::Floor($Valor) -eq $Valor) {
    return $Valor.ToString(
      "0",
      [System.Globalization.CultureInfo]::InvariantCulture
    )
  }

  return [string]$Valor
}

# ============================================================
# 2. INICIALIZAÇÃO DAS REFERÊNCIAS COM
# ============================================================

$excel = $null
$workbooks = $null
$workbook = $null
$worksheets = $null
$worksheet = $null
$rows = $null
$cells = $null
$lastCell = $null
$lastUsedCell = $null
$range = $null

try {

  # ==========================================================
  # 3. VERIFICA SE A PLANILHA EXISTE
  # ==========================================================

  if (-not (Test-Path -LiteralPath $CaminhoPlanilha)) {
    throw "Planilha não encontrada: $CaminhoPlanilha"
  }

  # ==========================================================
  # 4. INICIA O EXCEL EM SEGUNDO PLANO
  # ==========================================================

  $excel = New-Object -ComObject Excel.Application
  $excel.Visible = $false
  $excel.DisplayAlerts = $false

  # ==========================================================
  # 5. ABRE A PLANILHA EM MODO SOMENTE LEITURA
  # ==========================================================

  $workbooks = $excel.Workbooks
  $workbook = $workbooks.Open($CaminhoPlanilha, 0, $true)

  # ==========================================================
  # 6. ACESSA A ABA CONTROLE
  # ==========================================================

  $worksheets = $workbook.Worksheets
  $worksheet = $worksheets.Item("CONTROLE")

  # ==========================================================
  # 7. LOCALIZA A ÚLTIMA LINHA PREENCHIDA NA COLUNA B
  # ==========================================================

  $rows = $worksheet.Rows
  $cells = $worksheet.Cells

  $lastCell = $cells.Item($rows.Count, 2)

  # -4162 representa xlUp no Excel.
  $lastUsedCell = $lastCell.End(-4162)

  $ultimaLinha = $lastUsedCell.Row

  # ==========================================================
  # 8. LÊ OS DADOS DA PLANILHA
  # ==========================================================

  $range = $worksheet.Range("A4:K$ultimaLinha")
  $dados = $range.Value2

  $ferramentas = @()

  # ==========================================================
  # 9. PERCORRE OS REGISTROS
  # ==========================================================

  for ($linha = 1; $linha -le $dados.GetLength(0); $linha++) {

    $patrimonio = Convert-ParaTexto $dados[$linha, 2]

    # Ignora linhas sem patrimônio, conforme a lógica original.
    if ([string]::IsNullOrWhiteSpace($patrimonio)) {
      continue
    }

    # Monta o objeto da ferramenta.
    $ferramentas += [PSCustomObject]@{
      item        = Convert-ParaTexto $dados[$linha, 1]
      patrimonio  = $patrimonio
      descricao   = Convert-ParaTexto $dados[$linha, 3]
      ordemCompra = Convert-ParaTexto $dados[$linha, 4]
      situacao    = Convert-ParaTexto $dados[$linha, 7]
      local       = Convert-ParaTexto $dados[$linha, 8]
      responsavel = Convert-ParaTexto $dados[$linha, 9]
      supervisor  = Convert-ParaTexto $dados[$linha, 10]
    }
  }

  # ==========================================================
  # 10. RETORNA OS DADOS EM JSON
  # ==========================================================

  [PSCustomObject]@{
    total = $ferramentas.Count
    ferramentas = @($ferramentas)
  } | ConvertTo-Json -Depth 3 -Compress

}
finally {

  # ==========================================================
  # 11. FECHA A PLANILHA E ENCERRA O EXCEL
  # ==========================================================

  if ($null -ne $workbook) {
    $workbook.Close($false)
  }

  if ($null -ne $excel) {
    $excel.Quit()
  }

  # ==========================================================
  # 12. LIBERA OS OBJETOS COM
  # ==========================================================

  $objetosCOM = @(
    $range,
    $lastUsedCell,
    $lastCell,
    $cells,
    $rows,
    $worksheet,
    $worksheets,
    $workbook,
    $workbooks,
    $excel
  )

  foreach ($objeto in $objetosCOM) {

    if ($null -ne $objeto) {

      try {
        [void][System.Runtime.InteropServices.Marshal]::FinalReleaseComObject($objeto)
      }
      catch {
        [Console]::Error.WriteLine(
          "Falha ao liberar objeto COM: $($_.Exception.Message)"
        )
      }
    }
  }

  # ==========================================================
  # 13. COLETA DE LIXO
  # ==========================================================

  [GC]::Collect()
  [GC]::WaitForPendingFinalizers()
}

exit 0
