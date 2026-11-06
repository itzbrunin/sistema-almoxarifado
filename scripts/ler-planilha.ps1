
param(
  [Parameter(Mandatory = $true)]
  [string]$CaminhoPlanilha
)

$ErrorActionPreference = "Stop"

function Converter-ParaTexto {
  param([object]$Valor)

  if ($null -eq $Valor) {
    return ""
  }

  if ($Valor -is [double] -and [math]::Floor($Valor) -eq $Valor) {
    return $Valor.ToString(
      "0",
      [System.Globalization.CultureInfo]::InvariantCulture
    )
  }

  return [string]$Valor
}

# Inicializa todas as referências COM.
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
  if (-not (Test-Path -LiteralPath $CaminhoPlanilha)) {
    throw "Planilha não encontrada: $CaminhoPlanilha"
  }

  # Inicia o Excel em segundo plano.
  $excel = New-Object -ComObject Excel.Application
  $excel.Visible = $false
  $excel.DisplayAlerts = $false

  # Abre a planilha em modo somente leitura.
  $workbooks = $excel.Workbooks
  $workbook = $workbooks.Open($CaminhoPlanilha, 0, $true)

  # Acessa a aba CONTROLE.
  $worksheets = $workbook.Worksheets
  $worksheet = $worksheets.Item("CONTROLE")

  # Localiza a última linha preenchida na coluna B.
  $rows = $worksheet.Rows
  $cells = $worksheet.Cells

  $lastCell = $cells.Item($rows.Count, 2)
  $lastUsedCell = $lastCell.End(-4162)

  $ultimaLinha = $lastUsedCell.Row

  # Lê os dados da planilha.
  $range = $worksheet.Range("A4:K$ultimaLinha")
  $dados = $range.Value2

  $ferramentas = @()

  for ($linha = 1; $linha -le $dados.GetLength(0); $linha++) {

    $patrimonio = Converter-ParaTexto $dados[$linha, 2]

    if ([string]::IsNullOrWhiteSpace($patrimonio)) {
      continue
    }

    $ferramentas += [PSCustomObject]@{
      item        = Converter-ParaTexto $dados[$linha, 1]
      patrimonio  = $patrimonio
      descricao   = Converter-ParaTexto $dados[$linha, 3]
      ordemCompra = Converter-ParaTexto $dados[$linha, 4]
      situacao    = Converter-ParaTexto $dados[$linha, 7]
      local       = Converter-ParaTexto $dados[$linha, 8]
      responsavel = Converter-ParaTexto $dados[$linha, 9]
      supervisor  = Converter-ParaTexto $dados[$linha, 10]
    }
  }

  # Retorna os dados em JSON.
  [PSCustomObject]@{
    total = $ferramentas.Count
    ferramentas = @($ferramentas)
  } | ConvertTo-Json -Depth 3 -Compress
}
finally {

  # Libera os objetos na ordem inversa de criação.
  if ($workbook -ne $null) {
    $workbook.Close($false)
  }

  if ($excel -ne $null) {
    $excel.Quit()
  }

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

  [GC]::Collect()
  [GC]::WaitForPendingFinalizers()
}

exit 0