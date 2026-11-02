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
    return $Valor.ToString("0", [System.Globalization.CultureInfo]::InvariantCulture)
  }

  return [string]$Valor
}

$excel = $null
$workbook = $null
$worksheet = $null

try {
  if (-not (Test-Path -LiteralPath $CaminhoPlanilha)) {
    throw "Planilha não encontrada: $CaminhoPlanilha"
  }

  $excel = New-Object -ComObject Excel.Application
  $excel.Visible = $false
  $excel.DisplayAlerts = $false

  $workbook = $excel.Workbooks.Open($CaminhoPlanilha, $null, $true)
  $worksheet = $workbook.Worksheets.Item("CONTROLE")
  $xlUp = -4162
  $ultimaLinha = $worksheet.Cells.Item($worksheet.Rows.Count, 2).End($xlUp).Row
  $dados = $worksheet.Range("A4:K$ultimaLinha").Value2
  $ferramentas = @()

  for ($linha = 1; $linha -le $dados.GetLength(0); $linha += 1) {
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

  [PSCustomObject]@{
    total       = $ferramentas.Count
    ferramentas = @($ferramentas)
  } | ConvertTo-Json -Depth 3 -Compress
}
finally {
  if ($worksheet -ne $null) {
    [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($worksheet)
  }

  if ($workbook -ne $null) {
    $workbook.Close($false)
    [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($workbook)
  }

  if ($excel -ne $null) {
    $excel.Quit()
    [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($excel)
  }

  [GC]::Collect()
  [GC]::WaitForPendingFinalizers()
}
