<#
.SYNOPSIS
    Exporta o e-CPF A1 do repositorio de certificados do Windows para um .pfx,
    para que o container Docker possa usa-lo.

.DESCRIPTION
    O container Linux nao enxerga o repositorio de certificados do Windows.
    Este script lista os certificados pessoais e exporta o escolhido para
    .\certs\certificado.pfx (pasta ja ignorada pelo git e pelo build Docker).

    So funciona se a chave privada tiver sido marcada como exportavel na
    instalacao. Se nao for o caso, reinstale o A1 a partir do arquivo original
    marcando "Marcar esta chave como exportavel".

.EXAMPLE
    powershell -ExecutionPolicy Bypass -File docker\exportar_certificado.ps1
#>

$ErrorActionPreference = 'Stop'
$destino = Join-Path $PSScriptRoot '..\certs'
$destino = [System.IO.Path]::GetFullPath($destino)
New-Item -ItemType Directory -Force -Path $destino | Out-Null

Write-Host "`nCertificados pessoais instalados:`n" -ForegroundColor Cyan

$certs = @(Get-ChildItem -Path Cert:\CurrentUser\My |
    Where-Object { $_.HasPrivateKey -and $_.NotAfter -gt (Get-Date) })

if ($certs.Count -eq 0) {
    Write-Host "Nenhum certificado com chave privada valido encontrado." -ForegroundColor Red
    exit 1
}

for ($i = 0; $i -lt $certs.Count; $i++) {
    $c = $certs[$i]
    $exportavel = if ($c.PrivateKey -and $c.PrivateKey.CspKeyContainerInfo.Exportable) { "exportavel" } else { "? (tente)" }
    "{0}) {1}`n   valido ate {2:dd/MM/yyyy}  [{3}]" -f $i, $c.Subject, $c.NotAfter, $exportavel | Write-Host
}

$escolha = Read-Host "`nNumero do certificado a exportar"
$cert = $certs[[int]$escolha]

Write-Host "`nDefina uma senha para o arquivo .pfx (pode ser diferente da senha original)."
$senha = Read-Host -AsSecureString "Senha do .pfx"

$caminho = Join-Path $destino 'certificado.pfx'
Export-PfxCertificate -Cert $cert -FilePath $caminho -Password $senha -Force | Out-Null

Write-Host "`nExportado para: $caminho" -ForegroundColor Green
Write-Host "Essa pasta esta no .gitignore e no .dockerignore - o .pfx nao vai para o git nem para a imagem." -ForegroundColor DarkGray
Write-Host "`nAgora rode:  docker compose run --rm sonda /certs/certificado.pfx`n" -ForegroundColor Cyan
