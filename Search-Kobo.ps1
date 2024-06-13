using module .\KoboAPI.psm1

param(
    [Parameter(Mandatory)][string]$Title,
    [string]$Genre,
    [string]$Keyword,
    [Parameter(Mandatory)][string]$AppId
)

$api = [KoboAPI]::New()
$api.Init($AppId)
$api.SearchKobo($Title, $Genre, $KeyWords)
