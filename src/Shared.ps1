<#
    .SYNOPSIS
    Defines shared variables.

    .DESCRIPTION
    See synopsis.
#>

$SourceDir = $PSCommandPath | Split-Path -Parent
$RepoRoot = $SourceDir | Split-Path -Parent
$EpplusLibName = "EPPlus.dll"
$EpplusLibPath = $SourceDir | Join-Path -ChildPath $EpplusLibName
