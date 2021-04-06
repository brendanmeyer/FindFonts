#https://github.com/wormeyman/FindFonts
#########################################################################################
#   MICROSOFT LEGAL STATEMENT FOR SAMPLE SCRIPTS/CODE
#########################################################################################
#   This Sample Code is provided for the purpose of illustration only and is not 
#   intended to be used in a production environment.
#
#   THIS SAMPLE CODE AND ANY RELATED INFORMATION ARE PROVIDED "AS IS" WITHOUT WARRANTY 
#   OF ANY KIND, EITHER EXPRESSED OR IMPLIED, INCLUDING BUT NOT LIMITED TO THE IMPLIED 
#   WARRANTIES OF MERCHANTABILITY AND/OR FITNESS FOR A PARTICULAR PURPOSE.
#
#   We grant You a nonexclusive, royalty-free right to use and modify the Sample Code 
#   and to reproduce and distribute the object code form of the Sample Code, provided 
#   that You agree: 
#   (i)    to not use Our name, logo, or trademarks to market Your software product 
#          in which the Sample Code is embedded; 
#   (ii)   to include a valid copyright notice on Your software product in which 
#          the Sample Code is embedded; and 
#   (iii)  to indemnify, hold harmless, and defend Us and Our suppliers from and 
#          against any claims or lawsuits, including attorneys’ fees, that arise 
#          or result from the use or distribution of the Sample Code.
#########################################################################################

#******************************************************************************
# File:     Add-Font.ps1
# Date:     08/28/2013
# Version:  1.0.1
#
# Purpose:  PowerShell script to install Windows fonts.
#
# Usage:    Add-Font -help | -path "<Font file or folder path>"
#
# Copyright (C) 2010 Microsoft Corporation
#
#
# Revisions:
# ----------
# 1.0.0   09/22/2010   Created script.
# 1.0.1   08/28/2013   Fixed help text.  Added quotes around paths in messages.
# 1.0.2   06/04/2021   Split common code to separate file
#
#******************************************************************************

#requires -Version 2.0

. $scriptPath\Font-Common.ps1

#*******************************************************************
# Declare Parameters
#*******************************************************************
param(
    [string] $path = "",
    [switch] $help = $false
)

#*******************************************************************
# Declare Functions
#*******************************************************************


#*******************************************************************
# Function Add-SingleFont()
#
# Purpose:  Install a font file
#
# Input:    $file    Font file path
#
# Returns:  0 - success, 1 - failure
#
#*******************************************************************
function Add-SingleFont($filePath)
{
    try
    {
        [string]$filePath = (resolve-path $filePath).path
        [string]$fileDir  = split-path $filePath
        [string]$fileName = split-path $filePath -leaf
        [string]$fileExt = (Get-Item $filePath).extension
        [string]$fileBaseName = $fileName -replace($fileExt ,"")

        $shell = new-object -com shell.application
        $myFolder = $shell.Namespace($fileDir)
        $fileobj = $myFolder.Items().Item($fileName)
        $fontName = $myFolder.GetDetailsOf($fileobj,21)

        if ($fontName -eq "") { $fontName = $fileBaseName }

        copy-item $filePath -destination $fontsFolderPath

        $fontFinalPath = Join-Path $fontsFolderPath $fileName
        $retVal = [FontResource.AddRemoveFonts]::AddFont($fontFinalPath)

        if ($retVal -eq 0) {
            Write-Host "Font `'$($filePath)`'`' installation failed"
            Write-Host ""
            1
        }
        else
        {
            Write-Host "Font `'$($filePath)`' installed successfully"
            Write-Host ""
            Set-ItemProperty -path "$($fontRegistryPath)" -name "$($fontName)$($hashFontFileTypes.item($fileExt))" -value "$($fileName)" -type STRING
            0
        }
        ""
    }
    catch
    {
        Write-Host "An error occured installing `'$($filePath)`'"
        Write-Host ""
        Write-Host "$($error[0].ToString())"
        Write-Host ""
        $error.clear()
        1
    }
}


#*******************************************************************
# Function Show-Usage()
#
# Purpose:   Shows the correct usage to the user.
#
# Input:     None
#
# Output:    Help messages are displayed on screen.
#
#*******************************************************************
function Show-Usage()
{
$usage = @'
Add-Font.ps1
This script is used to install Windows fonts.

Usage:
Add-Font.ps1 -help | -path "<Font file or folder path>"

Parameters:

    -help
     Displays usage information.

    -path
     May be either the path to a font file to install or the path to a folder 
     containing font files to install.  Valid file types are .fon, .fnt,
     .ttf,.ttc, .otf, .mmm, .pbf, and .pfm

Examples:
    Add-Font.ps1
    Add-Font.ps1 -path "C:\Custom Fonts\MyFont.ttf"
    Add-Font.ps1 -path "C:\Custom Fonts"
'@

$usage
}


#*******************************************************************
# Function Process-Arguments()
#
# Purpose: To validate parameters and their values
#
# Input:   All parameters
#
# Output:  Exit script if parameters are invalid
#
#*******************************************************************
function Process-Arguments()
{
    ## Write-host 'Processing Arguments'

    if ($unnamedArgs.Length -gt 0)
    {
        write-host "The following arguments are not defined:"
        $unnamedArgs
    }

    if ($help -eq $true) 
    { 
        Show-Usage
        break
    }

    if ((Test-Path $path -PathType Leaf) -eq $true)
    {
        If ($hashFontFileTypes.ContainsKey((Get-Item $path).Extension))
        {
            $retVal = Add-SingleFont $path
            if ($retVal -ne 0)
            {
                exit 1
            }
            else
            {
                exit 0
            }
        }
        else
        {
            "`'$($path)`' not a valid font file type"
            ""
            exit 1
        }
    }
    elseif ((Test-Path $path -PathType Container) -eq $true)
    {
        $bErrorOccured = $false
        foreach($file in (Get-Childitem $path))
        {

            if ($hashFontFileTypes.ContainsKey($file.Extension))
            {
                $retVal = Add-SingleFont (Join-Path $path $file.Name)
                if ($retVal -ne 0)
                {
                    $bErrorOccured = $true
                }
            }
            else
            {
                "`'$(Join-Path $path $file.Name)`' not a valid font file type"
                ""
            }
        }

        If ($bErrorOccured -eq $true)
        { 
            exit 1 
        }
        else
        {
            exit 0
        }
    }
    else
    {
        "`'$($path)`' not found"
        ""
        exit 1
    }
}


#*******************************************************************
# Main Script
#*******************************************************************

$fontsFolderPath = Get-SpecialFolder($CSIDL_FONTS)
Process-Arguments

