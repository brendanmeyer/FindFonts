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


#*******************************************************************
# Declare Parameters
#*******************************************************************
param(
	[Parameter(Mandatory=$true)]
	[String] $file = [String]::Empty,
    [switch] $help = $false
)

. "$(Split-Path $MyInvocation.MyCommand.Path)\Font-Common.ps1"


#*******************************************************************
# Declare Functions
#*******************************************************************

#*******************************************************************
# Function Get-RegistryStringNameFromValue()
#
# Purpose:  Return the Registry value name
#
# Input:    $keyPath    Regsitry key drive path
#           $valueData  Regsitry value sting data
#
# Returns:  Registry string value name
#
#*******************************************************************
function Get-RegistryStringNameFromValue([String] $keyPath, [String] $valueData)
{
    $pattern = [Regex]::Escape($valueData)

    foreach($property in (Get-ItemProperty $keyPath).PsObject.Properties)
    {
        ## Skip the property if it was one PowerShell added
        if(($property.Name -eq "PSPath") -or
            ($property.Name -eq "PSChildName"))
        {
            continue
        }
        ## Search the text of the property
        $propertyText = "$($property.Value)"
        if($propertyText -match $pattern)
        {
            "$($property.Name)"
        }
    }
}


#*******************************************************************
# Function Remove-SingleFont()
#
# Purpose:  Uninstall a font file
#
# Input:    $file    Font file name
#
# Returns:  0 - success, 1 - failure
#
#*******************************************************************
function Remove-SingleFont($file)
{
    try
    {
        $fontFinalPath = Join-Path $fontsFolderPath $file
        $retVal = [FontResource.AddRemoveFonts]::RemoveFont($fontFinalPath)
        if ($retVal -eq 0) {
            Write-Host "Font `'$($file)`' removal failed"
            Write-Host ""
            1
        }
        else
        {
            $fontRegistryvaluename = (Get-RegistryStringNameFromValue $fontRegistryPath $file)
            Write-Host "Font: $($fontRegistryvaluename)"
            if ($fontRegistryvaluename -ne "")
            {
                Remove-ItemProperty -path $fontRegistryPath -name $fontRegistryvaluename
            }
            Remove-Item $fontFinalPath
            if ($null -ne $error[0])
            {
                Write-Host "An error occured removing $`'$($file)`'"
                Write-Host ""
                Write-Host "$($error[0].ToString())"
                $error.clear()
            }
            else
            {
                Write-Host "Font `'$($file)`' removed successfully"
                Write-Host ""
            }
            0
        }
        ""
    }
    catch
    {
        Write-Host "An error occured removing `'$($file)`'"
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
Remove-Font.ps1
This script is used to uninstall a Windows font.

Usage:
Remove-Font.ps1 -help | -path "<Font file name>"

Parameters:

    -help
     Displays usage information.

    -file
     Font file name.  Files located in \Windows\Fonts.  Valid file 
     types are .fon, .fnt, .ttf,.ttc, .otf, .mmm, .pbf, and .pfm

Examples:
    Remove-Font.ps1
    Remove-Font.ps1 -file "MyFont.ttf"
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
function ProcessArguments()
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

    $fontFilePath = Join-Path $fontsFolderPath $file
    if ((Test-Path $fontFilePath -PathType Leaf) -eq $true)
    {
        If ($hashFontFileTypes.ContainsKey((Get-Item $fontFilePath).Extension))
        {
            $retVal = Remove-SingleFont $file
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
            "`'$($fontFilePath)`' not a valid font file type"
            ""
            exit 1
        }
    }
    else
    {
        "`'$($fontFilePath)`' not found"
        ""
        exit 1
    }
}


#*******************************************************************
# Main Script
#*******************************************************************

$fontsFolderPath = [System.Environment]::GetFolderPath($CSIDL_FONTS)
ProcessArguments

