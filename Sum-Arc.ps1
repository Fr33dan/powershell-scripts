<#
  .SYNOPSIS
  Get the total uncompressed size of a list of archive files using 7z

  .PARAMETER ArchiveListFile
  A list of archive file path(s) to summarize. Paths may be escaped with quotes or not.

  .PARAMETER FolderPath
  Path to a folder containing archive file(s) to be summarized.
  
  .PARAMETER ArchiveExtension
  Extension of archive file(s) to summarize.
  
  .PARAMETER ArchiveListFile
  File containing list of containing path(s) to archive file(s).

  .EXAMPLE
  PS> Sum-Arc -FolderPath C:\MyCollectionOfArchives\
  PS> Sum-Arc C:\MyCollectionOfArchives\
  
  Summary based on folder location

  .EXAMPLE
  PS> Sum-Arc -ArchiveListFile C:\MyCollectionOfArchives\MyListOfFiles.txt
  
  Load the list of archive file(s) from a text file.
  
  .EXAMPLE
  PS> @myListOfArchives | Sum-Arc
  
  Summarize list of files from pipeline input.
  
  .LINK
  https://www.7-zip.org/
#>
[CmdletBinding(DefaultParameterSetName = 'WithPath')]
Param(
	[Parameter(Mandatory,
    ParameterSetName = "WithPath",
	HelpMessage = "Path to a folder containing 7z archives.",
	Position = 0)]
	[string]
	$FolderPath = $null,
	
	[Parameter(ParameterSetName = "WithPath",
	HelpMessage = "Text file containing path(s) to 7z file(s).")]
	[string]
	$ArchiveExtension = "7z",
	
	[Parameter(Mandatory,
    ParameterSetName = "WithList",
	HelpMessage = "List of containing path(s) to archive file(s).",
	ValueFromPipeline)]
	[string[]]
	$ArchiveFile = $null,
	
	[Parameter(Mandatory,
    ParameterSetName = "WithListFile",
	HelpMessage = "File containing list of containing path(s) to archive file(s).")]
	[string[]]
	$ArchiveListFile = $null
)
Begin
{
	$uncompressedHeader = 'Uncompressed Size (GB)'
	$uncompressedSum = 0
	
	$compressedHeader = 'Compressed Size (GB)'
	$compressedSum = 0
	$list = @()
	if ($FolderPath -ne "") 
	{
		Write-Host "Getting file List"
		$ArchiveFile = Get-ChildItem $FolderPath -Filter "*.$ArchiveExtension"
	}
	elseif ($ArchiveListFile -ne "")
	{
		$ArchiveFile = Get-Content $ArchiveListFile
	}
}
Process
{
	# If using pipeline parameter input only process one item.
	$count = if ($FolderPath -ne $null) { $ArchiveFile.length } else { 1 }
	for($i = 0; $i -lt $count;$i++)
	{ 
		$pathItem = $ArchiveFile[$i]
		if ($pathItem -is [string] -and $pathItem.StartsWith('"'))
		{
			$pathItem = $pathItem.Substring(1,$pathItem.Length - 2)
		}
		
		$archivePath = Get-Item -LiteralPath "$pathItem"
		$7zOutput = 7z l "$archivePath"
		$outputRow = $7zOutput.length - 1
		if ($7zOutput[$outputRow].StartsWith("Warnings"))
		{
			$outputRow -= 2
		}
		
		$totalBytes = $7zOutput[$outputRow].Substring(27,12)
		$uncompressedSize = $totalBytes / [Math]::Pow(1024, 3)
		
		$totalBytes = $7zOutput[$outputRow].Substring(40,12)
		$compressedSize = $totalBytes / [Math]::Pow(1024, 3)
		
		$uncompressedSum += $uncompressedSize;
		$compressedSum += $compressedSize;
		
		new-object psobject -Property @{File = $archivePath.Name;
										$compressedHeader = $compressedSize
										$uncompressedHeader = $uncompressedSize}
	}
}
End
{
	new-object psobject -Property @{File="----"
									$compressedHeader=("-" * $compressedHeader.Length)
									$uncompressedHeader=("-" * $uncompressedHeader.Length)}
	new-object psobject -Property @{File = "Total"
									$compressedHeader = $compressedSum
									$uncompressedHeader = $uncompressedSum}
}
