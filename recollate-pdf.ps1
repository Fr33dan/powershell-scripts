<#
  .SYNOPSIS
  Converts a scanned PDF from to a print and fold format.

  .DESCRIPTION
  The recollate-pdf.ps1 script re-collates a PDF of a book scanned to be read on a
  screen to a format that can be printed and folded back into a facsimile of 
  the original book. The source PDF must have the first page be the front cover,
  the last page be the back cover and the pages in between each be a spread of
  two pages of the original book. Requires Imagemagick and ghostscript. 
  Additionally there is a speed boost if pdfinfo is available in the path as well.

  .PARAMETER SourcePDF
  Source PDF to re-collate to a printable format.

  .PARAMETER DestPDF
  Specifies the name and path for the generated foldable PDF.
  
  .PARAMETER ThreadCount
  Number of threads to use when processing pages.  -1 (Default) is the CPU core count.

  .INPUTS
  None. You cannot pipe objects to recollate-pdf.ps1.

  .OUTPUTS
  None. recollate-pdf.ps1 does not generate any output.

  .EXAMPLE
  PS> .\recollate-pdf.ps1 C:\ScannedPDF.pdf
  
  Simple conversion

  .EXAMPLE
  PS> .\recollate-pdf.ps1 -SourcePDF C:\ScannedPDF.pdf -DestPDF C:\PrintablePDF.pdf
  
  Specify output file.

  .EXAMPLE
  PS> .\recollate-pdf.ps1 -SourcePDF C:\ScannedPDF.pdf -ThreadCount 2
  
  Limit thread count (not recommended)
  
  .LINK
  https://imagemagick.org/
  
  .LINK
  https://www.ghostscript.com/
  
  .LINK
  https://www.xpdfreader.com/pdfinfo-man.html
#>

param ([Parameter(Mandatory=$true)][string]$SourcePDF
	, [string]$DestPDF = $((Get-Item $SourcePDF).BaseName + "-foldable.pdf")
	, [int]$ThreadCount = -1)

if(Test-Path -Path $DestPDF -PathType Leaf){
	$confirmation = Read-Host "$DestPDF already exists overwrite? [y/N]"
	if($confirmation -ne "y")
	{
		exit;
	}
}

if( $ThreadCount -le 0){
	$ThreadCount = (
	 (Get-CimInstance –ClassName Win32_Processor).NumberOfCores |
	   Measure-Object -Sum
	).Sum
}

$tempFolderName = "pdf-repage-$([guid]::NewGuid())"
$tempDir = Join-Path -Path $env:TEMP -ChildPath $tempFolderName
$pdfPagesPath = Join-Path -Path $tempDir -ChildPath "PDF";
$splitFilesPath = Join-Path -Path $tempDir -ChildPath "Split";
$finalPagesPath = Join-Path -Path $tempDir -ChildPath "Final";


$null = New-Item -Path $env:TEMP -Name $tempFolderName -ItemType "directory";

try {
	$null = New-Item -Path $tempDir -Name "PDF" -ItemType "directory";
	$null = New-Item -Path $tempDir -Name "Split" -ItemType "directory";

	if (Get-Command "pdfinfo" -ErrorAction SilentlyContinue) 
	{ 
		$sourcePageCount = (pdfinfo $SourcePDF | Select-String -Pattern '(?<=Pages:\s*)\d+').Matches.Value
	} else {
		$sourcePageCount = [int]((magick identify -format "%n\n" $SourcePDF) | Select-Object -First 1)
	}
	$sourcePageRange = 0 .. ($sourcePageCount - 1)

	$pageCount = ($sourcePageCount - 1) * 2;

	$sourcePageRange | ForEach-Object -ThrottleLimit $ThreadCount -Parallel {
		$pdfFileName = "s-$_.jpg"
		$pdfPageFile = Join-Path -Path $using:pdfPagesPath -ChildPath $pdfFileName
		magick convert -density 150 "$using:SourcePDF[$_]" -quality 90 $pdfPageFile
		
		if($_ -eq 0){
			$moveDest = "$using:splitFilesPath\0000.jpg"
			Move-Item -Path $pdfPageFile -Destination $moveDest
		} elseif ($_ -eq ($using:sourcePageCount - 1)){
			$moveDest = "$using:splitFilesPath\$(($using:pageCount - 1).ToString('0000')).jpg"
			Move-Item -Path $pdfPageFile -Destination $moveDest
		} else {
			$destName = Join-Path -Path $using:splitFilesPath -ChildPath $pdfFileName;
			magick convert -crop 50%x100% +repage $pdfPageFile $($destName + ".jpg");
			
			$firstPage = 1 + (($_ - 1) * 2);
			$secondPage = $firstPage + 1;
				
			$firstPage= $firstPage.ToString("0000") + ".jpg";
			$secondPage = $secondPage.ToString("0000") + ".jpg";
			
			Rename-Item $($destName + "-0.jpg") $firstPage;
			Rename-Item $($destName + "-1.jpg") $secondPage;
		}
	};

	$null = New-Item -Path $tempDir -Name "Final" -ItemType "directory";


	$rawPageRange = 0 .. ($sourcePageCount - 2)

	$rawPageRange | ForEach-Object -ThrottleLimit $ThreadCount -Parallel {
		$destFile = $_.ToString("0000") + ".jpg";
		$destFile = Join-Path -Path $using:finalPagesPath -ChildPath $destFile;
		
		if(($_ % 2) -eq 0) {
			$firstPage = $using:pageCount - $_ - 1;
			$secondPage = $_;
		}else {
			$firstPage = $_;
			$secondPage = $using:pageCount - $_ - 1;
		}
		
		$firstPage = $firstPage.ToString("0000") + ".jpg";
		$firstPage = Join-Path -Path $using:splitFilesPath -ChildPath $firstPage;
		
		$secondPage = $secondPage.ToString("0000") + ".jpg";
		$secondPage = Join-Path -Path $using:splitFilesPath -ChildPath $secondPage;
		
		magick convert $firstPage $secondPage +append $destFile;
	};

	magick convert "$finalPagesPath\*.jpg" -quality 90 $DestPDF;
} finally
{
	Remove-Item -Path $tempDir -Recurse
}
