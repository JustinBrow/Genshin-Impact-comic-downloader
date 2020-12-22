#Requires -Version 5.1
param (
   $outputFolder = $PSScriptRoot
)

$ErrorActionPreference = 'Stop'

if (-not ($outputFolder -match '\\$'))
{
   $outputFolder = $outputFolder + '\'
}
$outputFolder = $outputFolder + 'Genshin Impact Official Manga'

if (-not (Test-Path $outputFolder))
{
   [void](New-Item -Path $outputFolder -ItemType Directory)
}

$URL = 'https://genshin.mihoyo.com/en/manga/'
$IE = New-Object -ComObject 'InternetExplorer.Application'
$IE.FullScreen = $false
$IE.Visible = $false
$IE.navigate($URL)

do {sleep 1}
until ($IE.Busy -eq $False -and $IE.ReadyState -eq 4)

$chapters = @()
sleep 2

try
{
   $navButtons = $IE.Document.getElementsByClassName('mihoyo-pager-rich__button')
}
catch
{
   try # No idea what HRESULT: 0x800A01B6 means so we're going to reload the browser and try again.
   {
      $IE.Quit()
      $IE = New-Object -ComObject 'InternetExplorer.Application'
      $IE.navigate($URL)

      do {sleep 1}
      until ($IE.Busy -eq $False -and $IE.ReadyState -eq 4)

      sleep 2
      $navButtons = $IE.Document.getElementsByClassName('mihoyo-pager-rich__button')
   }
   catch
   {
      $IE.Quit()
      throw 'Try again later'
   }
}
$navButtons = $navButtons | Where-Object {$_.textContent -ne $null}

$navCount = ($navButtons | Select-Object textContent | Sort-Object textContent).textContent[-1]

do
{
   [int]$navNumber = ($IE.Document.getElementsByClassName('mihoyo-pager-rich__button mihoyo-pager-rich__current') | Select-Object textContent).textContent

   $chapters += $IE.Document.getElementsByClassName('chapters__item') | ForEach-Object {
      [void]($_.innerHTML -match '(detail\/\d{2,}).*?(https:\/\/.*?\.jpg).*?chapters__number">(.*?)<\/div>.*')
      [PSCustomObject]@{URL = $URL + $Matches.1; Thumbnail = $Matches.2;Title = $Matches.3 -replace ': ', ' - '}
   }

   $navNumber++

   try
   {
      &{($navButtons | Where-Object {$_.textcontent -eq $navNumber}).Click()} 2> $null
   }
   catch
   {
      continue
   }

   sleep 5

} until ($navNumber -gt $navCount)

$i = 1
$countChapters = $chapters.count

foreach ($chapter in $chapters)
{
   Write-Progress -Id 1 -Activity "Downloading chapter $i of $countChapters" -PercentComplete ($i/$countChapters * 100)

   if (-not (Test-Path "$outputFolder\$($chapter.Title)"))
   {
      [void](New-Item -Path $outputFolder -Name $chapter.Title -ItemType Directory)
   }

   Invoke-WebRequest -Uri $chapter.Thumbnail -OutFile "$outputFolder\$($chapter.Title)\chapter.$($chapter.Thumbnail.Split('.')[-1])"

   $IE.navigate($chapter.URL)

   do {sleep 1}
   until ($IE.Busy -eq $False -and $IE.ReadyState -eq 4)

   $pages = @()
   $pageNumber = 1
   sleep 2

   $pages += $IE.Document.getElementsByClassName('swiper-slide') | ForEach-Object {
      [void]($_.innerHTML -match '(https:\/\/.*?\.(?:png|jpg))')
      [PSCustomObject]@{PageNumber = $pageNumber; URL = $Matches.1}
      $pageNumber++
   }

   $j = 1
   $countPages = $pages.count

   foreach ($page in $pages)
   {
      Write-Progress -Id 2 -Activity "Downloading page $j of $countPages" -PercentComplete ($j/$countPages * 100)

      Invoke-WebRequest -Uri $page.URL -OutFile "$outputFolder\$($chapter.Title)\pg$($page.PageNumber).$($page.URL.Split('.')[-1])"
      sleep 1
      $j++
   }

   $i++
}
$IE.Quit()