# Forgive me father for I have sinned  
Gehshin Impact comic downloader is a tool to well... download the official Genshin Impact comic.  
The tool creates a hidden Internet Explorer browser window \*shudder\* and manipulates the browser programatically to grab the URLs for the pages of the comic.  
I wrote this as a challenge to myself, however, upon completion of this most unholy of tools I experienced a sudden wave of regret and felt it would be better to send it back into the dark. Alas, I did not and so here it is to darken your day.  

## How to use  
I'm setting the minimum version of PowerShell to 5.1 as I have not tested against any other versions.  

Most common ways to run the tool:  
 - Right click on the script file > Run with PowerShell.  
 - From a PowerShell window. Do note that by default PowerShell is set to not allow the running of scripts which will need to be temporarily disabled.  
```
Set-ExecutionPolicy Bypass -Scope Process -Force
{drive}:\{path}\{to}\{script}\Get-GenshinComic.ps1
```

By default the script will download the comic to the location the script was ran. You can also provide a folder you'd like to save it in.
```
{drive}:\{path}\{to}\{script}\Get-GenshinComic.ps1 -outputFolder 'S:\Private Stash'
```
---
### Notes  
Was not tested against the non-English langauge versions of the comic. Simply changinging the URL to the desired languange version in the script *may* do the trick, but I make no promises. Additionally, I don't take any steps to verify that the file/folder names are acceptible to Windows.  

Tool downloads images one by one rather than asynchronously as I didn't take the time to develop teh cability to do so. Also, I thought the progress bar was pretty, but mostly the former.  

Tool does not currently have the ability specific chapters.  

### Seal of approval  
![seal of approval](https://i.imgur.com/SkCMOti.png)
