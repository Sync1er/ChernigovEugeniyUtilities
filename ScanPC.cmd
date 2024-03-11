start "ScanPC" /D "%~dp0" powershell.exe "$comp="""%1"""; $a= Get-Content '.\%~nx0' -Encoding Oem | Select-Object -Skip 3; Invoke-Expression ($a -join """`r`n""")" 
EXIT /B

# https://github.com/Sync1er/ChernigovEugeniyUtilites/
# Chernigov Eugeniy 2024. For non commercial use only

[void] [System.Reflection.Assembly]::LoadWithPartialName("System.Windows.Forms")
Add-Type -Name Window -Namespace Console -MemberDefinition '
[DllImport("Kernel32.dll")]
public static extern IntPtr GetConsoleWindow();

[DllImport("user32.dll")]
public static extern bool ShowWindow(IntPtr hWnd, Int32 nCmdShow);
'
[Console.Window]::ShowWindow(( $consolePtr = [Console.Window]::GetConsoleWindow()), 0)

Function Prompt {"::"}

if(-not $comp){$comp=$env:COMPUTERNAME}
$jobCount=4
$excl="GENUINEINTEL","PRINTENUM","MSRRAS","}\"
$hide="RAS Async Adapter","WAN Miniport","Virtual Adapter","FAX","Microsoft XPS","OneNote","Š®à­¥¢ ï ®ç¥à¥¤ì ¯¥ç â¨","PDF","Microsoft Kernel Debug Network Adapter","HTREE"
$repl="Intel(R) Core(TM) ","11th Gen Intel(R) Core(TM) ","Intel(R) ","System Product Name"," 82371AB/EB PCI Bus Master"," To be filled by O.E.M."," Virtual Platform"," Express Chipset Family"," Express Chipset","\(ª®à¯®à æ¨ï Œ ©ªà®á®äâ - WDDM 1.1\)"," ¡®à ¬¨ªà®áå¥¬ ","‘¥¬¥©áâ¢® ­ ¡®à®¢ ¬¨ªà®áå¥¬ ","/C200 Series 6 Port","6th Generation Core Processor Family Platform "," Family Controller","APU with Radeon(tm) HD Graphics"," Storage Controllers"," Universal Host Controller","\(Œ ©ªà®á®äâ\)","Series/C200 Series Chipset Family USB ","Š®à­¥¢®© ","®á«¥¤®¢ â¥«ì­ë© ¯®àâ "," áè¨à¥­­ë© "," áè¨àï¥¬ë© ","‘â ­¤ àâ­ë© "," Enhanced Host Controller"," Storage Controller"," Chipset Family"," LaserJet Professional"
$sprtr="-------------------------------------------------"
$list= @()
$SlctnStrt= 0
$global:usrSz= @()
function txt($txt) {$ListBox3.AppendText("$txt")}

function Output($txt="`r`n", $color="black", $font=$font){
  $start=$ListBox3.TextLength
  $ListBox3.AppendText($txt)
  $ListBox3.SelectionStart = $start
  $ListBox3.SelectionLength = $txt.Length
  $ListBox3.SelectionFont = $font
  $ListBox3.SelectionColor = $color
  $start=$ListBox3.TextLength
  $ListBox3.AppendText('')
  $ListBox3.SelectionLength = 0
  $ListBox3.SelectionColor = 'Black'
}

$help="

Buttons:
 Devices - show devices
 Scan    - show hardware information and users

Scan option:
 Print - show printers
 Size  - Scan and show system and users files sizes 
 HW    - show CPU, memory, HDD, vendor
 Usr   - chek users now logged-in and *last logon (Gray: full list)

*last logon time not always correct

CR button - clean C:\`$Recycle.Bin by delete files older than 2 month
CR button shown than size of C:\`$Recycle.Bin more than 500MB
 






 https://github.com/Sync1er/ChernigovEugeniyUtilites/
 Chernigov Eugeniy 2024. For non commercial use only


"

$ScriptBlock= {
   $usrDir="$input" -split ';'
  $fldrs='Downloads;Pictures;Desktop;Videos;Documents;Music;AppData\Local\Temp;AppData\Local\Microsoft\Outlook' -split ';'       
  $tmp=$fldrs | ForEach {
    $a=$($(Get-ChildItem -Path "$($usrDir -join '\')\$_" -Recurse -Force | ForEach {$_.Length}) 2>$null  | Measure-Object  -Sum).Sum / 1MB
    if($a -gt 100){@{usrDir=$usrDir[1];Dir=$_;Size=$a}}
  }
    if($tmp){$tmp}
}

function PrProgres {  $job=$((Get-Job | Where { $_.State -eq "Running"}).count)
  if($job){$job="J$job"}else{$job=""}
  $Label.Text= "$job  $prcnt"
}

function OutputSize ($usrSz){
  ($usrSz.usrDir | Sort-Object -Unique | ForEach {$usrDir=$_; @{usrDir=$usrDir;Sum= $(($usrSz | Where {$_.usrDir -eq $usrDir}).Size | Measure-Object -Sum).Sum}} | `
` Sort-Object {$_.sum} -Descending).usrDir | ForEach {
    $usrDir=$_
    $out= $usrSz | Where {$_.usrDir -eq $usrDir} | ForEach {"`r`n"+$($_.Dir).PadRight(22," ")+$(" {0:n0}" -f $_.Size+" MB").PadLeft(10," ")}
    $usrSz | Where {$_.usrDir -eq $usrDir} | ForEach { $global:sum+=$_.Size}
    $usr= $null; $usr= $usrs | Where {$_.Login -eq $usrDir} | Select-Object -First 1
    $dep=(($usr.DistinguishedName -split ',OU=')[1] -split ',')[0]
    $fullname= Shrt "$($usr.fullname)"
    $clr="SlateBlue"; if(-not($usr.Enabled) -and $usr.fullname){$clr="red"}
    $txt1=$(addSp "$usrDir $($usr.LastLogin)"  22)
    $txt2=''; if($dep){$txt2="/$dep"}
    $bar=(0..("$txt1$fullname$txt2".Length + 3) | ForEach {'-'}) -join ''
    Output "`n$txt1 "
    Output "$fullname" $clr
    Output $("$txt2`r`n$bar-$out`r`n`r`n")
  }
}

$step=0
$tmrTick=10
if($timer){$timer.Dispose()}
$timer = New-Object System.Windows.Forms.Timer
$timer.Interval = 200
$timer.add_tick({

  $prcnt=""
  if($step -eq 6){
    if($sum){
      $sum=[int](($sum | Measure-Object  -Sum).Sum /1024)
      Output "`n`n  All summary - $sum GB"
    }
    Output "`n`n -= Done =-`n`n"; $global:step=$false; $timer.Stop()
  }

  if($step -eq 5){    
    $job= Get-Job | Where { $_.Name -eq 'GetSize' -and $_.State -eq "Completed"}
    if($job){$global:jobCnt+=$job.Count; $tmp= $($job | Receive-Job) 2>$null; $job | Remove-Job -force}
    if($drs.Count){$prcnt="{0:P0}" -f ($global:jobCnt/$($drs.Count))}
    if($global:usrSz){OutputSize $global:usrSz; $global:usrSz=@()}
    if($tmp){OutputSize $tmp; $tmp=@()}
    if(-not((Get-Job | Where {$_.Name -eq 'GetSize'}) -or $cnt)){$global:step=6}
  }

  if($cnt){
    $run=(Get-Job | Where {$_.Name -eq 'GetSize' -and $_.State -eq "Running"}).count
    if($run -le $jobCount -and $cnt){$global:cnt--; Start-Job -InputObject "\\$compAdr\C$\Users;$($drs[$cnt])" -ScriptBlock $ScriptBlock -Name GetSize}
  }

  if($step -eq 4.5){    
    if(Get-Job | Where {$_.Name -eq 'System files'}){
      $job= Get-Job | Where { $_.Name -eq 'System files' -and $_.State -eq 'Completed'}
      if($job){$out= $($job | Receive-Job); $job | Remove-Job -force; Output $out; ShCRButton $out}
      if(-not(Get-Job | Where { $_.Name -eq 'System files'})){
        $global:step=5
        if($os -and $CheckBox2.Checked){Output "`nScan users folders size more than 100MB... `n`n"}
      }
    }else{$global:step=5}
  }

  if($step -eq 4.2){    
    if(Get-Job | Where {$_.Name -eq 'Users files Invoke'}){
      $job= Get-Job | Where { $_.Name -eq 'Users files Invoke' -and $_.State -eq "Completed"}
      if($job){
        $out= $job | Receive-Job; $job | Remove-Job -force
        if($out){$global:usrSz= $out; $global:step=4.5}
        else{
          $global:drs= $(Get-ChildItem -Path "\\$compAdr\C$\Users" -Directory -name) 2>$null
          $global:cnt= $drs.Count
          $global:step=4.5
        }
      }
    }else{$global:step=6}
  }

  if($step -eq 4.1){    
    if(Get-Job | Where {$_.Name -eq 'System files Invoke'}){
      $job= Get-Job | Where { $_.Name -eq 'System files Invoke' -and $_.State -eq "Completed"}
      if($job){$out= $job | Receive-Job; $job | Remove-Job -force}
      if($out){Output $out; $global:step=4.2; ShCRButton $out; Output "`n`nScan users folders size more than 100MB... `n`n"}
      else{GetSizeSys; $global:step=4.2}
    }else{$global:step=6}
  }

  if($step -eq 3){
    if($CheckBox4.Checked){
      $job= Get-Job | Where { $_.name -eq "Get date" -and $_.State -eq "Completed"}
      if($job.count -eq (Get-Job | Where { $_.name -eq "Get date"}).count){
        $shrt=0
        $global:usrs= $($job | Receive-Job | Sort-Object {$_.LastLogin -as [datetime]} -Descending) 2>$null; Get-Job | Where { $_.name -eq "Get date"} | Remove-Job -force 
        if($usrs){
          if($CheckBox4.CheckState -ne "Indeterminate"){$shrt=$usrs.Count-10; $usrs=$usrs[0..9]}
          Output $("`r`nUsers last logget-in")
          Output "`n-------------------------------------------------------------------------------"
          $usrs | ForEach {
            if($_.fullname.Length -gt 40){$_.fullname=($_.fullname).Substring(0,39)+">"}
            Output $("`n"+(addSp $($(addSp $_.Login  22) + $_.fullname) 62)+$_.LastLogin)
          }
          if($shrt -gt 0){Output "`n                      +$shrt users.." "SlateBlue" }
          Output "`n-------------------------------------------------------------------------------`n"
        }else{Output "`r`n Can't read C:\Users folder`n" "SaddleBrown"}
        $global:step=4.1
      }
    }else{$global:step=4.1} 
  }

  if($step -eq 2.5){
    if(-not $CheckBox4.Checked){$global:step=3}
    $job= Get-Job | Where { $_.Name -eq "ShowUsr" -and $_.State -eq "Completed"}
    if($job){
      Output "`r`nlogged-in user - "
      $out= $($job | Receive-Job) 2>$null; $job | Remove-Job -Force
      $out | ForEach {$a=$( $_ -split ';'); Output $a[0] $a[1]}
      $global:step=3
    }
  }


  if($step -eq 2.2){
    $job= Get-Job | Where { $_.Name -eq "Get Printers" -and $_.State -eq "Completed"}
    if($job.count -eq (Get-Job | Where { $_.Name -eq "Get Printers"}).count){
      $printTabl= $($job | Receive-Job)
      $job | Remove-Job -Force
      $printrs= PrintTabl $printTabl
      if($printrs){PrintersOut $($printrs)}
      $global:step=2.5
    }
  }

  if($step -eq 2){
    $job= Get-Job | Where { $_.Name -eq 'Get Printers Invoke' -and $_.State -eq 'Completed'}
    if($job){
      $global:printTabl= $($job | Receive-Job)
      $job | Remove-Job -Force
      if($printTabl){$printrs= PrintTabl $printTabl}else{GetPrinters;$global:step=2.2; return}
      if($printrs){PrintersOut $printrs}
    }
    if( -not(Get-Job | Where { $_.Name -eq 'Get Printers Invoke'})){$global:step=2.5}
  }

  if($step -eq 1){
    if($os -and $CheckBox1.Checked){
      $global:step=2
    }else{$global:step=2.5}
  }

  PrProgres

})




function ChkCmp ($comp){
  if ( [system.Net.Sockets.TcpClient]::new().BeginConnect($comp,455,$null,$null).AsyncWaitHandle.WaitOne(150,$false)){return $true}
  $cnt=0; do {$cnt++; if($(ping -4 $comp -n 1 -w 200 | Select-String -Pattern 'TTL=')){return $True}} while ($cnt -lt 4)
}
function GetIP ($cmp){return $([System.Net.Dns]::GetHostAddresses("$cmp") | where {$_.AddressFamily -eq "InterNetwork"} | ForEach {[string]$_}) 2>$nul}
$ping={ping.exe $compAdr -n 1 -4 -w 300 | Where {$_} | Select-Object -First 1 -Skip 1}
function ChkPrt ($cmp, $prt){[system.Net.Sockets.TcpClient]::new().BeginConnect($cmp,$prt,$null,$null).AsyncWaitHandle.WaitOne(500,$false)}
function ChkOS ($comp){ ((Get-WmiObject -Class Win32_OperatingSystem -ComputerName $comp).name -split"\|")[0]}
function addSp([string]$word, $len){ for ($i = $word.Length; $i -lt $len; $i++){$word+=" "}; return [string]$word }
function Exclude ($Id) {$a=$True; $excl | ForEach {if ("$Id" -like "*$_*"){$a=$False}}; return $a}
function Replace ($word) {$repl | ForEach {$word=$word -replace "$_",""}; return $word}
function Hide ($word) {$a=$True; $hide | ForEach { if ("$word" -like "*$_*"){$a=$False}}; return $a}
function Fltr ($id){
  if($id -like "*\*" -and $id -notlike "?:\*"){
    $id= ($id -split '\\') -replace "___",""
    $id[1]=($id[1] -split "&")[0..1] -join "&"
    return $id[0]+"\"+$id[1]
    } else{ return $id}
}
function Shrt($adUsr){ $adUsrSh=$adUsr; $a=($adUsr.trim() -split ' '); if($a[1] -and $a[2]){$adUsrSh=$a[0]+" "+($a[1])[0]+"."+($a[2])[0]+"."}; return $adUsrSh}


  function SMBchk ($cmp) {
    [system.Net.Sockets.TcpClient]::new().BeginConnect($cmp,445,$null,$null).AsyncWaitHandle.WaitOne(500,$false)
  }



function Chek {
  Get-Job | Remove-Job -Force
  if($global:CRButton){$global:CRButton.Dispose()}; $global:CRButton=$null

  $global:sum=@()
  $global:os=$null
  $global:compAdr= ($ListBox2.Text).trim()
  if(-not $global:compAdr){$global:compAdr=$env:COMPUTERNAME;$ListBox2.Text=$global:compAdr}

  $ListBox3.text="`r`n"
  $ip= GetIP $compAdr
  $cnt=2; while($cnt) {$cnt--; $a= $(.$ping); Output "$a`r`n" $(if($a -like '*TTL=*'){"green"; $png=$true}else{"red"})}
  if(-not $ip){Output "`r`nComputer $compAdr unaccessible.."; return}
  if(-not $png){3389,5900,445,22,80,8080,1112,2575,10050 | ForEach {  if(ChkPrt $compAdr 3389){$png=$true;Output "Port $_ open`r`n" "green"}}}

  if(-not $png){Output "`r`nComputer $compAdr unaccessible.."; return}
  if($CheckBox4.Checked){GetUsers}
  PrProgres
  GetSizeInvoke
  PrProgres
  $global:os=ChkOS $compAdr
  PrProgres
  if($os -and $CheckBox1.Checked){GetPrintersInvoke}
  PrProgres
  if($os -and $CheckBox4.Checked){ShowUsr}
  $ListBox3.text="`r`n"
  Output "`n$($os)`n`n"
  PrProgres

  $dt= ((Get-WmiObject Win32_OperatingSystem -ComputerName $compAdr -ErrorVariable err).LastBootUpTime)
  if($dt){
    $curTm= Get-Date -date  $(Get-WmiObject -Class Win32_CurrentTime -ComputerName $compAdr | Where {$_.__CLASS -eq 'Win32_LocalTime'} | select Day,Hour,Minute,Month,Year | ForEach {"$($_.Year,$_.Month,$_.Day  -join '.') $($_.Hour,$_.Minute -join ':')"})
    $dt= Get-Date -date  $($dt.Substring(0,4)+'.'+$dt.Substring(4,2)+'.'+$dt.Substring(6,2)+' '+$dt.Substring(8,2)+':'+$dt.Substring(10,2))
    $tmUp=$curTm - $dt
    if($dt){Output " System boot up date        Uptime`r`n --------------------ÿÿÿÿÿ-----------`r`n $($dt.ToString('dd MMM yyyy HH:mm'))         $($tmUp.Days)d $($tmUp.Hours):$("{0:d2}" -f $tmUp.Minutes)`r`n`r`n"}
  }


if($os){
  $a= Get-WmiObject win32_logicaldisk -Computername  $compAdr | select DeviceID, FreeSpace, Size, VolumeName | Where {$_.Size} 
  if($a){
    Output "`nÿDriveÿÿÿÿSizeÿÿFreeSpaceÿÿÿÿUsageÿÿÿVolumeÿName`n------------------------------------------------`n"
    $a | ForEach {
      $free=$($(100-[int](100*($_.FreeSpaceÿ/ÿ$_.Size)))); $clr="DarkGreen"; if($free -ge 80){$clr="Chocolate"}; if($free -ge 90){$clr="Crimson"}
      Output "ÿÿ$($_.DeviceID.PadRight(3,"ÿ"))ÿ$("$([int]($_.Sizeÿ/ÿ1GB))GB".PadLeft(8,"ÿ")) $("$([int]($_.FreeSpaceÿ/ÿ1GB))GB".PadLeft(10,"ÿ"))"; Output "ÿ$("$free%".PadLeft(8,"ÿ"))" $clr; Outputÿ"ÿÿÿ$($_.VolumeName)`n"
    }
    Output "------------------------------------------------`n"
  }
}

  PrProgres

if($os -and $CheckBox3.Checked){CompHw}

  PrProgres

  if(ChkPrt $compAdr 3389){Output "`nRDP port  3389" "DarkBlue"}
  if(ChkPrt $compAdr 5900){Output "`nVNC port  5900" "SeaGreen"}
  if(ChkPrt $compAdr 445 ){Output "`nSMB port   445" "MediumVioletRed"}
  if(ChkPrt $compAdr 22  ){Output "`nSSH port    22" "Violet"}
  if(ChkPrt $compAdr 80  ){Output "`nHTTP port   80" "DarkOrange"}
  if(ChkPrt $compAdr 8080){Output "`nHTTP port 8080" "Blue"}
  if(ChkPrt $compAdr 1541){Output "`n 1C port  1541" "SeaGreen"}
  if(ChkPrt $compAdr 10050){Output "`nZabbix port 10050" "DarkViolet"}

  PrProgres

  Output "`n"

  $global:step=1
  $timer.Start()
}

function CompHw {

  if (-not $ip){return}
  $hst= ((get-wmiobject -list "StdRegProv" -computername $compAdr -namespace root\default).GetStringValue(2147483650,"SOFTWARE\Microsoft\Virtual Machine\Guest\Parameters","HostName")).sValue
  $memSum=0; $model=""; $memModule=""
  $cpu=(Get-WmiObject -Class CIM_Processor -ComputerName $compAdr).Name | Select-Object -first 1
  $cpu=$cpu -replace '  ',' '; $cpu=$cpu -replace '  ',' '; $cpu=$cpu -replace '  ',' ' ;$cpu=$cpu -replace '  ',' '
  $cpu=$cpu -replace ' CPU ',' '
  $cpu=$cpu -replace 'Intel\(R\) Core\(TM\)2 Duo','Core2Duo'
  $cpu=$cpu -replace 'Intel\(R\) Pentium\(R\)','Pentium'
  $cpu=$cpu -replace 'Intel\(R\) Core\(TM\)','Core'
  $cpu=$cpu -replace 'Intel\(R\) Celeron\(R\)','Celeron'
  $cpu=$cpu -replace 'Intel\(R\) Xeon\(R\)','Xeon'
  $cpu=$cpu -replace 'Pentium\(R\) Dual-Core','Pentium'
  $cpu=$cpu -replace ' with Radeon\(tm\) HD Graphics',''
  $cpu=$cpu -replace ' with Radeon Vega Mobile Gfx',''
  $cpu=$cpu -replace ' with Radeon Vega Graphics','' 
  $cpu=$cpu -replace ' @ ',' '
  if ("$cpu" -ne ""){
    $memModule=((Get-WmiObject -Class CIM_PhysicalMemory -ComputerName $compAdr).Capacity | ForEach {$mem=[int]($_*10/1073741824)/10; if($mem -ge 0.5){$mem; $memSum=$memSum+$mem}}) -join ' '
    $sys=Get-WmiObject -Class CIM_ComputerSystem -ComputerName $compAdr
    $model=$sys.model
    $model=("$model").Trim()
    $model=$model -replace 'System Product Name','Noname'
    $model=$model -replace 'To be filled by O.E.M.','Noname'
    $model=$model -replace 'VMware Virtual Platform','VMware' 
    $cores=$sys.NumberOfLogicalProcessors
    $compName=$sys.Name+$(if($sys.Domain){"."+$sys.Domain})
    $hdd=""; Get-WmiObject -Class CIM_diskdrive -ComputerName $compAdr | ForEach {if($_.Size -gt 0){$hdd+=[string]($_.Model+" "+[int]($_.Size / (1000000000)))+"GB`r`n"}}
    Output "`r`n    $(addSp $compName 23) $ip`r`n-------------------------------------------------------`r`n$(addSp "CPU" 15) :  $cpu($cores core) `r`n$(addSp "Memory" 15) :  $memModule (Sum:$memSum`GB) `r`n$(addSp "Manufacturer" 15) :  $model"
    $(if($hst){Output " ("; Output "$hst" "MediumVioletRed"; Output ")"}); Output "`r`n"
    if($hdd){Output "`r`nHard Disk Drive:"; Output "`r`n$hdd" "DarkBlue"} 
  }
}



function ShowUsr {

  Start-Job -InputObject "$(1*$CheckBox5.Checked);$compAdr" -Name 'ShowUsr' -ScriptBlock {
    $srch = New-Object -TypeName System.DirectoryServices.DirectorySearcher
    function Shrt($adUsr){ $adUsrSh=$adUsr; $a=($adUsr.trim() -split ' '); if($a[1] -and $a[2]){$adUsrSh=$a[0]+" "+($a[1])[0]+"."+($a[2])[0]+"."}; return $adUsrSh}
    $in="$input" -split ';'
    $compAdr=$in[1]
    $cnt=0; $job=@(); $act=@(); $usrs=@(); $col="SaddleBrown"
    $getNm={$nm=Shrt ([string]$($srch.Filter = "(&(objectCategory=person)(sAMAccountName=$a))"; $srch.FindOne().Properties.displayname) 2>$null); if($nm){$nm}else{"($a)"}}
    $query=$null
    $job = Start-Job -ScriptBlock {Start-Sleep -Milliseconds 4500; Get-Process -ProcessName "quser" | Stop-Process}
    $query= quser.exe  /SERVER:$compAdr 2>&1
    $job | Remove-Job -force
    if(-not $query.TargetObject -and $query){
      $usrs=@(); $act=@(); $list=@(); $col="SaddleBrown"
      $query | Select-Object -Skip 1 | ForEach {
        $a=($_ -split " " | Where {$_})[0]
        if($_ -like "*console*"){$act+="$(.$getNm) (console)"}
        elseif($_ -like "*rdp-tcp*"){$act+="$(.$getNm) (rdp)"}
        else{$usrs+=$(.$getNm)}
      }
    }

    if($query.TargetObject -or -not $query ){
      $job = Start-Job -ScriptBlock {Start-Sleep -Milliseconds 7500; Get-Process -ProcessName "tasklist" | Stop-Process}
      $query= $(tasklist.exe /S $compAdr /FO "CSV" /V /FI "IMAGENAME eq explorer.exe" 2>&1)
      $job | Remove-Job -force
      if($query.TargetObject){$err= $query.TargetObject; $col="red"}
      elseif(-not $query){$usrs="computer not responding"; $col="black"}
      else{$list=@(); $usrs= $query -replace '"' | Select-Object -Skip 1 | ForEach {
          $qr= $_ -split ','; $indx=5; if($qr.count -gt 7){$indx=6}
          $a= ($qr[$indx] -split '\\')[1]
          if($a -notin $list){
            $list+=$a
            if($qr[2]){$act+="$(.$getNm) ($(($qr[2] -split '-')[0]))"}else{$(.$getNm)}
          }
        }
        if(-not $act -and -not $usrs){$usrs="nobody"; $col="black"}
      }
    }

    $cnt=0; $lmt=7; $all=$act.Count+$usrs.Count; if($all -gt 7){$lmt=5}; 
    if("1" -eq $in[0]){$lmt= $all}
    $act  | ForEach {if($cnt -lt $lmt){"$_`r`n                 ;SlateBlue"; $cnt++}}
    $usrs | ForEach {if($cnt -lt $lmt){"$_`r`n                 ;$col"; $cnt++}}
    if($all-$cnt){"..+$($all -$cnt)`r`n;SlateBlue"}
    "`r`n ;black"
  }
}

function GetSizeInvoke {
  if(-not $CheckBox2.Checked){return}
  Start-Job -InputObject "$compAdr" -Name 'System files Invoke' -ScriptBlock {
    $compAdr="$input"
    $ScriptBlock= {
      $cnt=0
      'C:\Windows\Temp','C:\$Recycle.Bin','C:\ProgramData\Microsoft\Diagnosis','C:\ProgramData\Microsoft\Search\Data\Applications\Windows\Windows.edb' | ForEach {
        $cnt++;$tmp=""; $a=[int](($(Get-ChildItem -Path "$_" -Recurse -force | ForEach {$_.Length}) 2>$null | Measure-Object  -Sum).Sum / 1MB)
        if($a -ge 500 -or $cnt -le 2){$("`n$_ ").PadRight(22," ")+$("{0:n0}" -f $a+" MB`n").PadLeft(9," ")}
      }
      
      $fls='C:\hiberfil.sys','C:\pagefile.sys'
      $fls | ForEach {
        $fl=$_; $tmp=""; $a=$(Get-ChildItem -Path 'C:\' -Force | Where {$_.FullName -eq $fl}).Length / 1MB 2>$null 
        if($a -ge 1){$("`n$fl").PadRight(22," ")+$("{0:n0}" -f $a+" MB`n").PadLeft(9," ")}
      }
    }
    $(Invoke-Command -ComputerName $compAdr -ScriptBlock $ScriptBlock -ErrorVariable err) 2>$null
  }

  Start-Job -InputObject "$compAdr" -Name 'Users files Invoke' -ScriptBlock {
    $compAdr="$input"
    Invoke-Command -ComputerName $compAdr -ScriptBlock {
      $fldrs='Downloads;Pictures;Desktop;Videos;Documents;Music;AppData\Local\Temp;AppData\Local\Microsoft\Windows\INetCache;AppData\Local\Google\Chrome\User Data\Default\Cache;AppData\Local\Microsoft\Outlook' -split ';'       
      $(Get-ChildItem -Path "C:\Users" | Where {$_.Mode -like "d????*"}) 2>$null | ForEach {
        $usrDir=$_.Name
        $tmp= $fldrs | ForEach {
          $a=[int](($(Get-ChildItem -Path "C:\Users\$usrDir\$_" -Recurse -Force | ForEach {$_.Length}) 2>$null  | Measure-Object  -Sum).Sum / 1MB)
          if($a -gt 100){@{usrDir=$usrDir;Dir=$_;Size=$a}}
        }
        if($tmp){$tmp}
      }
    }
  }
}

function GetSizeSys {
    $global:jobCnt=0
    if($CheckBox2.Checked -and (Test-Path "\\$compAdr\c$\Users")){
      $Script={
        $tmp=""; $a=$(($(Get-ChildItem -Path "$input" -Recurse -force | ForEach {$_.Length}) 2>$null | Measure-Object  -Sum).Sum / 1MB)
        $("`nC:\Windows\Temp ").PadRight(22," ")+$("{0:n0}" -f $a+" MB`n").PadLeft(9," ")
      }
      Start-Job -InputObject "\\$compAdr\c$\Windows\Temp" -ScriptBlock $Script -Name 'System files'

      $Script={
        $tmp=""; $a=$(($(Get-ChildItem -Path "$input" -Recurse -force | ForEach {$_.Length}) 2>$null | Measure-Object  -Sum).Sum / 1MB)
        if($a -ge 1GB){$("`nC:\ProgramData\Microsoft\Diagnosis").PadRight(36," ")+$("{0:n0}" -f $a+" MB`n").PadLeft(9," ")}
      }
      Start-Job -InputObject "\\$compAdr\c$\ProgramData\Microsoft\Diagnosis" -ScriptBlock $Script -Name 'System files'

      $Script={
        $tmp=""; $a=$(Get-ChildItem -Path "$input" -Force | Where {$_.Name -eq 'hiberfil.sys'}).Length / 1MB 2>$null 
        if($a -ge 1){$("`nC:\hiberfil.sys").PadRight(22," ")+$("{0:n0}" -f $a+" MB`n").PadLeft(9," ")}
      }
      Start-Job -InputObject "\\$compAdr\c$" -ScriptBlock $Script -Name 'System files'

      $Script={
        $tmp=""; $a=$(Get-ChildItem -Path "$input" -Force | Where {$_.Name -eq 'pagefile.sys'}).Length / 1MB 2>$null 
        $("`nC:\pagefile.sys").PadRight(22," ")+$("{0:n0}" -f $a+" MB`n").PadLeft(9," ")
      }
      Start-Job -InputObject "\\$compAdr\c$" -ScriptBlock $Script -Name 'System files'

      $Script={
        $tmp=""; $a=$(($(Get-ChildItem -Path "$input" -Recurse -force | ForEach {$_.Length}) 2>$null | Measure-Object  -Sum).Sum / 1MB)
        $("`nC:\`$Recycle.Bin ").PadRight(22," ")+$("{0:n0}" -f $a+" MB`n").PadLeft(9," ")
      }
      Start-Job -InputObject "\\$compAdr\c$\`$Recycle.Bin" -ScriptBlock $Script -Name 'System files'

      $Script={
        $tmp=""; $a=($(Get-ChildItem -Path "$input" -force) 2>$null ).Length / 1MB
        if($a -ge 1024){"`nC:\ProgramData\Microsoft\Search\Data\Applications\Windows\Windows.edb "+$("{0:n0}" -f $a+" MB`n").PadLeft(9," ")}
      }
      Start-Job -InputObject "\\$compAdr\c$\ProgramData\Microsoft\Search\Data\Applications\Windows\Windows.edb" -ScriptBlock $Script -Name 'System files'
   }
}

function GetUsers {

  $usrDir=@()

  if(SMBchk $compAdr){if (Test-Path "\\$compAdr\C$\Users"){ Get-ChildItem -Path "\\$compAdr\C$\Users\" -Directory -name -Exclude $exld | ForEach {$usrDir+=@($_)}}}

  $Script={
    $usrDir= @($input)
    $srch = New-Object -TypeName System.DirectoryServices.DirectorySearcher
    $usrDir | ForEach {
      $usr=($_ -split '\\')[-1]; $acces=""; $adu=@()
      $acces=$(((Get-ChildItem -Path "$_\AppData\Local\Temp" -force).LastAccessTime | Sort-Object {$_ -as [datetime]} -Descending)[0].ToString('dd.MM.yyyy')) 2>$null
      if(-not $acces){$acces=$(((Get-ChildItem -Path "$_\NTUSER.DAT" -force).LastWriteTime).ToString('dd.MM.yyyy')) 2>$null}
      $adu= $($srch.Filter = "(&(objectCategory=person)(sAMAccountName=$usr))"; $srch.FindOne()).Properties 2>$null
      if (-not $adu -and $acces){
        $sid=$(((Get-Acl "$_\NTUSER.DAT").Sddl -split 'G:')[0] -replace 'O:','') 2>$null 
        if($sid){$adu= $($srch.Filter = "(&(objectCategory=person)(objectSID=$sid))"; $srch.FindOne()).Properties 2>$null}
      }
      @{Login=$usr;fullname=$adu.displayname;LastLogin=$acces;DistinguishedName=[string]$adu.distinguishedname;Enabled=-not [bool]($($adu["useraccountcontrol"]) -band 0x000002)}
    }
  }

  $jobCnt=[int]($usrDir.Count/40)+1
  $usrsCount=[int](($usrDir.Count/$jobCnt)+.5)
  $a=@()  

  if($usrDir -ne $nul){
    $usrDir | ForEach {
      $a+="\\$compAdr\C$\Users\$_"
      if($a.Count -ge $usrsCount){$a | Start-Job -ScriptBlock $Script -Name "Get date"; $a=@()}
    }
    if($a){$a | Start-Job -ScriptBlock $Script -Name "Get date"}
  }
}
  
function GetPrintersInvoke {  

  Start-Job -InputObject "$compAdr" -Name 'Get Printers Invoke' -ScriptBlock { 
    Invoke-Command -ComputerName "$input" -ScriptBlock {

    function GetIP ($cmp){return $([System.Net.Dns]::GetHostAddresses("$cmp") | where {$_.AddressFamily -eq "InterNetwork"} | ForEach {[string]$_}) 2>$nul}

    $pth= 'SYSTEM\ControlSet001\Control\Print\Monitors'
    $key1="Ports","$pth\Standard TCP/IP Port\Ports","HostName"
    $key2="Ports","$pth\HP Standard TCP/IP Port\Ports","IPAddress"
    $key3="Ports","$pth\Advanced TCP/IP Port Monitor\Ports","IPAddress"
    $key4="Ports","$pth\WSD Port\Ports","Printer UUID"
    $pth= 'SOFTWARE\Microsoft\Windows NT\CurrentVersion\Print'
    $key5="Printers","$pth\Printers","port"
    $key6="UsrPrinters","$pth\Providers\Client Side Rendering Print Provider","Printers\Connections"
    $gtReg= {(Get-ChildItem -Path "HKLM:\$pth\$pth2").PSChildName 2>$null}
    $key1,$key2,$key3,$key4,$key5 | foreach {
      $class=$_[0]; $pth=$_[1]; $vl=$_[2];$pth2=''
      .$gtReg | Where {$_} | Where {$_ -notlike "*:OneNote*"} | ForEach {
        $prt=$_;$pth2="$prt\"; $prtIP= get-ItemPropertyValue "HKLM:\$pth\$pth2" $vl
        $ip=''; if($prtIP -notmatch '^\d{1,3}\.\d{1,3}\.\d{1,3}\.\d{1,3}$'){ $ip= GetIP $prtIP}
        @{Class=$class; Name=$prt;Value=$prtIP; IP=$ip}
      }
    }

    $pth=$key6[1];$class=$key6[0];$vl=$key6[2];$pth2='' 
    .$gtReg | Where {$_ -notlike "*:OneNote*"} | Where {$_}| ForEach {$name=$_;$pth2="$_\$vl"; $prnr= .$gtReg; if($prnr -or $name){@{Class=$class; Name=$name; Value=$prnr}}}
    }
  }
}


function GetPrinters {  
  $compName=$compAdr
  $reg1='get-wmiobject -list "StdRegProv" -computername $compAdr -namespace root\default'
  $reg2='(($reg1).EnumKey(2147483650,"$key")).snames'
  $key1="Ports","SYSTEM\ControlSet001\Control\Print\Monitors\Standard TCP/IP Port\Ports\",'(($reg1).GetStringValue(2147483650,"$($key+"\"+$_)","HostName")).sValue',$compAdr
  $key2="Ports","SYSTEM\ControlSet001\Control\Print\Monitors\HP Standard TCP/IP Port\Ports\",'(($reg1).GetStringValue(2147483650,"$($key+"\"+$_)","IPAddress")).sValue',$compAdr
  $key3="Ports","SYSTEM\ControlSet001\Control\Print\Monitors\Advanced TCP/IP Port Monitor\Ports\",'(($reg1).GetStringValue(2147483650,"$($key+"\"+$_)","IPAddress")).sValue',$compAdr
  $key4="Ports","SYSTEM\ControlSet001\Control\Print\Monitors\WSD Port\Ports",'(($reg1).GetStringValue(2147483650,"$($key+"\"+$_)","Printer UUID")).sValue',$compAdr
  $key5="Printers","SOFTWARE\Microsoft\Windows NT\CurrentVersion\Print\Printers\",'(($reg1).GetStringValue(2147483650,"$($key+"\"+$_)","Port")).sValue',$compAdr
  $key6="UsrPrinters","SOFTWARE\Microsoft\Windows NT\CurrentVersion\Print\Providers\Client Side Rendering Print Provider\",'(($reg1).EnumKey(2147483650,("$($key+$_+"\Printers\Connections")"))).sNames',$compAdr
  $out=@();
  $Script={
    $inpt=@($input); $class=$inpt[0]; $key=$inpt[1]; $code=$inpt[2];$compAdr=$inpt[3]
    function Get-Reg ($code){Invoke-Expression $ExecutionContext.InvokeCommand.ExpandString($code)}
    $reg1='get-wmiobject -list "StdRegProv" -computername $compAdr -namespace root\default'
    $reg2='(($reg1).EnumKey(2147483650,"$key")).snames'
    Get-Reg $reg2 | foreach {$vl=Get-Reg $code; @{CompName=$compAdr;Name=$_; Value=$vl; Class=$class}} #) 2>$nul
  }
  $key1,$key2,$key3,$key4,$key5,$key6 | foreach { $_ | Start-Job -ScriptBlock $Script -Name "Get Printers"}
}

function  PrintTabl ($printTabl){
  $srch = New-Object -TypeName System.DirectoryServices.DirectorySearcher
    $printersOut=@()
    $printTabl | Where {$_.Class -eq "Printers"} | foreach { 
      $prName=$null; $Port=$null; $prPort=$null; $dhcpName=$null; $prName=$_.Name; $prPort=$_.Value
      if($prPort){
        if ($prPort -like "USB*" -or $prPort -like "DOT4_0*" -or $prPort -like "CNM*"){ "$prName;USB"}
        else{
          $Port= $printTabl | Where {$_.Class -eq "Ports" -and $_.Name -eq $prPort};
          if($Port){
            $prPort= $Port.Value; 
            if($Port.Name -like "WSD-*"){
                $a=''; $reg= 'LocalMachine',"SYSTEM\CurrentControlSet\Enum\SWD\DAFWSDProvider\uuid:$($Port.Name)",'LocationInformation'
                $reg2= 'LocalMachine',"SYSTEM\CurrentControlSet\Enum\SWD\DAFWSDProvider\urn:uuid:$($Port.Name)",'LocationInformation'
                $reg, $reg2 | ForEach {if(-not $a){$a= $([Microsoft.Win32.RegistryKey]::OpenRemoteBaseKey($_[0], $compName).OpenSubKey($_[1], $true).GetValue($_[2])) 2>$null}}
                if($a){$prPort="WSD:"+($a -replace 'http://' -split ':')[0]; if($a -like "*``[*::*]*"){$prPort="WSD:"+($a -replace 'http://\[' -split ']:')[0]}}else{$prPort=$Port.Name}
            }
            "$prName;$prPort;$($Port.IP)"
          }
        }
      }
    }
  
  $printersSort=@()

  $printTabl | Where {$_.Class -eq "UsrPrinters" -and $_.Value -ne $nul} | foreach {
    $sid=$_.Name; $adUser=$nul; $adUser= [string]$($srch.Filter = "(&(objectCategory=person)(objectSID=$sid))"; $srch.FindOne()).Properties.displayname 2>$null
    if($adUser){$_.Value | ForEach {$printersSort+="$(Shrt $adUser);$($_ -replace ',','\')"}}
  }
  $printersSort | Sort-Object | Get-Unique | foreach {$a=($_ -split ';'); $a[1]+";"+$a[0]}
}



function PrintersOut ($printersOut){
      $compName=(@($printersOut)[0] -split ';')[0]
      $compCol="Red"; if (($printersOut[0] -split ';')[1]){$compCol="Green"};
      Output "`r`n`r`nPrinters"
      Output "`n------------------------------------------------------------------------`n" 
      $printersOut | foreach {
        $a=($_ -split ';')                               
        $clr="Blue"; if([string]$a[0] -like '\\*'){$clr="black"}
        $a[0]=$a[0] -replace 'LaserJet','LJ'
        $a[0]=$a[0] -replace 'Professional','Pro' 
        if($a[0].Length -gt 43){$a[0]=$a[0].Substring(0,42)+">"}
        Output $(addSp $a[0] 45) $clr
        Output $(addSp $a[1] 17); Output $($a[2]) $clr; Output "`n"
      }
    Output "------------------------------------------------------------------------`n" 
}

function PrintOut ($dev){
  $list=@(); $ListBox3.text="`r`n"
  $cl=@(); $($dev.Class | ForEach { if($_ -notin $cl){$cl+=$_; $_}}) | ForEach { 
    $Class=$_
    $dev | Where {$_.Class -eq $Class} | Where {Hide $_.FriendlyName} | Sort-Object {$_.FriendlyName} | ForEach-Object{
      $Id=""; if( (exclude $_.InstanceId) ){$Id= $_.InstanceId}
      $list+=@{Class=$Class;dev=("   "+(addSp (Replace $_.FriendlyName) 24)+"    ")+";"+(Fltr $ID)+"`r`n"}
    }
  }

  If (-not($global:list)){$global:list=$list}

  $list.Class | Get-Unique | ForEach { 
    $Class=$_; $out=@();$outOff=@(); $dev=@()
    txt "$Class`r`n$sprtr`r`n"
    $dev= ($list | Where {$_.Class -eq $Class}).dev
    ($global:list | Where {$_.Class -eq $Class}).dev | ForEach {if ($_ -in $dev){$Out+=$_}else{$outOff+=$_}}
    $out | ForEach {$a=$_ -split ";"; Output $a[0] "blue"; Output $a[1] "green"} 
    $dev | Where { $_ -notin $Out} | ForEach {$a=$_ -split ";"; Output $a[0] "blue" $fontBold ; Output $a[1] "green" $font}
    $outOff | ForEach {$a=$_ -split ";"; Output $a[0] "red"; Output $a[1] "green"}
  txt "`r`n"
  }

txt "`r`n`r`n"; 
$ListBox3.SelectionStart=0
$global:list=$list
}


function ClrRecycle {

  $code='$comp="'+"$compAdr"+'"
    $ScriptBlock= {
      $dir= "$Input"
      $mnth=2
      cls; $expr= (Get-Date).AddMonths(-$mnth);
      Write-Host "`n`n Script run on $env:COMPUTERNAME"
      Write-Host "`n`n Clean directory by delete files older than $mnth month `n`n Folder - $dir  Scan.."  -n
      $($a= Get-ChildItem -Path $dir -Recurse -Force | Where {$_.Mode -like "-????*"} | select Length,LastWriteTime,FullName) 2>$null 
      Write-Host ".." -n
      $all= [int]$((($a | Where {$_.Length} | ForEach {$_.Length}) | Measure-Object  -Sum).Sum / 1MB)
      $del= $a | Where {$_.LastWriteTime -le $expr}
      $old= [int]$((( $del | ForEach {$_.Length}) | Measure-Object  -Sum).Sum / 1MB)
      $prcnt="{0:P0}" -f ($old/(1+$all))
    
      Write-Host ". Done `n`n Folder $dir, size - $all MB"
      Write-Host "`r Files to delete - $((" {0:n0}" -f $old+" MB").PadLeft(9," ")), $($del.Count) files ($prcnt)"
      Write-Host "`n`n`n Delete files..." -n
    
      $time= Measure-Command {$($del | ForEach { Remove-Item $_.FullName -Force}) 2>$null} 
    
      Write-Host ".. Done in $([int]$time.TotalSeconds) seconds  "
    
      Write-Host "`n`n Delete empty folders..." -n
      $GetDir={$(Get-ChildItem -Path $dir -Force | Where {$_.Mode -like "d????*"}).FullName 2>$null}
      function ChkFldr ($dir){if($(Get-ChildItem -Path $dir -File -Recurse -Force) 2>$null){$true; return};$false}
      function DelFldr ($dir){if(ChkFldr $dir){.$GetDir | Where {$_} | ForEach {DelFldr $_}}else{$(Remove-Item $_ -Force -Recurse) 2>$null}}
      .$GetDir | ForEach {DelFldr $_}
      Write-Host ".. Done" 
    
      $all= [int]$((($(Get-ChildItem -Path $dir -File -Recurse -Force) 2>$null | Where {$_.Length} | ForEach {$_.Length}) | Measure-Object  -Sum).Sum / 1MB)
      Write-Host "`n`n Folder $dir, size - $all MB`n"
      return "OK"
   }

   $dir= "C:\`$Recycle.Bin"
   $out=$(Invoke-Command -ComputerName $comp -ScriptBlock $ScriptBlock -InputObject $dir) 2>$null
   if($out -ne "OK"){
   $err=$null
   $dir= "\\$comp\C$\`$Recycle.Bin"
   $err; timeout 3 >$null
   $(Invoke-Command -ScriptBlock $ScriptBlock -InputObject $dir -ErrorVariable err) 2>$null

  }
  Write-Host "`n   -= Press any key =- `n`n"
  timeout 300 >$null
  ' 

  Start-Process cmd -ArgumentList " /C powershell.exe  -EncodedCommand $([Convert]::ToBase64String([System.Text.Encoding]::Unicode.GetBytes($code)))"

}



function GetDev ($comp){
  $class="Win32_OperatingSystem","Win32_VideoController","Win32_DesktopMonitor","Win32_DiskDrive","Win32_Share","Win32_IDEController","Win32_SoundDevice","Win32_Processor","Win32_USBController","Win32_USBControllerDevice","Win32_USBHub","Win32_Printer","Win32_NetworkAdapter"
  $dev=@();txt "`r`n`r`n`r`n"
  $class | ForEach { 
    $clss=$_ -replace "Win32_",""; Get-WmiObject -Class $_ -ComputerName $comp | Where {$_.PhysicalAdapter -or $_.SpoolEnabled -or $_.Status -eq "OK"} | Select Name, PNPDeviceID, Model, Path, Size |ForEach {
    if($clss -like "USB*"){$clss="USB"}
    if($clss -eq "Share"){$_.PNPDeviceID= $_.Path}
    if($clss -eq "OperatingSystem"){$_.name=($_.name -split"\|")[0]}
    if($clss -eq "DiskDrive"){ $_.Name=$_.Model; $_.PNPDeviceID="";if($_.Size -gt 0){ $_.PNPDeviceID="$([int]($_.Size / (1000000000)))"+"GB"}}

    @{FriendlyName=$_.Name; InstanceId=$_.PNPDeviceID; Class=$clss}

    if($clss -eq "Processor"){ @{FriendlyName=(Get-WmiObject -Class CIM_ComputerSystem -ComputerName $comp).model; InstanceId=""; Class="Manufacturer"}}
    }
  }
  Get-WmiObject  -Class Win32_PnPEntity -ComputerName $comp | Select Name, PNPDeviceID, Manufacturer | ForEach {
    if($_.name -like "*(COM*"){@{FriendlyName=$_.Name; InstanceId=$_.PNPDeviceID; Class="Com Ports"}}
    if($_.Manufacturer -eq $nul -and $_.PNPDeviceID -notlike "ROOT\LEGACY*"){@{FriendlyName=$_.Name; InstanceId=$_.PNPDeviceID; Class="Unknow Devices"}}
  }
}


function Device ($comp){
  Get-Job | Remove-Job -Force
  $compAdr= ($ListBox2.Text).trim()
  if(-not $compAdr){$compAdr=$env:COMPUTERNAME;$ListBox2.Text=$compAdr}
  $ListBox3.text="`r`n"
  $ip=GetIP $compAdr
  if($lstcmp -ne $compAdr){$global:list=@();$global:lstcmp=$compAdr }
  $cnt=2; while($cnt) {$cnt--; $a= $(.$ping); Output "$a`r`n" $(if($a -like '*TTL=*'){"green"; $png=$true}else{"red"})}

  if(-not $png){txt "`r`nComputer $compAdr unaccessible.."; return}
  $os=ChkOS $compAdr
  Output "`n$($os)`n`n"

  $dev=@(); $dev+=@{FriendlyName=$os; Class="OperatingSystem"}
  if($dev){PrintOut (GetDev $compAdr)}
  else{txt "`r`nComputer $compAdr unaccessible.."}
}



$Form = New-Object System.Windows.Forms.Form
$Form.Size = New-Object System.Drawing.Size(687,842)
$Form.Text ="ScanPC"
$Form.MaximizeBox = $true
$Form.MinimizeBox = $true
$Form.SizeGripStyle = [System.Windows.Forms.SizeGripStyle]::Hide
$Form.BackColor = "#c0c0c0"
$Form.ShowIcon = $true
$Form.MaximumSize = @{Width=1300; Height=900}
$Form.MinimumSize = @{Width=400; Height=500}
$Form.WindowState = "Normal"
$Form.Opacity = 1.0
$Form.TopMost = $false

$ListBox2 = New-Object System.Windows.Forms.TextBox
$ListBox2.Location = New-Object System.Drawing.Size(2,2)
$ListBox2.Size = New-Object System.Drawing.Size(190,16)
$ListBox2.Text=$comp
$ListBox2.Font = "Courier New,10"
$ListBox2.WordWrap = $False
$ListBox2.AllowDrop = $True
$ListBox2.Anchor='Top,Left'
$Form.Controls.Add($ListBox2)

$cntxtMenu = New-Object System.Windows.Forms.ContextMenuStrip
$cntxtMenu.Items.Add("Copy").add_Click({ $ListBox3.Copy() })

$ListBox3 = New-Object System.Windows.Forms.RichTextBox
$ListBox3.Location = New-Object System.Drawing.Point(2,26)
$ListBox3.Size = New-Object System.Drawing.Size(665,770)
$ListBox3.Text=$help
$ListBox3.MultiLine = $True
$ListBox3.DetectUrls = $false
$ListBox3.Font = "Courier New,10"
$ListBox3.ReadOnly = $true
$ListBox3.WordWrap = $False
$ListBox3.ScrollBars = "Vertical"
$ListBox3.ReadOnly = $True
$ListBox3.AllowDrop = $True
$ListBox3.Anchor='Top,Left,Bottom,Right'
$ListBox3.ContextMenuStrip = $cntxtMenu
$ListBox3.ShortcutsEnabled = $false
$Form.Controls.Add($ListBox3)

$font=$ListBox3.Font

$fontBold = New-Object Drawing.Font($font.FontFamily, $font.Size, [Drawing.FontStyle]::Bold)

$Label = New-Object System.Windows.Forms.label
$Label.Location = New-Object System.Drawing.Size(595,3)
$Label.Size = New-Object System.Drawing.Size(100,16)
$Label.Text = ""
$Label.Anchor='Top,Right'
$Label.Font = New-Object Drawing.Font("Terminal",12, [Drawing.FontStyle]::Bold) #Regular
$Form.Controls.Add($Label)

$DeviceButton = New-Object System.Windows.Forms.Button
$DeviceButton.Location = New-Object System.Drawing.Point(200,4)
$DeviceButton.Size = New-Object System.Drawing.Size(60,20)
$DeviceButton.Text = 'Devices'
$DeviceButton.Add_Click({Device})
$form.Controls.Add($DeviceButton)

$ScanButton = New-Object System.Windows.Forms.Button
$ScanButton.Location = New-Object System.Drawing.Point(280,4)
$ScanButton.Size = New-Object System.Drawing.Size(60,20)
$ScanButton.Text = 'Scan'
$ScanButton.Add_Click({Chek})
$form.Controls.Add($ScanButton)

$CheckBox1 = New-Object System.Windows.Forms.CheckBox
$CheckBox1.Location = New-Object System.Drawing.Size(345,5)
$CheckBox1.Size = New-Object System.Drawing.Size(50,16)
$CheckBox1.Text = "Print"
$Form.Controls.Add($CheckBox1)

$CheckBox2 = New-Object System.Windows.Forms.CheckBox
$CheckBox2.Location = New-Object System.Drawing.Size(400,5)
$CheckBox2.Size = New-Object System.Drawing.Size(45,16)
$CheckBox2.Text = "Size"
$Form.Controls.Add($CheckBox2)


$CheckBox3 = New-Object System.Windows.Forms.CheckBox
$CheckBox3.Location = New-Object System.Drawing.Size(450,5)
$CheckBox3.Size = New-Object System.Drawing.Size(50,16)
$CheckBox3.Text = "HW"
$Form.Controls.Add($CheckBox3)

$CheckBox4 = New-Object System.Windows.Forms.CheckBox
$CheckBox4.Location = New-Object System.Drawing.Size(500,5)
$CheckBox4.Size = New-Object System.Drawing.Size(50,16)
$CheckBox4.ThreeState = $True
$CheckBox4.Text = "Usr"
$CheckBox4.Checked=0
$Form.Controls.Add($CheckBox4)

function ShCRButton ($in){
  if($global:CRButton){$global:CRButton.Dispose()}; $global:CRButton=$null
  $sz= (($in | Select-String 'Recycle.Bin' -SimpleMatch) -split '' | Where {$_ -match '\d'}) -join ''
  if(500 -le $sz){
    $CRButton = New-Object System.Windows.Forms.Button
    $CRButton.Location = New-Object System.Drawing.Point(550,3)
    $CRButton.Size = New-Object System.Drawing.Size(30,20)
    $CRButton.Text = 'CR'
    $CRButton.Add_Click({ClrRecycle})
    $global:CRButton=$CRButton
    $Form.Controls.Add($global:CRButton)
  }
}

$Form.Add_Shown({$Form.Activate(); $DeviceButton.focus()})

[void] $Form.ShowDialog()

[Console.Window]::ShowWindow(( $consolePtr = [Console.Window]::GetConsoleWindow()), 4)