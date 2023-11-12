[String]$runningvms = & "C:\Program Files\Oracle\VirtualBox\VBoxManage.exe" list runningvms
$allvms = $runningvms -Split " "
Write-Host "Seperating all the fields into Array...". $vms
Write-Host "VM...". $runningvms

## Lock the host
Start-Process -FilePath "C:\Windows\System32\rundll32.exe" -ArgumentList "user32.dll,LockWorkStation"

[String]$handlevm = @()
[Int]$tellen = 0
foreach ($onevm in $allvms)
  {
    Write-Host "Examinating: ". $onevm
    Write-Host "-->"
    If ($onevm.Contains("{"))
      {
        Write-Host "Skip: ". $onevm
        Write-Host "-->"
      }
    else
      {
        Write-Host "Use: ". $onevm
        Write-Host "-->"
        
        ## Get the information about the vm
        [String]$infovm = & "C:\Program Files\Oracle\VirtualBox\VBoxManage.exe" showvminfo $onevm

        If ($infovm.Contains("Windows"))
          {
            Write-Host "Windows VM, Executing Lock Command in Windows"
            ## Send Ctrl+ESC
            Start-Process -FilePath "C:\Program Files\Oracle\VirtualBox\VBoxManage.exe" -ArgumentList "controlvm $onevm keyboardputscancode 1d 01 81 9d"
            Write-Host "Sending CTRL+ESC: Start-Process -FilePath "C:\Program Files\Oracle\VirtualBox\VBoxManage.exe" -ArgumentList \"controlvm $onevm keyboardputscancode 1d 01 81 9d\""
            sleep 3

            ## Send the shutdown command and lock command
            Start-Process -FilePath "C:\Program Files\Oracle\VirtualBox\VBoxManage.exe" -ArgumentList "controlvm $onevm keyboardputstring C:\Windows\System32\panicshut.bat"
            sleep 3

            ## Press enter
            Start-Process -FilePath "C:\Program Files\Oracle\VirtualBox\VBoxManage.exe" -ArgumentList "controlvm $onevm keyboardputscancode 1c 9c"
            Sleep 3
          }
        Else
          {
            Write-Host "Not Windows VM, Executing CTRL+ALT+L"

            ## Send CTRL+ALT+L
            Start-Process -FilePath "C:\Program Files\Oracle\VirtualBox\VBoxManage.exe" -ArgumentList "controlvm $onevm keyboardputscancode 1d 38 26 9d a8"
            Write-Host "Sending CTRL+ALT+L: Start-Process -FilePath "C:\Program Files\Oracle\VirtualBox\VBoxManage.exe" -ArgumentList \"controlvm $onevm keyboardputscancode 1d 38 26 9d a8\""
            Sleep 2

            ## Send CTRL+ALT+L
            Start-Process -FilePath "C:\Program Files\Oracle\VirtualBox\VBoxManage.exe" -ArgumentList "controlvm $onevm keyboardputscancode 1d 38 26 9d a8"
            Write-Host "Sending CTRL+ALT+L: Start-Process -FilePath "C:\Program Files\Oracle\VirtualBox\VBoxManage.exe" -ArgumentList \"controlvm $onevm keyboardputscancode 1d 38 26 9d a8\""
            Sleep 2
          }
       }
    $tellen = $tellen + 1
    Write-Host $handlevm
  }

## Save the state in one hour
#
sleep 3600

## Same routine, but now for vm shutdown
[Int]$tellen = 0
foreach ($onevm in $allvms)
  {
    Write-Host "Examinating: ". $onevm
    Write-Host "-->"
    If ($onevm.Contains("{"))
      {
        Write-Host "Skip: ". $onevm
        Write-Host "-->"
      }
    else
      {
        Write-Host "Use: ". $onevm
        Write-Host "-->"
        Start-Process -FilePath "C:\Program Files\Oracle\VirtualBox\VBoxManage.exe" -ArgumentList "controlvm $onevm acpipowerbutton"
      }
    $tellen = $tellen + 1
  }

## Now Make sure all VM's are off
sleep 600

## Same routine, but now for vm poweroff
[Int]$tellen = 0
foreach ($onevm in $allvms)
  {
    Write-Host "Examinating: ". $onevm
    Write-Host "-->"
    If ($onevm.Contains("{"))
      {
        Write-Host "Skip: ". $onevm
        Write-Host "-->"
      }
    else
      {
        Write-Host "Use: ". $onevm
        Write-Host "-->"
        Start-Process -FilePath "C:\Program Files\Oracle\VirtualBox\VBoxManage.exe" -ArgumentList "controlvm $onevm poweroff"
      }
    $tellen = $tellen + 1
  }

## Now shutdown the host
sleep 300

Start-Process -FilePath "C:\Windows\System32\shutdown.exe" -ArgumentList "-s -t 300"
