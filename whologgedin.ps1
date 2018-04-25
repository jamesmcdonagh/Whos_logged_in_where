#Test if C:\whos_logged_in_where\ directory exists, if not create it.
if((Test-Path C:\whos_logged_in_where) -eq 0){
      New-Item -ItemType Directory -Path C:\whos_logged_in_where | Out-Null
      Write-Host "Created directory C:\whos_logged_in_where"
}

#Create todays file, add CSV headers
New-Item -ItemType File "C:\whos_logged_in_where\$((Get-Date).ToString('yyyy-MM-dd')).csv" -force | Out-Null
$line = "NetBios Name" + "," + "IP Address" + "," + "Users"
Write-Output $line | Out-File "C:\whos_logged_in_where\$((Get-Date).ToString('yyyy-MM-dd')).csv" -Encoding ASCII

#Get list of AD computers, can filter on subnet
$list_of_AD_computers = Get-ADComputer -Properties Name, IPv4Address -Filter * | where {$_.ipv4address -like “10.1.*” } 

$list_of_AD_computers | ForEach-Object {
      if (Test-Connection -ComputerName $_.Name -Count 1 -BufferSize 1 -Quiet){
            # Get-WmiObject is not robust, you will get "RPC server is unavailable" errors  
            #Ping + Get-WmiObject takes ~5-20 seconds per machine       
            $users = Get-WmiObject -Class win32_process -Computername $_.Name | Foreach {$_.GetOwner().User} | Where {$_ -ne "NETWORK SERVICE" -and $_ -ne "LOCAL SERVICE" -and $_ -ne "SYSTEM"} | sort -unique
            $line = $_.Name + "," + $_.IPv4Address + "," + $users
            Write-Output $line | Out-File "C:\whos_logged_in_where\$((Get-Date).ToString('yyyy-MM-dd')).csv" -Append -Encoding ASCII
            Write-Host $_.Name $_.IPv4Address $users -ForegroundColor Green}
      
      else{
            #Ping unsuccessful
            Write-Host $_.Name -ForegroundColor Red}
}