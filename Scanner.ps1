# synthese de voix pour annoncer le chargement des infos
add-type -AssemblyName System.speech
$speak = New-Object System.speech.Synthesis.SpeechSynthesizer
$speak.speak("Veuillez patienter pendant le chargement des informations...")



$content ="<b>INFORMATIONS GENERALES</b></br></br>"
# récupération des infos globales
$hostname = hostname.exe
$infos = Get-ComputerInfo
$winproductname = $infos.WindowsProductName
$winversion = $infos.windowsversion
$winedition = $infos.WindowsEditionId
$processor = $infos.CsProcessors.name
$osinstalldate = $infos.OsInstallDate
$uptime = $infos.OsLastBootUpTime
$ramsize = (Get-CimInstance Win32_PhysicalMemory | Measure-Object -Property capacity -Sum).sum /1gb
$content += "Nom du poste : $hostname<br />"
$content += "Version de l'OS : $winproductname<br />"
$content += "Version de Windows : $winversion<br />"
$content += "Edition de Windows : $winedition<br />"
$content += "Date d'install du système : $osinstalldate<br />"
$content += "Date dernier démarrage : $uptime<br /><br />"
$content += "Processeur installé : $processor<br />"
$content += "Quantité de RAM : $ramsize Go<br /><br />"





$content +="</br></br><b>ANTIVIRUS</b></br></br>"
# Recherche de l'Antivirus installé
$Antivirus_Class = "AntiVirusProduct"
$Win_Antivirus = gwmi -Namespace "root\SecurityCenter2" -Class $Antivirus_Class
$Antivirus_Name = $Win_Antivirus.displayname
$content += "L'antivirus par défaut sur le poste est $Antivirus_Name"
$content += "<br />"
$content += "<br />"
$content += "<br />"






$content +="</br><b>DISQUES ET PARTITIONS</b></br></br>"
# recherche des disques physiques et affichage de leur type (SSD / méca)
$disks = Get-WmiObject -Class MSFT_PhysicalDisk -Namespace root\Microsoft\Windows\Storage  | Select FriendlyName, MediaType
Foreach($Disk in $disks){
    $name = $disk.friendlyname
    $content += "$name"
    switch($disk.MediaType){
        3 {$content += " de type HDD"}
        4 {$content += " de type SSD"}
    }
    $content += "<br />"
}
$content += "<br /><br />"

# Recherche des partitions sur disques physiques et récupération de l'espace disponible
$ListDisk = Get-WmiObject -Class Win32_LogicalDisk | Where-Object {$_.DriveType -eq "3"}
Foreach($Disk in $ListDisk){
   $DiskFreeSpace = ($Disk.freespace/1GB).ToString('F2')
   $TotalSize = ($Disk.size/1GB).ToString('F2')
   $content += "L'espace disque restant sur $($Disk.DeviceID) est de $DiskFreeSpace Go sur $TotalSize Go"
   $content += "<br />"
}




$content +="</br></br></br></br><b>INTERFACES RESEAU</b></br></br>"
# affichage des adaptateurs reseau
$content += Get-NetAdapter | select name, interfacedescription, status, macaddress | ConvertTo-Html
$content += "<br />"

# Affichage des adresses IPv4
$content += Get-NetIPAddress -AddressFamily IPv4 | select ipaddress, interfacealias | ConvertTo-Html
$content += "<br /><br /><br /><br />"






$content +="<b>IMPRIMANTES ET SPOOLER</b></br>"
# Recherche de l'imprimante par défaut et de son statut actuel
$defaultprintername = gwmi win32_printer | where { $_.Default } | select name
$defaultprintername = $defaultprintername.name
$content += "Imprimante par défaut : $defaultprintername<br />"
$code_defaultprinterstatus = gwmi win32_printer | where { $_.Default } | select printerstate
switch($code_defaultprinterstatus.printerstate){
    0{$content+="L'imprimante est en état Idle -- Etat OK"}
    1{$content+="L'imprimante est en état Paused"}
    2{$content+="L'imprimante est en état Error"}
    3{$content+="L'imprimante est en état Pending Deletion"}    
    4{$content+="L'imprimante est en état Paper Jam"}    
    5{$content+="L'imprimante est en état Paper Out"}    
    6{$content+="L'imprimante est en état Manual Feed"}    
    7{$content+="L'imprimante est en état Paper Problem"}    
    8{$content+="L'imprimante est en état Offline"}    
    9{$content+="L'imprimante est en état I/O Active"}
    10{$content+="L'imprimante est en état Busy"}    
    11{$content+="L'imprimante est en état Printing"}    
    12{$content+="L'imprimante est en état Output Bin Full"}    
    13{$content+="L'imprimante est en état Not Available"}    
    14{$content+="L'imprimante est en état Waiting"}    
    15{$content+="L'imprimante est en état Processing"}    
    16{$content+="L'imprimante est en état Initialization"}    
    17{$content+="L'imprimante est en état Warming Up"}    
    18{$content+="L'imprimante est en état Toner Low"}    
    19{$content+="L'imprimante est en état No Toner"}    
    20{$content+="L'imprimante est en état Page Punt"}    
    21{$content+="L'imprimante est en état User Intervention Required"}    
    22{$content+="L'imprimante est en état Out of Memory"}    
    23{$content+="L'imprimante est en état Door Open"}    
    24{$content+="L'imprimante est en état Server_Unknown"}    
    25{$content+="L'imprimante est en état Power Save"}
}
$content += "<br /><br />"

# Affichage du spooler d'impression
$content += "Statistiques des impressions :<br />"
$content += Get-WMIObject Win32_PerfFormattedData_Spooler_PrintQueue | Select Name, @{Expression={$_.jobs};Label="CurrentJobs"}, TotalJobsPrinted, JobErrors | ConvertTo-Html
$content += "<br /><br /><br />"








$content +="<br /><br /><b>LOGS SYSTEME</b></br>"
# Récupération des derniers log en erreur pour le système
$content += "<br />Derniers logs système :"
$content += "<br /><br />"
$content += Get-EventLog system | Where-Object {$_.EntryType -eq "Error"} | select timegenerated, source, eventid, message | select -first 40 | ConvertTo-Html
$content += "<br />"
$content += "<br />"







# envoi par mail
$smtpserver = "my.smtp.server"
$SmtpPort = 25
$subject = "powershell scanner by Proc"
$mailfrom = "powershell-scanner@mysmtp.server"
$mailto = "processus@thiefin.fr"
$Username = "powershell@mysmtp.server"
$Password = "mysupersecurepassword"

$Message = New-Object System.Net.Mail.MailMessage $mailfrom,$mailto
$Message.IsBodyHTML = $true
$Message.Subject = $subject
$Message.Body = $content


$Smtp = New-Object Net.Mail.SmtpClient($smtpserver,$SmtpPort)
$Smtp.EnableSsl = $false
$Smtp.Credentials = New-Object System.Net.NetworkCredential($Username,$Password)
$Smtp.Send($Message)


