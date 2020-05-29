# Le but de ce script est de générer les signatures outlook à la volee
# definition des répertoires locaux et distants
$localpath= get-content env:APPDATA # definition du répertoire local dans lequel doivent être crées les signatures
$remotepath='\\RemoteServer\Share\Signatures\Template'  # path pour accéder au partage qui contient les sources html et les images
$defaultPhone='+33 (0)5.57.00.00.00' # Num de téléphone du standard
$defaultfax=''# Num pour le fax


# Controle de présence sur le réseau
if ((test-path $remotepath) -eq 'true'){ 

# suppression des anciennes signatures A décommenter si on veut nettoyer le repertoire avant la upload des templates
# C:\Users\$login\AppData\Roaming\Microsoft\Signatures
#remove-item -recurse $localpath\Microsoft\Signatures\*.*
#remove-item -recurse $localpath\Microsoft\Signatures\"EN REPLY_fichiers"\*.*
#remove-item -recurse $localpath\Microsoft\Signatures\EN_fichiers\*.*
#remove-item -recurse $localpath\Microsoft\Signatures\"FR REPLY_fichiers"\*.*
#remove-item -recurse $localpath\Microsoft\Signatures\FR_fichiers\*.*

# récupération des fichiers templates
robocopy /e /r:3 /w:1 $remotepath\ $localpath\Microsoft\Signatures

# récuperation des informations liées à l'utilisateur
# récupération de la chaine pour contacter l'AD local
$SysInfo = New-Object -ComObject "ADSystemInfo"
$ADpath2=$SysInfo.GetType().InvokeMember("username", "GetProperty", $Null, $SysInfo, $Null)

# recuperation des informations personnelles
$Info=[ADSI] "LDAP: / / $ADpath2"
$name=$info.displayname # récupération du display name. Ca evite les soucis avec les accents, changement de nom d'épouse, ..
$mail=$info.mail # récupération de l'adresse email
$mobile=$info.mobile # récupération du num de portable
$adresse=$info.streetaddress # récupération de l'adresse postale
$postalCode=$info.postalcode # récupération du code postal
$City=$info.l # récupération de la ville
$FullAddress= $adresse+"-"+$postalCode+$City # mise en forme de type Adresse-Ville
$PhoneNumber=$info.telephonenumber # récupération du téléphone fixe
$fax=$info.facsimileTelephoneNumber # récupération du fax
$title=$info.title # récupération de l'intitulé de poste
$titleUK=$info.department # récupération de l'intitulé de poste en Anglais. J'avais mis cette info dans le département de l'utilisateur, on peut utiliser les objets personnalisés
$office=$info.physicalDeliveryOfficeName # récupération du bureau

# Mise en forme des N° de telephones en version FR et UK. mise en forme annulée mais gardée pour compatibilité avec le reste du script.
if ($PhoneNumber -like '00.33*') {$PhoneNumberform= $PhoneNumber | foreach-object {'+'+ $_.substring(3,2) +' (0)'+ $_.substring(6,1) +'.'+ $_.substring(8,2) +'.'+ $_.substring(11,2) +'.'+ $_.substring(14,2)+'.'+ $_.substring(17,2)}} 
if ($mobile -like '00.44.*') {$mobileform= $mobile} 
if ($mobile -like '00.33.*') {$mobileform= $mobile | foreach-object {'+'+ $_.substring(3,2) +' (0)'+ $_.substring(6,1) +'.'+ $_.substring(8,2) +'.'+ $_.substring(11,2) +'.'+ $_.substring(14,2)+'.'+ $_.substring(17,2)}}



$faxform= $fax
$officeForm = "MyTestCompany - "+$office

# A decommenter en mode pas à pas pour voir si nous récupérons bien les bonnes valeurs
#write-host "Nom et prenom" $name
#write-host "mail" $mail
#write-host "mobile" $mobile
#write-host "tel fixe" $PhoneNumber
#write-host "titre" $title
#write-host "titre UK" $titleUK

#Personnalisation HTML
# creation de la fonction de personnalisation
# $perso est la variable qui va contenir le template html à personnaliser
# le but est de remplacer des variables du style %xxx% contenue dans le template
 function PersonnalisationHTML ($file)
 { 
    $Perso=Get-Content $file
    $Perso = $Perso -replace "%name%","$name" # remplacement de %name% par le displayname d'AD
    $Perso = $Perso -replace "%mail%","$mail"
    $Perso = $Perso -replace "%adresse%","$FullAddress"
    $Perso = $Perso -replace "%rue%","$Addresse"
    $Perso = $Perso -replace "%codepostal%","$postalcode"
    $Perso = $Perso -replace "%ville%","$city"
	# de if permettait d'avoir une mise en forme du bureau différente en fonction du bureau (siege vs succursales)
    if ($office -eq 'Siege Social') {
        $perso=$perso -replace '%office%<br />'," "
    } Else {
        $perso=$perso -replace '%office%',"$officeform"
    }
	# Si N° de téléphone vide, on supprime la ligne du contenu html afin de ne pas avoir une info du style Tél : et rien derriére
        if ($PhoneNumber.count -eq 0) {
        $perso=$perso -replace 'Tel : %pager%<br />'," "
    } Else {
        $perso=$perso -replace '%pager%',"$PhoneNumberform"
    }

    $Perso = $Perso -replace "%title%","$title" #personnalisation des intitulés de poste
    $Perso = $Perso -replace "%titleUK%","$titleUK" #personnalisation des intitulés de poste
    if ($fax.count -eq 0) {
        $perso=$perso  -replace "%fax%","$defaultfax"
    } Else {
        $perso=$perso  -replace "%fax%","$faxform"
    }
	# Si N° de téléphone Mobile vide, on supprime la ligne du contenu html afin de ne pas avoir une info du style Tél : et rien derriére. Attention dans ce genre de remplacement au nombre d'espace entre dans la ligne
    if ($mobile.count -eq 0) {
        $perso=$perso  -replace '      Mob : %mobile%<br />'," "
    } Else {
    $perso=$perso  -replace '%mobile%',"$mobileform"
    }
    set-content $file $perso #on ecrase le contenu du template par le contenu personnalisé
}

#Personnalisation TXT
# Même principe que la personnalisation html
 function PersonnalisationTXT ($file)
 { 
    $Perso=Get-Content $file
    $Perso = $Perso -replace "%name%","$name"
    $Perso = $Perso -replace "%mail%","$mail"
    $Perso = $Perso -replace "%adresse%","$FullAddress"
    $Perso = $Perso -replace "%rue%","$Addresse"
    $Perso = $Perso -replace "%codepostal%","$postalcode"
    $Perso = $Perso -replace "%ville%","$city"
	# Si Num de téléphone vide, on remplace la ligne par un espace
    if ($PhoneNumber.count -eq 0) {
        $perso=$perso -replace "Tel : %pager%"," "
    } Else {
        $perso=$perso -replace "%pager%","$PhoneNumberform"
    }
    $Perso = $Perso -replace "%title%","$title"
    $Perso = $Perso -replace "%titleUK%","$titleUK"
	# si pas de fax, on utilise le Num par défaut
    if ($fax.count -eq 0) {
        $perso=$perso  -replace "%fax%","$defaultfax"
    } Else {
        $perso=$perso  -replace "%fax%","$faxform"
    }
    if ($mobile.count -eq 0) {
        $perso=$perso  -replace "Mob : %mobile%"," "
    } Else {
    $perso=$perso  -replace "%mobile%","$mobileform"
    }
    set-content $file $perso
}


# Personnalisation des fichiers HTML
PersonnalisationHTML ($localpath+'\Microsoft\Signatures\FR.htm')
PersonnalisationHTML ($localpath+'\Microsoft\Signatures\EN.htm')
PersonnalisationHTML ($localpath+'\Microsoft\Signatures\FR_Reply.htm')
PersonnalisationHTML ($localpath+'\Microsoft\Signatures\EN_Reply.htm')


# Personnalisation des fichiers TXT
PersonnalisationTXT ($localpath+'\Microsoft\Signatures\FR.txt')
PersonnalisationTXT ($localpath+'\Microsoft\Signatures\EN.txt')
PersonnalisationTXT ($localpath+'\Microsoft\Signatures\FR_Reply.txt')
PersonnalisationTXT ($localpath+'\Microsoft\Signatures\EN_Reply.txt')
}

# sortie de script
exit
