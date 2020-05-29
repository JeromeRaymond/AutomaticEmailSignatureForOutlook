# Le but de ce script est de g�n�rer les signatures outlook � la volee
# definition des r�pertoires locaux et distants
$localpath= get-content env:APPDATA # definition du r�pertoire local dans lequel doivent �tre cr�es les signatures
$remotepath='\\RemoteServer\Share\Signatures\Template'  # path pour acc�der au partage qui contient les sources html et les images
$defaultPhone='+33 (0)5.57.00.00.00' # Num de t�l�phone du standard
$defaultfax=''# Num pour le fax


# Controle de pr�sence sur le r�seau
if ((test-path $remotepath) -eq 'true'){ 

# suppression des anciennes signatures A d�commenter si on veut nettoyer le repertoire avant la upload des templates
# C:\Users\$login\AppData\Roaming\Microsoft\Signatures
#remove-item -recurse $localpath\Microsoft\Signatures\*.*
#remove-item -recurse $localpath\Microsoft\Signatures\"EN REPLY_fichiers"\*.*
#remove-item -recurse $localpath\Microsoft\Signatures\EN_fichiers\*.*
#remove-item -recurse $localpath\Microsoft\Signatures\"FR REPLY_fichiers"\*.*
#remove-item -recurse $localpath\Microsoft\Signatures\FR_fichiers\*.*

# r�cup�ration des fichiers templates
robocopy /e /r:3 /w:1 $remotepath\ $localpath\Microsoft\Signatures

# r�cuperation des informations li�es � l'utilisateur
# r�cup�ration de la chaine pour contacter l'AD local
$SysInfo = New-Object -ComObject "ADSystemInfo"
$ADpath2=$SysInfo.GetType().InvokeMember("username", "GetProperty", $Null, $SysInfo, $Null)

# recuperation des informations personnelles
$Info=[ADSI] "LDAP: / / $ADpath2"
$name=$info.displayname # r�cup�ration du display name. Ca evite les soucis avec les accents, changement de nom d'�pouse, ..
$mail=$info.mail # r�cup�ration de l'adresse email
$mobile=$info.mobile # r�cup�ration du num de portable
$adresse=$info.streetaddress # r�cup�ration de l'adresse postale
$postalCode=$info.postalcode # r�cup�ration du code postal
$City=$info.l # r�cup�ration de la ville
$FullAddress= $adresse+"-"+$postalCode+$City # mise en forme de type Adresse-Ville
$PhoneNumber=$info.telephonenumber # r�cup�ration du t�l�phone fixe
$fax=$info.facsimileTelephoneNumber # r�cup�ration du fax
$title=$info.title # r�cup�ration de l'intitul� de poste
$titleUK=$info.department # r�cup�ration de l'intitul� de poste en Anglais. J'avais mis cette info dans le d�partement de l'utilisateur, on peut utiliser les objets personnalis�s
$office=$info.physicalDeliveryOfficeName # r�cup�ration du bureau

# Mise en forme des N� de telephones en version FR et UK. mise en forme annul�e mais gard�e pour compatibilit� avec le reste du script.
if ($PhoneNumber -like '00.33*') {$PhoneNumberform= $PhoneNumber | foreach-object {'+'+ $_.substring(3,2) +' (0)'+ $_.substring(6,1) +'.'+ $_.substring(8,2) +'.'+ $_.substring(11,2) +'.'+ $_.substring(14,2)+'.'+ $_.substring(17,2)}} 
if ($mobile -like '00.44.*') {$mobileform= $mobile} 
if ($mobile -like '00.33.*') {$mobileform= $mobile | foreach-object {'+'+ $_.substring(3,2) +' (0)'+ $_.substring(6,1) +'.'+ $_.substring(8,2) +'.'+ $_.substring(11,2) +'.'+ $_.substring(14,2)+'.'+ $_.substring(17,2)}}



$faxform= $fax
$officeForm = "MyTestCompany - "+$office

# A decommenter en mode pas � pas pour voir si nous r�cup�rons bien les bonnes valeurs
#write-host "Nom et prenom" $name
#write-host "mail" $mail
#write-host "mobile" $mobile
#write-host "tel fixe" $PhoneNumber
#write-host "titre" $title
#write-host "titre UK" $titleUK

#Personnalisation HTML
# creation de la fonction de personnalisation
# $perso est la variable qui va contenir le template html � personnaliser
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
	# de if permettait d'avoir une mise en forme du bureau diff�rente en fonction du bureau (siege vs succursales)
    if ($office -eq 'Siege Social') {
        $perso=$perso -replace '%office%<br />'," "
    } Else {
        $perso=$perso -replace '%office%',"$officeform"
    }
	# Si N� de t�l�phone vide, on supprime la ligne du contenu html afin de ne pas avoir une info du style T�l : et rien derri�re
        if ($PhoneNumber.count -eq 0) {
        $perso=$perso -replace 'Tel : %pager%<br />'," "
    } Else {
        $perso=$perso -replace '%pager%',"$PhoneNumberform"
    }

    $Perso = $Perso -replace "%title%","$title" #personnalisation des intitul�s de poste
    $Perso = $Perso -replace "%titleUK%","$titleUK" #personnalisation des intitul�s de poste
    if ($fax.count -eq 0) {
        $perso=$perso  -replace "%fax%","$defaultfax"
    } Else {
        $perso=$perso  -replace "%fax%","$faxform"
    }
	# Si N� de t�l�phone Mobile vide, on supprime la ligne du contenu html afin de ne pas avoir une info du style T�l : et rien derri�re. Attention dans ce genre de remplacement au nombre d'espace entre dans la ligne
    if ($mobile.count -eq 0) {
        $perso=$perso  -replace '      Mob : %mobile%<br />'," "
    } Else {
    $perso=$perso  -replace '%mobile%',"$mobileform"
    }
    set-content $file $perso #on ecrase le contenu du template par le contenu personnalis�
}

#Personnalisation TXT
# M�me principe que la personnalisation html
 function PersonnalisationTXT ($file)
 { 
    $Perso=Get-Content $file
    $Perso = $Perso -replace "%name%","$name"
    $Perso = $Perso -replace "%mail%","$mail"
    $Perso = $Perso -replace "%adresse%","$FullAddress"
    $Perso = $Perso -replace "%rue%","$Addresse"
    $Perso = $Perso -replace "%codepostal%","$postalcode"
    $Perso = $Perso -replace "%ville%","$city"
	# Si Num de t�l�phone vide, on remplace la ligne par un espace
    if ($PhoneNumber.count -eq 0) {
        $perso=$perso -replace "Tel : %pager%"," "
    } Else {
        $perso=$perso -replace "%pager%","$PhoneNumberform"
    }
    $Perso = $Perso -replace "%title%","$title"
    $Perso = $Perso -replace "%titleUK%","$titleUK"
	# si pas de fax, on utilise le Num par d�faut
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
