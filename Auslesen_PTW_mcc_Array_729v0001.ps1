# https://github.com/ah7rfu/powershellptwarray/edit/master/Auslesen_PTW_mcc_Array_729v0001.ps1 steht unter der MIT License
#
# Pfad zur PTW Messdatei, Dateiendung .mcc
# Es handelt sich um eine Messdatei vom PTW Array 729

# Informationen zur Erstellung dieses Skriptes stammen, u.a. von http://techgenix.com/read-text-file-powershell/
# weiter interessante Links https://www.youtube.com/watch?v=YrZLCDsh5a0
# https://blog.stefanrehwald.de/2013/03/05/powershell-04-textdatei-auslesen-bearbeiten-anlegen-befullen/
# https://docs.microsoft.com/en-us/previous-versions/windows/it-pro/windows-powershell-1.0/ee176843(v=technet.10)?redirectedfrom=MSDN
# https://powershellexplained.com/2017-03-18-Powershell-reading-and-saving-data-to-files/

$Nachkommastellen = 4

# path - input
$inputClosed = "C:\Users\christoph\Downloads\PTWFiles\0X0 VER 200429 16'16 close.mcc"
$inputopen = "C:\Users\christoph\Downloads\PTWFiles\0X0 VER 200429 16'15.mcc"


# path - output
#$destinationTXT = "C:\Users\christoph\SkriptlokalLaufwerkC\aFiltered.txt"
#$destinationXLS = "C:\Users\christoph\SkriptlokalLaufwerkC\aFiltered.xls"
$destinationCSV = "C:\Users\christoph\SkriptlokalLaufwerkC\aFiltered.csv"

#Getting Content from input files
$contentClosed = Get-Content $inputClosed
$contentopen = Get-Content $inputopen

#Data filtering and extraction
# Measurement Characteristics
$Meas_DateClosed =  $contentClosed | Where-Object {$_ -match 'File_Creation_Date'}
#Measurement Data
$N100P100Closed = $contentClosed | Where-Object {$_ -match '#105'}
$P100P100Closed = $contentClosed | Where-Object {$_ -match '#645'}
$N100Z000Closed = $contentClosed | Where-Object {$_ -match '#95'}
$P100Z000Closed = $contentClosed | Where-Object {$_ -match '#635'}
$N100N100Closed = $contentClosed | Where-Object {$_ -match '#85'}
$P100N100Closed = $contentClosed | Where-Object {$_ -match '#625'}
$Z000P100Closed = $contentClosed | Where-Object {$_ -match '#375'}
$Z000Z000Closed = $contentClosed | Where-Object {$_ -match '#365'}
$Z000N100Closed = $contentClosed | Where-Object {$_ -match '#355'}

$N100P100open = $contentopen | Where-Object {$_ -match '#105'}
$P100P100open = $contentopen | Where-Object {$_ -match '#645'}
$N100Z000open = $contentopen | Where-Object {$_ -match '#95'}
$P100Z000open = $contentopen | Where-Object {$_ -match '#635'}
$N100N100open = $contentopen | Where-Object {$_ -match '#85'}
$P100N100open = $contentopen | Where-Object {$_ -match '#625'}
$Z000P100open = $contentopen | Where-Object {$_ -match '#375'}
$Z000Z000open = $contentopen | Where-Object {$_ -match '#365'}
$Z000N100open = $contentopen | Where-Object {$_ -match '#355'}


#Measurement Data - copying Content to array
#$N100P100ClosedarrayGy = $N100P100Closed.Split()
#$N100P100ClosedmGy = 1000 * $N100P100Closedarray[5]


$P100P100Closedarray = $P100P100Closed.Split()
$N100Z000Closedarray = $N100Z000Closed.Split()
$P100Z000Closedarray = $P100Z000Closed.Split()
$N100N100Closedarray = $N100N100Closed.Split()
$P100N100Closedarray = $P100N100Closed.Split()
$Z000P100Closedarray = $Z000P100Closed.Split()
$Z000Z000Closedarray = $Z000Z000Closed.Split()
$Z000N100Closedarray = $Z000N100Closed.Split()

$N100P100openarray = $N100P100open.Split() 
$P100P100openarray = $P100P100open.Split()
$N100Z000openarray = $N100Z000open.Split()
$P100Z000openarray = $P100Z000open.Split()
$N100N100openarray = $N100N100open.Split()
$P100N100openarray = $P100N100open.Split()
$Z000P100openarray = $Z000P100open.Split()
$Z000Z000openarray = $Z000Z000open.Split()
$Z000N100openarray = $Z000N100open.Split()

#Ausgabe
"`n" #new line
'Datum und Zeit der Messung des geschlossenen Feldes: '+$Meas_DateClosed

'Messwerte bei den Positionen (crossplane [mm] / inplane [mm]) in mGy'
'Closed Field'
'-100, +100:' +"`t"  + $N100P100ClosedmGy + ' mGy'
$P100P100Closedarray[5]
$N100Z000Closedarray[5]
$P100Z000Closedarray[5]
$N100N100Closedarray[5]
$P100N100Closedarray[5]
$Z000P100Closedarray[5]
$Z000Z000Closedarray[5]
$Z000N100Closedarray[5]


"`n" #new line


$N100P100openarray[5]
$P100P100openarray[5]
$N100Z000openarray[5]
$P100Z000openarray[5]
$N100N100openarray[5]
$P100N100openarray[5]
$Z000P100openarray[5]
$Z000Z000openarray[5]
$Z000N100openarray[5]

# Berechnung der Transmissionswerte 
#$N100P100transm = $N100P100Closedarray[5]/$N100P100openarray[5]
$P100P100transm = $P100P100Closedarray[5]/$P100P100openarray[5]
$N100Z000transm = $N100Z000Closedarray[5]/$N100Z000openarray[5]
$P100Z000transm = $P100Z000Closedarray[5]/$P100Z000openarray[5]
$N100N100transm = $N100N100Closedarray[5]/$N100N100openarray[5]

$P100N100transm = $P100N100Closedarray[5]/$P100N100openarray[5]
$Z000P100transm = $Z000P100Closedarray[5]/$Z000P100openarray[5]
$Z000Z000transm = $Z000Z000Closedarray[5]/$Z000Z000openarray[5]
$Z000N100transm = $Z000N100Closedarray[5]/$Z000N100openarray[5]

# Ausgabe der Transmissionswerte unter Eink�rzung der Werte auf 4 Nachkommastellen


"`n" #new line

'Transmissionen bei den Positionen (crossplane [mm] / inplane [mm])'
'-100, +100 :' +"`t"  +"{0:n$Nachkommastellen}" -f $N100P100transm
'+100, +100 :' +"`t"  +"{0:n$Nachkommastellen}" -f $P100P100transm
'-100,    0 :' +"`t"  +"{0:n$Nachkommastellen}" -f $N100Z000transm
'+100,    0 :' +"`t"  +"{0:n$Nachkommastellen}" -f $P100Z000transm
'-100, -100 :' +"`t"  +"{0:n$Nachkommastellen}" -f $N100N100transm
'+100, -100 :' +"`t"  +"{0:n$Nachkommastellen}" -f $P100N100transm
'   0, +100 :' +"`t"  +"{0:n$Nachkommastellen}" -f $Z000P100transm
'   0,    0 :' +"`t"  +"{0:n$Nachkommastellen}" -f $Z000Z000transm
'   0, -100 :' +"`t"  +"{0:n$Nachkommastellen}" -f $Z000N100transm
