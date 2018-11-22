<# 
 .Synopsis
  Crea il file csv in "SA-esami2018_06_20_09-44.csv" da AlmaEsami usando la funzione Get-examsListXml
  richiede powershell 4.0 o superiore

 .Description
  Crea un file csv con tutti gli eami relativi a tre corsi di laurea di Scienze degli Alimenti.
  Il file viene generato da questi link:

	https://corsi.unibo.it/laurea/TecnologieAlimentari/appelli?appelli=
	https://corsi.unibo.it/laurea/ViticolturaEnologia/appelli?appelli=
	https://corsi.unibo.it/magistrale/ScienzeTecnologieAlimentari/appelli?appelli=



   Le pagine citate visualizzano 30 erami per pagina con i relativi appelli. Il programma è composto da due procedure
   esamitotali che prende come parametro gli url riportati e naviga le pagine di visualizzazione degli esami.
   Es. se gli esami sono 42, ci saranno 2 pagine, nella prima  ci compariranno 30 righe e nella seconda 12.
   EsamiTotali richiama la procedura Get-examsListfromXml passandogli l'xml di ogni singola pagina
   

   
  Problemi di lentezza su windows 7 (in 10 tutto ok) creata funzione Get-examsListXml per ovviare...
  https://stackoverflow.com/questions/14202054/why-is-this-powershell-code-invoke-webrequest-getelementsbytagname-so-incred

 .Parameter URL
  esempio di parametro 
  "https://corsi.unibo.it/magistrale/ScienzeTecnologieAlimentari/appelli?appelli="
  todo:Manca il controllo sul sito pagina etc.
  


.OUTPUTS
  Senza modifiche l'output finisce nella directory dello script 
  ToDo: controllo file e folder di detinazione

  .NOTES
  Version:        	1.0.0.0
  Author:         	gastone.canali
  Creation Date:  	29/6/2016 - 19/06/2018 giugno - 13/09/2018 
  Change: 			risolti problemi appelli mancanti - da terminare

 .Example
   # Elenca gli esami i cui nomi iniziano per B
   EsamiTotali  -url "https://corsi.unibo.it/laurea/TecnologieAlimentari/appelli?appelli=&b_start:int=30" 
   esamitotali -url  "https://corsi.unibo.it/laurea/TecnologieAlimentari/appelli?appelli="   
   E' cambiato tutto il sito esami compresi !!!!
 .ToDo
	Aggiungere il parametro DataEsame e implementare gli esami del tal giorno
		esami dal tal giorno in avanti ... es. -dateEsame OGGI file snello solo
		con gli esami a venire...
	levare i &nbsp; all'interno della funzione 
	Header da parametrizzare
	Parametrizzazione e Scelta campi da trasferire nel csv (ora tutti per debug)
	Logging assente mettere un write-debug da qualche parte
 
 .Idea
	Esami tal giorno alla tal ora, forse troppo...
	Sito web pagina dinamicamente creata alla consultazione con rappresentazione gragica delle date occupate ...
	
#>

function convertcsv2excel {
    param(  
        [Parameter(
            Position = 0, 
            Mandatory = $true, 
            ValueFromPipeline = $true,
            ValueFromPipelineByPropertyName = $true)
        ]
		[string]$csvfile,
		[string]$excelfile
    ) 
	 
    
    $xl = new-object -comobject excel.application
    $xl.visible = $true
    $Workbook = $xl.workbooks.open($csvfile)
    ##$Worksheets = $Workbooks.worksheets
    #https://stackoverflow.com/questions/24662980/powershell-excel-finding-string-in-cell-color-row-delete-others
    #Create the missing type
    $MissingType = [System.Type]::Missing
    $Sheet1 = $xl.Worksheets.Item(1)
    $r1 = $xl.Worksheets.Item(1).Range('A1:G1')
    $r1.Font.ColorIndex = 3
    $Sheet1.Name = "Elenco Esami"
    #Create a collection of sheets, separate by comma.
    $colSheets = ($Sheet1)
    #Set the first row to have bold text and set auto filter. Autofit column in the end
    foreach ($sheet in $colSheets) {
        #Lock the first row
        $sheet.Application.ActiveWindow.SplitColumn = 0
        $sheet.Application.ActiveWindow.SplitRow = 1
        $sheet.Application.ActiveWindow.FreezePanes = $true
        ### $xlThin=2 
        #Set other sheet top row defaults
        $Range = $sheet.UsedRange
        #Uncomment to set back color
        #$Range.Interior.ColorIndex = 20
        #Set the font to bold and color to blue
        $Range.Font.ColorIndex = 11
        #$selection = $c.range("A3:C$($DNSResults.Count+2)")
        #
        #$Range.Font.Bold = $True
        #Set autofilter
        #Field, Criteria1, XlAutoFilterOperator Operator, Criteria2, VisibleDropDown
        #XLAutoFilterOperators: xlAnd *default, xlBottom10Items, xlBottom10Percent, xlOr, xlTop10Items, xlTop10Percent
        $Range.AutoFilter(1, $MissingType, 1, $MissingType, $MissingType) | Out-Null
        $Range.EntireColumn.AutoFit() | Out-Null
        $r1 = $xl.Worksheets.Item(1).Range('A1:G1')
        $r1.HorizontalAlignment = -4108
        $r1.Font.ColorIndex = 54
        $r1.Font.size = 12
        $r1.Interior.ColorIndex = 34
        $r1.Font.Bold = $True
    }
    $Workbook.SaveAs("$excelfile", 51)
    $Workbook.Saved = $True
    $xl.Quit()
}

function esamitotali {
    param(  
        [Parameter(
            Position = 0, 
            Mandatory = $true, 
            ValueFromPipeline = $true,
            ValueFromPipelineByPropertyName = $true)
        ]
        [string]$URL
    ) 
    
###
$template=@"
<html xmlns="http://www.w3.org/1999/xhtml">
<head><title>
	Occupazione Aule
</title>


<div id="overlay" class="overlay" style="display: none;"></div>       
<div id="modal_box" style="display: none;"><p class="close_box">x chiudi</p></div>

 <h1>{Facolta:Agraria e Medicina veterinaria}</h1>
 <div id="CurrentTimePanel">

        <div class="header"><span>{Quando:giovedì 8 novembre 2018}</span></div>



<div class="appointment" id='appointment-1'>
<div class="subject">QUALITÀ E INNOVAZIONE NELLE PRODUZIONI PRIMARIE</div>


<div class="teacher">
        luigi filippo d'antuono
    </div>

<div class="corso">
        <p>{Corso*:SCIENZE E TECNOLOGIE ALIMENTARI [LM]}</p>
    </div>
<div class="detail-content" style="display:none;">
    <div class="detail-title">
        <strong>{Esame*:66048 - QUALITÀ E INNOVAZIONE NELLE PRODUZIONI PRIMARIE}</strong>
    </div>                            
    <div class="teacher">
            <strong>luigi filippo d'antuono</strong>
        </div>
    <div class="time">9:00 - 11:00</div>
    <div class="corso">
            <p>8531 - SCIENZE E TECNOLOGIE ALIMENTARI [LM]</p>
        </div>
</div>
</div>

                            </div>
                        </div>
                    </div>
                </div>
            </div>
        </div></td><td><div class="rsWrap" style="z-index:22;">
            <div id="DayResourceScheduler_0_0" title="PRODUZIONI ANIMALI (C.I.) - ZOOCOLTURE" class="rsApt" style="height:148px;width:100%;left:0%;">
                <div class="rsAptOut">
                    <div class="rsAptMid">
                        <div class="rsAptIn">
                            <div class="rsAptContent">

"@
$template1=@"
<html xmlns="http://www.w3.org/1999/xhtml">
<head><title>
	Occupazione Aule
</title>

    <style type="text/css">

<body>

<div class="appointment" id='appointment-4'>
<div class="subject">CONDIZIONAMENTO E IMBALLAGGIO</div>


<div class="teacher">
        santina romani
    </div>

<div class="corso">
        <p>{Corso*:8531 - SCIENZE E TECNOLOGIE ALIMENTARI [LM]}</p>
    </div>

<div class="detail-content" style="display:none;">
    <div class="detail-title">
        <strong>{Esame*:69168 - CONDIZIONAMENTO E IMBALLAGGIO}</strong>
    </div>                            
    <div class="teacher">
            <strong>{Prof:santina romani}</strong>
        </div>
    <div class="time">{Ora:10:00 - 13:00}</div>
    <div class="corso">
            <p>8531 - SCIENZE E TECNOLOGIE ALIMENTARI [LM]</p>
        </div>
</div>
</div>

                            </div>
                        </div>
                    </div>
                </div>
            </div>
        </div></td><td>&nbsp;</td>
    </tr><tr class="rsAlt" style="height:38px;">
        <td>&nbsp;</td><td>&nbsp;</td><td>&nbsp;</td><td>&nbsp;</td><td>&nbsp;</td><td>&nbsp;</td>
    </tr><tr style="height:38px;">
        <td><div class="rsWrap" style="z-index:18;">
            <div id="DayResourceScheduler_6_0" title="NUTRIZIONE UMANA" class="rsApt" style="height:148px;width:100%;left:0%;">
                <div class="rsAptOut">
                    <div class="rsAptMid">
                        <div class="rsAptIn">
                            <div class="rsAptContent">

"@



$prop=@"
[
{"key":"38714","text":"MASSIMILIANO PETRACCI","type":"Teacher","available":true,"cssClass":"","internalKey":"/wEFBTM4NzE0Cqu5YMBxjXyF1eId76/K8/+r9WfUwXOmj4tifP7uNEQ="},
{"key":"30592","text":"LUIGI FILIPPO D\\u0027ANTUONO","type":"Teacher","available":true,"cssClass":"","internalKey":"/wEFBTMwNTkyvRTqTxjDz2RZohfwepH9762HJ5XrpR+lBvD5ZpuURTM="},
{"key":"31427","text":"MARCO DALLA ROSA","type":"Teacher","available":true,"cssClass":"","internalKey":"/wEFBTMxNDI3tw/p6m7xQ6TDw0PbnO49vgd9c+rmT09VOSbMUEiB3X0="},
{"key":"16711","text":"CLAUDIO CIAVATTA","type":"Teacher","available":true,"cssClass":"","internalKey":"/wEFBTE2NzExUR6jBc4P16JthbFEtfkh8FA6rS9he1ZhD1nkkcC19mE="},
{"key":"35827","text":"SANTINA ROMANI","type":"Teacher","available":true,"cssClass":"","internalKey":"/wEFBTM1ODI3+puwHjqEFHfaYR8oPNtsShlwvbvkKRGUpiSif4Nh8vw="},
{"key":"24415","text":"FAUSTO GARDINI","type":"Teacher","available":true,"cssClass":"","internalKey":"/wEFBTI0NDE1SMeeuPEk+oyGCy9eDWihn6L0bgzE/P+4BKr5luZ8LaE="},
{"key":"35180","text":"ALESSANDRA BORDONI","type":"Teacher","available":true,"cssClass":"","internalKey":"/wEFBTM1MTgwCFcFfrA/z41W34rhAjTJhPPbrwHw9rrusZIuDaSWfsI="},
{"key":"23242","text":"MARIA CABONI","type":"Teacher","available":true,"cssClass":"","internalKey":"/wEFBTIzMjQyZTHjJq1t5lbv8OW+tyWABgyFhUsOyXim0UwStXssRNM="},
{"key":"31414","text":"ANGELO FABBRI","type":"Teacher","available":true,"cssClass":"","internalKey":"/wEFBTMxNDE0JiXGdDD7RfBDUMd4YAST1MoiaEMr7OiWte3OMUB2Z78="},
{"key":"31064","text":"CLAUDIO RIPONI","type":"Teacher","available":true,"cssClass":"","internalKey":"/wEFBTMxMDY0m5azBRC6XPvUjvQMDXl2mJ97fYT/JkUW6CGir7r5QV0="},
{"key":"65_WPTE_041","text":"AULA A","type":"Room","available":true,"cssClass":"","internalKey":"/wEFCzY1X1dQVEVfMDQxy/XeBn2QxiwGTXUgfbPh6d6RIQCAb3srmeW1FutrcGQ=","attributes":
{"Piano":"Piano Terra","Edificio":"Villa Almerici","Ubicazione":"Piazza Goidanich, 60 - Cesena","SortOrder":"65;AULA A"}},
{"key":"65_WPTE_059","text":"AULA B","type":"Room","available":true,"cssClass":"","internalKey":"/wEFCzY1X1dQVEVfMDU5wc5e0U5rGAqOkg+utee6d3qpepOQvpM6Bj3dm4HFsFQ=","attributes":
{"Piano":"Piano Terra","Edificio":"Villa Almerici","Ubicazione":"Piazza Goidanich, 60 - Cesena","SortOrder":"65;AULA B"}},
{"key":"65_WPTE_065","text":"AULA C","type":"Room","available":true,"cssClass":"","internalKey":"/wEFCzY1X1dQVEVfMDY1v1NRJFHyPJrE7/EbzzFgmHTYXhio97uo81b+q7ojuhY=","attributes":
{"Piano":"Piano Terra","Edificio":"Villa Almerici","Ubicazione":"Piazza Goidanich, 60 - Cesena","SortOrder":"65;AULA C"}},
{"key":"65_WPTE_068","text":"AULA D","type":"Room","available":true,"cssClass":"","internalKey":"/wEFCzY1X1dQVEVfMDY42Ffn7M78NjAIp2+g5J+UAkppfRDls1MBGIifhKAaL4Y=","attributes":
{"Piano":"Piano Terra","Edificio":"Villa Almerici","Ubicazione":"Piazza Goidanich, 60 - Cesena","SortOrder":"65;AULA D"}},
{"key":"65_WPTE_036","text":"AULA F","type":"Room","available":true,"cssClass":"","internalKey":"/wEFCzY1X1dQVEVfMDM2QFoGCYNNuEN1ylgUrfnbsSF3AABSWH/hiDAUK3sFmFc=","attributes":
{"Piano":"Piano Terra","Edificio":"Villa Almerici","Ubicazione":"Piazza Goidanich, 60 - Cesena","SortOrder":"65;AULA F"}},
{"key":"65_WPTE_055","text":"AULA MAGNA","type":"Room","available":true,"cssClass":"","internalKey":"/wEFCzY1X1dQVEVfMDU19zqDGd1PFnflmcu71qeYMHFkvR5C1f80RyWM/DMY4Dc=","attributes":
{"Piano":"Piano Terra","Edificio":"Villa Almerici","Ubicazione":"Piazza Goidanich, 60 - Cesena","SortOrder":"65;AULA MAGNA"}},
{"key":"8528","text":"TECNOLOGIE ALIMENTARI [L]","type":"Corso","available":true,"cssClass":"","internalKey":"/wEFBDg1MjjOCL0xIb/JEjux/W1t0ZpQVus6CHaxM49vumHliYGIUA=="},
{"key":"8531","text":"SCIENZE E TECNOLOGIE ALIMENTARI [LM]","type":"Corso","available":true,"cssClass":"","internalKey":"/wEFBDg1MzHQnTiAzvmBRV2IAQLT+bEvVvN9hZG8BzL5Od0ON4w2cw=="},
{"key":"8527","text":"VITICOLTURA ED ENOLOGIA [L]","type":"Corso","available":true,"cssClass":"","internalKey":"/wEFBDg1Mjcebs+X8YcbIjbE5AsjZxRG7H0RQL0tOUAEodUevYSrsg=="}]
"@
$ob=convertfrom-json -InputObject $prop
    #https://corsi.unibo.it/laurea/TecnologieAlimentari/appelli?appelli=&b_start:int=90
    #https://www.unibo.it/uniboweb/utils/orariolezioni/calendario.aspx?Scuola=1&Corso=8528,8527,8531&Edificio=65&Data=04/10/2018
    $xmlexa = ''
    $WebResponse = Invoke-WebRequest $URL
    $Cleaned=($webresponse.content).trim() -replace '\t|&nbsp;'
    $Cleaned=($webresponse.content).trim() -replace '\t|\  '

    $new=ConvertFrom-String -InputObject $Cleaned -TemplateContent $template -OutVariable exares

    $exares|ForEach-Object{  $_.corso}
    #estrae
    $start='//<![CDATA['
    $end='//]]>'
    $PAGE=$webresponse.Content
    #estrae fra parentesi graffe
    $regex="(?smi)\{(?:(?:\{(?:(?:\{(?:[^{}])*\})|(?:[^{}]))*\})|(?:[^{}]))*\}"
    $ma=$PAGE -match $regex
    #ora


    $regex1='(?m)\"(?:(?:\"(?:(?:\"(?:[^{}])*\")|(?:[^{}]))*\")|(?:[^{}]))*\"'
    $regex2='(?m)\[(?:(?:\[(?:(?:\[(?:[^[]])*\[)|(?:[^[]]))*\[)|(?:[^[]]))*\]'
    $ma=$PAGE.Substring(1, $PAGE.length-1)-match $regex
    $pclea=($page -replace "`r`n") -replace '\\"',"'"
    $r = $pclea  -split '"'
    $pclea -match '//<\!\[CDATA\[.*//\]\]>'

    $M=([regex]$regex).matches($pclea)
    $N=($pclea -match '\<\!\[CDATA\[(/?[^\]\]\>+)\]\]\>')
    $cleanpage=$page -replace '\"',"'"
    $splitted= $cleanpage -split '"'
    $sindex=$PAGE.IndexOf($start)
    $eIndex = $PAGE.IndexOf($end)
	$partialText,$xml = ''
    if(($sIndex -ge 0) -and ($eIndex -ge 0)) {
        $partialText = $PAGE.Substring($sIndex, $eIndex - $sIndex)
        $regex="^[^()]*(?>(?>(?'open'\()[^()]*)+(?>(?'-open'\))[^()]*)+)+(?(open)(?!))$"

        $ma1=$partialText -replace $regex
    }
    #Seleziona il blocco con gli esami e lo trasforma in xml
    #<TABLE style="WIDTH: 100%">
    $xmlexa1 = $webresponse.ParsedHtml.getElementsByTagName('DIV')
    $appelli=$xmlexa1|Where-Object {$_.getAttributeNode('class').Value -eq 'rsAptContent'}
    $xmlexa0 = ($webresponse.ParsedHtml.getElementsByTagName('H1')).outerhtml
    $xmlexa2=($webresponse.ParsedHtml.getElementsByTagName('H1')).innerhtml
    [xml]$xmlexa = ($WebResponse.AllElements | Where-Object {$_ -match 'H1'}).outerHTML
    # conta gli esami da elaborare
    $esamitotali = $esamiinpagina = ($xmlexa.SelectNodes("//a[@class='openclose-appelli']")	).Count
    # esami in pagina 1
    $pagina = 1
    # Estrae gli esami di pagina uno
    Get-examsListfromXml $xmlexa
	
	# Cicla tutte le pagine con gli esami
	# ho qualche dubbio se gli esami totali sono es. 120
	# perchè non ho usato la variabile $esamitoali... ci guarderemo
    while ($esamiinpagina -eq 30) {
        $xmlexa = ''
		# preleva pagina le pagine seguenti degli esami
		# nella prima pagina ci sono 30 appelli, la stringa seguente prende quelli da 30 a 60
		# tutto in funzione del numero di pagine rileato in precedenza
        $WebResponse = Invoke-WebRequest ( $URL + '&b_start:int=' + (30 * $pagina).ToString() )
        #Seleziona il blocco con gli esami e lo trasforma in xml
        [xml]$xmlexa = ($WebResponse.AllElements | Where-Object {$_.class -match 'dropdown-component'}).outerHTML
        # esami in pagina
        $esamiinpagina = ($xmlexa.SelectNodes("//a[@class='openclose-appelli']")	).Count	
        #conta gli esami totali (non serve ma per debug ...)
        $esamitotali = $esamitotali + $esamiinpagina
        # Passa alla pagina seguente
        $pagina++
        write-debug ((get-date).tostring() + "Esami pag. $esamiinpagina. Esami Tot.: $esamitotali")
        # Estrae gli esami di pagina seguente
        Get-examsListfromXml -xml $xmlexa
    }
}



function Get-examsListfromXml {
    param(  
        [Parameter(
            Position = 0, 
            Mandatory = $true, 
            ValueFromPipeline = $true,
            ValueFromPipelineByPropertyName = $true)
        ]
        [xml]$xml 
    ) 
    #preleva la parte xml relativa agli appelli
    $appelli = $xml.SelectNodes("//div[@class='items-container']")	
    #preleva la parte xml relativa agli esami
    $esami = $xml.SelectNodes("//a[@class='openclose-appelli']")	
		
    write-debug ((get-date).tostring() + " fine appelli")
    # Cicla fra gli esami
    for ($i = 0; $i -le $esami.count - 1; $i++) {
        # seleziona un singolo appello
        [xml]$appello = $appelli[$i].outerXml -replace '&nbsp;'
        $curAppello = 0
        # istanzia vuote
        $dataeora = $nome = $comp = $doc = $cod = $luogo = $tipo = ''
        # preleva tutti gli appelli dell'esame $i
        $appellifull = (Select-Xml -Xml $appello -XPath "//table[@class='single-item']/tbody/tr/th|//table[@class='single-item']/tbody/tr/td"  ).Node 
        #codice esame
        $cod = $esami[$i].span[0].'#text'
        #nome docente
        $doc = $esami[$i].span[1].'#text'
        #nome esame
        $nome = ($esami[$i].'#text').trim()
			
        #Cicla fra gli appelli delle'esame i esimo
        While ($curAppello -le $appellifull.Count - 1) {
            $val = $appellifull[$curAppello].'#text' 
				
            switch  -Regex  ($val) {
                'Data e ora' {   
												 
                    if (-not($dataeora -eq '')) { 
                        #scrivo il record 
                        "$dataeora;$nome;$comp;$doc;$cod;$luogo;$tipo"
                    }
                    # Svuoto DataeOra et al
                    $dataeora = $comp = $luogo = $tipo = ''
                    $curAppello++
                    # Tolgo la parola ore per avere un formato data e ora compatibile con excel
                    $dataeora = ($appellifull[$curAppello].'#text').trim() -replace '\s+|ore', ' '
                    break  
                }
                'Componente' { 
                    $curAppello++
                    $comp = ($appellifull[$curAppello].'#text').trim() -replace '\s+', ' '
                    break  
                }
                'Lista iscrizioni'	{ 
                    $curAppello++
                    $DalAl = 'Dal ' + $appellifull[$curAppello].span[0] + ' al ' + $appellifull[$curAppello].span[1]
                    break  
                }
                'Tipo prova' { 
                    $curAppello++
                    $tipo = ($appellifull[$curAppello].'#text').trim() -replace '\s+', ' '
                    break  
                }
                'Luogo' { 
                    $curAppello++
                    $luogo = ($appellifull[$curAppello].'#text').trim() -replace '\s+', ' '
                    break  
                }
                'Note' { 	
                    $curAppello++
                    $note = ($appellifull[$curAppello].'#text').trim() -replace '\s+', ' '
                    break  
                }     
            }
            # applello seguente
            #write-debug ("$dataeora;$nome;$comp;$doc;$cod;$luogo;$tipo") 
            $curAppello++
        }
        "$dataeora;$nome;$comp;$doc;$cod;$luogo;$tipo"
        write-debug ($i.tostring() + "--- $dataeora;$nome;$comp;$doc;$cod;$luogo;$tipo")
    }
}  

#--------------------------------------------------------------------------------------
# Main
#--------------------------------------------------------------------------------------
Clear-Host

$scriptPath = split-path -parent $MyInvocation.MyCommand.Definition
$DebugPreference = "Continue"
#$ProgressPreference = 'SilentlyContinue'
#start-transcript  -path "$scriptPath\log.log" -NoClobber

# ora attuale
$ora = get-date -f "yyyy_MM_dd_HH-mm"

# Da dove prelevo gli esami
$Tec = "https://corsi.unibo.it/laurea/TecnologieAlimentari/appelli?appelli="   
$eno = "https://corsi.unibo.it/laurea/ViticolturaEnologia/appelli?appelli="    
$mag = "https://corsi.unibo.it/magistrale/ScienzeTecnologieAlimentari/appelli?appelli=" 
$ORARIO="https://www.unibo.it/uniboweb/utils/orariolezioni/calendario.aspx?Scuola=1&Corso=8528,8527,8531&Edificio=65&Data=04/10/2018"
$ORARIO="https://www.unibo.it/uniboweb/utils/orariolezioni/calendario.aspx?Scuola=1&Corso=8528,8527,8531&Edificio=65&Data=08/11/2018"
# Nome del file csv generato
$outfile = "$scriptPath\SA-esami$ora.csv"
# Header del csv
###"dataeora;nome;componente;docente;codice;luogo;tipo"  |out-file $outfile -Append -Encoding "UTF8" 
###esamitotali $Tec                                      |out-file $outfile -Append -Encoding "UTF8" 
###esamitotali $eno                                      |out-file $outfile -Append -Encoding "UTF8" 
###esamitotali $mag                                      |out-file $outfile -Append -Encoding "UTF8" 
#Fine. Csv Creato.

esamitotali $orario                                      |out-file $outfile -Append -Encoding "UTF8" 

"--- Creo il file excel --- Aggiunta Dopo, così come veniva ... -------------------------------------------------------------"
# Nome del nuovo file excel  
$newexcelfile = "$scriptPath\SA-esami$ora.xlsx"
# Converto il csv in xslx con qualche formattazione
convertcsv2excel -csvfile  $outfile -excelfile  $newexcelfile
# Rimuovo il csv
Remove-Item $outfile
