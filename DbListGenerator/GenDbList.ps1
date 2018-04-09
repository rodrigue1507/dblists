Param(
    $dirPath = "C:\Users\rpouattara\Desktop\Sample-generate-dblist",
    $inputfile = "GenDbList.config.xlsx",
    [Parameter(Mandatory=$true)] 
    $Env
)


function ParseMainFolders($Env)
{
#Gestion du fichier excel
 


    $Excel = New-Object -ComObject Excel.Application 
    $Excel.visible = $False
    $excel.DisplayAlerts = $False
    $filepath = Join-Path $dirPath $inputfile
    $Workbook = $Excel.Workbooks.open($filepath)
    $workSheetDatabases = $Workbook.Sheets.Item("Databases")
    $workSheetFolders = $Workbook.Sheets.Item("Folders")

#on verifie que l'environnement donné par l'utilisateur existe
    $environmentTab= @{}

    $i = 2 
    do
    {
        $environmentTab.add($i,$workSheetDatabases.cells.item($i,4).text) #boucle sur les environnements et stockage dans le tableau
        $i++
    }while( $workSheetDatabases.cells.item($i,1).text -ne "")

    #on attribut $true à la varible $verifyenv si l'environnement existe sinon false
    $j=2  
    do
    {
        if($environmentTab[$j] -eq $env)
        {
            $Verifyenv = $true
            break
        }
        else
        {
            $Verifyenv = $False
        }
    $j++
    }while($j -lt $environmentTab.Count+2)



#Creation du tableau contries et stockage des pays et creation de la Dblist# 
    if($verifyenv -eq $true)
    {
        Write-Host "Creation en cours..."
        $countries = @{}
        $DbLists = @{}
        $environment =@{}
        $icolFolders = 3

            do
            {
                $countries.Add($icolFolders, $workSheetFolders.cells.Item(1,$icolFolders).Text) 
                $DbLists.Add($workSheetFolders.cells.Item(1,$icolFolders).Text, "")
            $icolFolders++
            }while($worksheetFolders.cells.Item($icolFolders,1).text -ne "")
 
#Stockage des données dans le tableau Dblists


            #Ligne courante
            $i = 2
            #Boucle sur les lignes de Folders
            do
            {

                $folderContent = $workSheetFolders.cells.item($i,1).text #colonne folder dans l'onglet Folders
                $databasecodeContent = $workSheetFolders.cells.item($i,2).text #colonne databasecode dans l'onglet Folders
    
                #Colonne courante
                $j = 3

                #Boucle sur les colonnes de Folders
                do
                {
                    if($workSheetFolders.cells.item($i,$j).text -eq "x")
                    {
                        $ilineDatabases = 2
                        do
                        {            
                            if ($databasecodeContent -eq $worksheetDatabases.cells.item($ilineDatabases,1).text -and $countries[$j] -eq $worksheetDatabases.cells.item($ilineDatabases,3).text -and $env -eq $worksheetDatabases.cells.item($ilineDatabases,4).text)
                            {
                                $Dblists[$j] += ":" + $folderContent + ";" + $worksheetDatabases.cells.item($ilineDatabases,2).text + ";" + $worksheetDatabases.cells.item($ilineDatabases,5).text + "¤"
                        
                            }
                         $ilineDatabases++
                         }while($worksheetDatabases.cells.item($ilineDatabases,1).text -ne "")

                     }
                $j++
                }while($j -lt $countries.Count+2)
            $i++
            }while($workSheetFolders.cells.item($i,1).text -ne "")

    $Excel.Quit()
        
    #Creation de la Dblist par pays
    
    foreach ($country in $DbLists.GetEnumerator())
    {
        if ($country.Value -ne "")
        {

            $DbListName = "DbList_" + $countries[$country.Name] + "-" +  $env + ".txt"
            $dbListPath = Join-Path $dirPath $DbListName 
        
            Set-Content -Path $dbListPath -Value "#:FOLDER;DATABASE;SERVER"  -Force
            Write-Host $DbListName -ForegroundColor Green
            $country.Value.Split("¤") | Foreach{
                
                Add-Content -Path $dbListPath -Value "$_"
         }

         } 
      }
}
if($verifyenv -eq $false)
{
    write-host "E1:environnement inconnu" -BackgroundColor Red -ForegroundColor Yellow 
}                     
}

 ParseMainFolders -env $Env
