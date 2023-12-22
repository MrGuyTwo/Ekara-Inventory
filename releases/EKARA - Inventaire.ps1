#####################################################################################################
#                           Example of use of the EKARA API                                         #
#####################################################################################################
# Swagger interface : https://api.ekara.ip-label.net                                                #
# Personalized parameters : username / password / Debug mode                                        #
# Purpose of the script : Full inventory                                                            #
#####################################################################################################
# Author : Guy Sacilotto
# Last Update : 18/05/2023
# Version : 8.3

<#
Authentication :  user / password
Methods called : 
    auth/login    
    adm-api/clients
    adm-api/client/users
    adm-api/scenarios
    adm-api/plannings
    adm-api/alerts
    script-api/script
    script-api/scripts
    adm-api/applications
    adm-api/workspaces
    adm-api/zones
    adm-api/sites
    adm-api/actions
    rum-restit/trk
    rum-restit/metrics
    rum-restit/trk/$trackerID/results/$metricID/overview
    rum-restit/trk/$trackerID/urlgroups
    rum-restit/customdim/business
    rum-restit/customdim/custom
    rum-restit/customdim/infra
    rum-restit/customdim/version
    adm-api/reports/views
    adm-api/reports/schedules
#>

Clear-Host

#region VARIABLES
#========================== SETTING THE VARIABLES ===============================
$error.clear()
add-type -assemblyName "Microsoft.VisualBasic"
$PSDefaultParameterValues['*:Encoding'] = 'utf8'
$global:API = "https://api.ekara.ip-label.net"                                                # Webservice URL
$global:UserName = "xxxxxxxxxxxxxxxx"                                                         # EKARA Account
$global:PlainPassword = "xxxxxxxxxxxx"                                                        # EKARA Password
$global:Result_OK = 0
$global:Result_KO = 0
[String]$global:date = [DateTime]::Now.ToString("yyyy-MM-dd HH-mm-ss")                        # Récupère la date du jour

$global:headers = $null
$global:headers = New-Object "System.Collections.Generic.Dictionary[[String],[String]]"       # Create Header
$headers.Add("Accept","application/json")                                                     # Setting Header
$headers.Add("Content-Type","application/json")                                               # Setting Header

# Recherche le chemin du script
if ($psISE) {
    [String]$global:Path = Split-Path -Parent $psISE.CurrentFile.FullPath
    if($Debug -eq $true){Write-Host "Path ISE = $Path" -ForegroundColor Yellow}
} else {
    #[String]$global:Path = split-path -parent $MyInvocation.MyCommand.Path
    [String]$global:Path = (Get-Item -Path ".\").FullName
    if($Debug -eq $true){Write-Host "Path Direct = $Path" -ForegroundColor Yellow}
}

# Authentication choice
    # 1 = Without asking for an account and password (you must configure the account and password in this script.)
    # 2 = Request the entry of an account and a password (default)
    $global:Auth = 2
    $global:Debug = $False     # Value : True ou False
#endregion

#region Functions
function Authentication{
    try{
        Switch($Auth){
            1{
                # Without asking for an account and password
                if(($null -ne $UserName -and $null -ne $PlainPassword) -and ($UserName -ne '' -and $PlainPassword -ne '')){
                    Write-Host "--- Automatic AUTHENTICATION (account) ---------------------------" -BackgroundColor Green
                    $uri = "$API/auth/login"                                                                                        # Webservice Methode
                    $response = Invoke-RestMethod -Uri $uri -Method POST -Body @{ email = "$UserName"; password = "$PlainPassword"} # Call WebService method
                    $global:Token = $response.token                                                                                        # Register the TOKEN
                    $global:headers.Add("authorization","Bearer $Token")                                                            # Adding the TOKEN into header
                }Else{
                    Write-Host "--- Account and Password not set ! ---------------------------" -BackgroundColor Red
                    Write-Host "--- To use this connection mode, you must configure the account and password in this script." -ForegroundColor Red
                    exit
                }
            }
            2{
                # Requests the entry of an account and a password (default) 
                Write-Host "------------------------------ AUTHENTICATION with account entry ---------------------------" -ForegroundColor Green
                $MyAccount = $Null
                $MyAccount = Get-credential -Message "EKARA login account" -ErrorAction Stop               # Request entry of the EKARA Account
                if(($null -ne $MyAccount) -and ($MyAccount.password.Length -gt 0)){
                    $UserName = $MyAccount.GetNetworkCredential().username
                    $PlainPassword = $MyAccount.GetNetworkCredential().Password
                    $uri = "$API/auth/login"
                    $response = Invoke-RestMethod -Uri $uri -Method POST -Body @{ email = "$UserName"; password = "$PlainPassword"} # Call WebService method
                    $Token = $response.token                                                               # Register the TOKEN
                    $global:headers.Add("authorization","Bearer $Token")
                }Else{
                    Write-Host "--- Account and password not specified ! ---------------------------" -BackgroundColor Red
                    Write-Host "--- To use this connection mode, you must enter Account and password." -ForegroundColor Red
                    exit
                }
            }
        }
    }Catch{

        Write-Host "-------------------------------------------------------------" -ForegroundColor red 
        Write-Host "Erreur ...." -BackgroundColor Red
        Write-Host $Error.exception.Message[0]
        Write-Host $Error[0]
        Write-host $error[0].ScriptStackTrace
        Write-Host "-------------------------------------------------------------" -ForegroundColor red

        Error_popup($Error[0])
        Break
    }
}

function Hide-Console{
    # .Net methods Permet de réduire la console PS dans la barre des tâches
    Add-Type -Name Window -Namespace Console -MemberDefinition '
    [DllImport("Kernel32.dll")]
    public static extern IntPtr GetConsoleWindow();

    [DllImport("user32.dll")]
    public static extern bool ShowWindow(IntPtr hWnd, Int32 nCmdShow);
    '
    $consolePtr = [Console.Window]::GetConsoleWindow()
    #0 hide / 1 normal / 2 réduit 
    [Console.Window]::ShowWindow($consolePtr, 2)
}

Function List_Clients{
    try{
        #========================== adm-api/clients =============================
        Write-Host "-------------------------------------------------------------" -ForegroundColor green
        Write-Host "------------------- Liste tous les client  -------------------" -BackgroundColor "White" -ForegroundColor "DarkCyan"
        $uri ="$API/adm-api/clients"
        $clients = Invoke-RestMethod -Uri $uri -Method POST -Headers $headers 
        $count = $clients.count

        Add-Type -AssemblyName System.Windows.Forms
        Add-Type -AssemblyName System.Drawing

        function ListIndexChanged { 
            #$label2.Text = $listbox.SelectedItems.Count
            $okButton.enabled = $True
        }

        $form = New-Object System.Windows.Forms.Form
        $form.Text = 'List all clients'
        $form.Size = New-Object System.Drawing.Size(350,400)
        $form.StartPosition = 'CenterScreen'
        $Form.Opacity = 1.0
        $Form.TopMost = $false
        $Form.ShowIcon = $true                                              # Enable icon (upper left corner) $ true, disable icon
        #$Form.FormBorderStyle = 'Fixed3D'                                  # bloc resizing form
        
        $okButton = New-Object System.Windows.Forms.Button
        $okButton.Location = New-Object System.Drawing.Point(75,330)
        $okButton.Size = New-Object System.Drawing.Size(75,23)
        $okButton.Text = 'OK'
        $okButton.AutoSize = $true
        $okButton.Anchor = [System.Windows.Forms.AnchorStyles]::Bottom 
        $okButton.DialogResult = [System.Windows.Forms.DialogResult]::OK
        $okButton.enabled = $False
        $form.AcceptButton = $okButton

        $cancelButton = New-Object System.Windows.Forms.Button
        $cancelButton.Location = New-Object System.Drawing.Point(150,330)
        $cancelButton.Size = New-Object System.Drawing.Size(75,23)
        $cancelButton.Text = 'Cancel'
        $cancelButton.AutoSize = $true
        $cancelButton.Anchor = [System.Windows.Forms.AnchorStyles]::Bottom 
        $cancelButton.DialogResult = [System.Windows.Forms.DialogResult]::Cancel
        $form.CancelButton = $cancelButton
        
        $label = New-Object System.Windows.Forms.Label
        $label.Location = New-Object System.Drawing.Point(10,20)
        $label.Size = New-Object System.Drawing.Size(280,20)
        $label.Text = 'Select the client to run inventory:'
        $label.AutoSize = $true
        $label.Anchor = [System.Windows.Forms.AnchorStyles]::Top `
        -bor [System.Windows.Forms.AnchorStyles]::Bottom `
        -bor [System.Windows.Forms.AnchorStyles]::Left `
        -bor [System.Windows.Forms.AnchorStyles]::Right

        $label2 = New-Object System.Windows.Forms.Label
        $label2.Location = New-Object System.Drawing.Point(10,335)
        $label2.Size = New-Object System.Drawing.Size(20,20)
        $label2.Text = $count
        $label2.AutoSize = $true
        $label2.Anchor = [System.Windows.Forms.AnchorStyles]::Bottom `
        -bor [System.Windows.Forms.AnchorStyles]::Left 

        $listBox = New-Object System.Windows.Forms.ListBox
        $listBox.Location = New-Object System.Drawing.Point(10,40)
        $listBox.Size = New-Object System.Drawing.Size(310,20)
        $listBox.Height = 280
        $listBox.SelectionMode = 'One'
        $ListBox.Anchor = [System.Windows.Forms.AnchorStyles]::Top `
        -bor [System.Windows.Forms.AnchorStyles]::Bottom `
        -bor [System.Windows.Forms.AnchorStyles]::Left `
        -bor [System.Windows.Forms.AnchorStyles]::Right

        $listboxCollection =@()

        foreach($client in $clients){
            $Object = New-Object Object 
            $Object | Add-Member -type NoteProperty -Name id -Value $client.id
            $Object | Add-Member -type NoteProperty -Name name -Value $client.name
            $listboxCollection += $Object
        }
        
        # Count selected item
        $ListBox.Add_SelectedIndexChanged({ ListIndexChanged })

        #Add collection to the $listbox
        $listBox.Items.AddRange($listboxCollection)
        $listBox.ValueMember = "$listboxCollection.id"
        $listBox.DisplayMember = "$listboxCollection.name"
        
        #Add composant into Form
        $form.Controls.Add($okButton)
        $form.Controls.Add($cancelButton)
        $form.Controls.Add($listBox)
        $form.Controls.Add($label2)
        $form.Controls.Add($label)
        $form.Topmost = $true
        $result = $form.ShowDialog()
        
        if (($result -eq [System.Windows.Forms.DialogResult]::OK) -and $listbox.SelectedItems.Count -gt 0)
        {
            Write-Host "------------------- Client sélectionné -------------------" -BackgroundColor "White" -ForegroundColor "DarkCyan"
            $ItemsName = $listBox.SelectedItems.name
            $global:ItemsID = $listBox.SelectedItems.id
            $global:clientId = $ItemsID
            Write-Host "Client name selected :$ItemsName (ID = $clientId)" -ForegroundColor Green
            
            # RUN All requests
            Write-Host "--> Request Plannings list" -ForegroundColor Blue
            get_plannings
            Write-Host "--> Request Alerts list" -ForegroundColor Blue
            $global:Alerts = get_alerts
            Write-Host "--> Request Users list" -ForegroundColor Blue
            $global:Users = get_users
            Write-Host "--> Request Applications list" -ForegroundColor Blue
            $global:applications = get_applications
            Write-Host "--> Request Workspaces list" -ForegroundColor Blue
            $global:workspaces = get_workspaces
            Write-Host "--> Request Zones list" -ForegroundColor Blue
            $global:zones = get_zones
            Write-Host "--> Request Sites list" -ForegroundColor Blue
            $global:sites_list = get_sites
            Write-Host "--> Request Scripts list" -ForegroundColor Blue
            $global:Scripts = get_scripts
            Write-Host "--> Request Actions list" -ForegroundColor Blue
            $global:actions = get_actions
            Write-Host "--> Request SSO Provider list" -ForegroundColor Blue
            $global:privider = get_SSO_provider
            Write-Host "--> Request RUM Tracker" -ForegroundColor Blue
            $global:trackers = get_rum_tracker
            Write-Host "--> Request RUM Metrics" -ForegroundColor Blue
            $global:RUM_metrics = get_RUM_metrics
            Write-Host "--> Request Publish reports" -ForegroundColor Blue
            $global:PublishReports = get_Publish_reports
            Write-Host "--> Request Share reports" -ForegroundColor Blue
            $global:ShareReports = get_Share_reports
            Write-Host "--> Request Shared data" -ForegroundColor Blue
            $global:shared_data = get_shared_data
            
            # RUN INVENTORY FOR CLIENT ------------------------------------------------------------       
            Inventory -clientName $ItemsName -clientID $ItemsID  

            Write-Host ("END Client inventory ["+$ItemsName+"]: " + $Result_OK) -BackgroundColor Green
            [System.Windows.Forms.MessageBox]::Show("Client inventory [$ItemsName] : $Result_OK","Resultat",[System.Windows.Forms.MessageBoxButtons]::OKCancel,[System.Windows.Forms.MessageBoxIcon]::Information)

        }else{
            Write-Host "Aucun client sélectionné" -ForegroundColor Red
            [System.Windows.Forms.MessageBox]::Show(`
                "------------------------------------`n`r Aucun client sélectionné`n`r------------------------------------`n`r",`
                "Resultat",[System.Windows.Forms.MessageBoxButtons]::OKCancel,[System.Windows.Forms.MessageBoxIcon]::Warning)
        }
    }
    catch{
        Write-Host "-------------------------------------------------------------" -ForegroundColor red
        Write-Host "Erreur ...." -BackgroundColor Red
        Write-Host $Error.exception.Message[0]
        Write-Host $Error[0]
        Write-host $error[0].ScriptStackTrace
        Write-Host "-------------------------------------------------------------" -ForegroundColor red

        Error_popup($Error[0])
    } 
}

function Inventory($clientName,$clientID){
    if(($clientName -ne "") -and ($clientID -ne "")){
        try{
            # Creation du ficheir XLS ---------------------------------------------------
            [String]$global:NewFile = 'EKARA_Inventaire_'+$clientName+"_"+$date + '.xlsx'          # Nom du Nouveau fichier XLS 
            Write-Host "Create file ["$NewFile"]"  
            Write-Host "Folder ["$path"]"         
            $ExcelPath = "$Path\$NewFile"                                                # Nom et Chemin du fichier
            $objExcel = new-object -comobject excel.application                          # Création du fichier XLS
            $objExcel.Visible = $False                                                   # Affiche Excel

            $FirstRowColor = 20                                                          # Cellule Couleur 1
            $SecondtRowColor = 2                                                         # Cellule Couleur 2
            $color = $FirstRowColor
            $VAlignTop = -4160                                                           # Cellule Vertical Aligne Haut
            $VAlignBottom = -4107                                                        # Cellule Vertical Aligne bas
            $VAlignCenter = -4108                                                        # Cellule Vertical Aligne Center
            $HAlignCenter = -4108                                                        # Cellule Horizontal Aligne Center
            $HAlignLeft = -4131                                                          # Cellule Horizontal Aligne Gauche
            $HAlignRight = -4152                                                         # Cellule Horizontal Aligne Droite

            # Microsoft.Office.Interop.Excel.XlBorderWeight
            $xlHairline = 1
            $xlThin = 2
            $xlThick = 4
            $xlMedium = -4138

            # Microsoft.Office.Interop.Excel.XlBordersIndex
            $xlDiagonalDown = 5
            $xlDiagonalUp = 6
            $xlEdgeLeft = 7
            $xlEdgeTop = 8
            $xlEdgeBottom = 9
            $xlEdgeRight = 10
            $xlInsideVertical = 11
            $xlInsideHorizontal = 12

            # Microsoft.Office.Interop.Excel.XlLineStyle
            $xlContinuous = 1
            $xlDashDot = 4
            $xlDashDotDot = 5
            $xlSlantDashDot = 13
            $xlLineStyleNone = -4142
            $xlDouble = -4119
            $xlDot = -4118
            $xlDash = -4115

            $xlAutomatic = -4105
            $xlBottom = -4107
            $xlCenter = -4108
            $xlContext = -5002
            $xlNone = -4142


            # color index
            $xlColorIndexBlue = 23 # <-- depends on default palette

            #region Rum_Inventaire _______________________________________________________________________________________________________
            Write-Host "------------------------------------------------------------------------" -ForegroundColor White
            Write-Host "--> Start RUM inventory for Client Name [$clientName], Client id [$clientID]" -ForegroundColor "Green"
            $row = 1
            $Column = 1
            $workbook = $objExcel.Workbooks.add()                                   # Ajout d'une feuille au fichier XLS
            $finalWorkSheet = $workbook.WorkSheets.item(1)                          # Sélection du nouvel onglet                    
            $finalWorkSheet.Name = "RUM"                                            # Nom de l'onglet
            
            # Création des entête de colonnes
            if($Debug -eq $true){Write-Host "Create header" -ForegroundColor Yellow}                       
            $finalWorkSheet.Cells.Item($row,$Column) = "Tracker Name" 
            $Column++
            $finalWorkSheet.Cells.Item($row,$Column) = "Application" 
            $Column++
            $finalWorkSheet.Cells.Item($row,$Column) = "ratio"
            $Column++
            $finalWorkSheet.Cells.Item($row,$Column) = "fav Metric"
            $Column++
            $finalWorkSheet.Cells.Item($row,$Column) = "PC"
            $Column++
            $finalWorkSheet.Cells.Item($row,$Column) = "Phone"
            $Column++
            $finalWorkSheet.Cells.Item($row,$Column) = "Tablet"
            $Column++
            $finalWorkSheet.Cells.Item($row,$Column) = "Page Groupe"
            $Column++
            $finalWorkSheet.Cells.Item($row,$Column) = "DIM business"
            $Column++
            $finalWorkSheet.Cells.Item($row,$Column) = "DIM custom"
            $Column++
            $finalWorkSheet.Cells.Item($row,$Column) = "DIM infra"
            $Column++
            $finalWorkSheet.Cells.Item($row,$Column) = "DIM version"
            
            # Format first line
            $range = $finalWorkSheet.UsedRange                                        # Sélectionne les cellules utilisées
            $range.Interior.ColorIndex = 14                                           # Change la couleur des cellule
            $range.Font.Name = "Arial"                                                # Fixe la police
            $range.Font.Size = 18                                                     # Fixe la taille de la police
            $range.Font.Bold = $True                                                  # Met en carractères gras
            $range.Font.ColorIndex = 2                                                # Fixe la couleur de la police
            $range.AutoFilter() | Out-Null                                            # Ajout du filtre automatique
            $range.EntireColumn.AutoFit() | Out-Null                                  # Auto ajustement de la taille de la colonne
            $range.VerticalAlignment = -4160                                          # Aligne vertivalement le contenu des cellules

            if($trackers.trackerId.count -gt 0){
                #========================== RUM =============================
                # Call WS : Gets all the Rum tracker of the current tenant
                Write-Host "--> Nb Trackers : " $trackers.trackerId.count -ForegroundColor Blue
                $Result1 = @()
                $global:Analyze_Rum = @()

                Foreach ($tracker in $trackers)
                {
                    Write-Host "." -NoNewline -ForegroundColor Blue

                    $favMetric = $RUM_metrics | Where-Object {$_.id -eq $tracker.favMetric}
                    
                    # Search détail for RUM tracker 
                    $RUM_Overview = get_RUM_overview -trackerID $tracker.trackerId  -metricID $tracker.favMetric
                    # Search pageGroup for RUM tracker 
                    $RUM_pageGroup = get_RUM_pageGroup($tracker.trackerId)
                    # Search business_Dimension for RUM tracker
                    $RUM_business_Dimension = get_RUM_business_Dimension($tracker.trackerId)
                    # Search custom_Dimension for RUM tracker
                    $RUM_custom_Dimension = get_RUM_custom_Dimension($tracker.trackerId)
                    # Search infra_Dimension for RUM tracker
                    $RUM_infra_Dimension = get_RUM_infra_Dimension($tracker.trackerId)
                    # Search version_Dimension for RUM tracker
                    $RUM_version_Dimension = get_RUM_version_Dimension($tracker.trackerId)

                    if($Debug -eq $true){Write-Host "Tracker Name : " $tracker.trackerName -ForegroundColor Yellow}
                    if($Debug -eq $true){Write-Host "Tracker Application Name : " $tracker.applicationName -ForegroundColor Yellow}
                    if($Debug -eq $true){Write-Host "Tracker ratio : " $tracker.ratio "%" -ForegroundColor Yellow} 
                    if($Debug -eq $true){Write-Host "Tracker tracker ID : " $tracker.trackerId -ForegroundColor Yellow}
                    if($Debug -eq $true){Write-Host "Tracker fav Metric : " $tracker.favMetric -ForegroundColor Yellow}
                    if($Debug -eq $true){Write-Host "Tracker fav Metric name : " $favMetric.label -ForegroundColor Yellow}
                    $PC = $RUM_Overview | Where-object {$_.device -eq "PC"}
                    if($Debug -eq $true){Write-Host "Tracker device PC count : " $PC.count -ForegroundColor Yellow}
                    if($Debug -eq $true){Write-Host "Tracker device PC sum : " $PC.sum -ForegroundColor Yellow}
                    if($Debug -eq $true){Write-Host ("Tracker device PC average : " + [math]::Round($PC.average)) -ForegroundColor Yellow}
                    $phone = $RUM_Overview | Where-object {$_.device -eq "phone"}
                    if($Debug -eq $true){Write-Host "Tracker device phone count : " $phone.count -ForegroundColor Yellow}
                    if($Debug -eq $true){Write-Host "Tracker device phone sum : " $phone.sum -ForegroundColor Yellow}
                    if($Debug -eq $true){Write-Host ("Tracker device phone average : " + [math]::Round($phone.average)) -ForegroundColor Yellow}
                    $tablet = $RUM_Overview | Where-object {$_.device -eq "tablet"}
                    if($Debug -eq $true){Write-Host "Tracker device tablet count : " $tablet.count -ForegroundColor Yellow}
                    if($Debug -eq $true){Write-Host "Tracker device tablet sum : " $tablet.sum -ForegroundColor Yellow}
                    if($Debug -eq $true){Write-Host ("Tracker device tablet average : " + [math]::Round($tablet.average)) -ForegroundColor Yellow}
                    if($Debug -eq $true){Write-Host "NB Tracker pageGroup : " $RUM_pageGroup.count -ForegroundColor Yellow}
                    if($Debug -eq $true){Write-Host "NB Tracker business_Dimension : " $RUM_business_Dimension.count -ForegroundColor Yellow}
                    if($Debug -eq $true){Write-Host "NB Tracker custom_Dimension : " $RUM_custom_Dimension.count -ForegroundColor Yellow}
                    if($Debug -eq $true){Write-Host "NB Tracker infra_Dimension : " $RUM_infra_Dimension.count -ForegroundColor Yellow}
                    if($Debug -eq $true){Write-Host "NB Tracker version_Dimension : " $RUM_version_Dimension.count -ForegroundColor Yellow}

                    $Result1 += New-Object -TypeName psobject -Property @{Name = [string]$tracker.trackerName;`
                                                                        Application = [string]$tracker.applicationName;`
                                                                        Ratio = [string]$tracker.ratio+"%";`
                                                                        fav_Metric = [string]$tracker.favMetric;`
                                                                        fav_Metric_name = [string]$favMetric.label;`
                                                                        PC_count = [int]$PC.count;`
                                                                        Phone_count = [int]$phone.count;`
                                                                        Tablet_count = [int]$tablet.count;`
                                                                        PageGroup_count = [int]$RUM_pageGroup.count;`
                                                                        DIM_Business = [int]$RUM_business_Dimension.count;`
                                                                        DIM_Custom = [int]$RUM_custom_Dimension.count;`
                                                                        DIM_Infra = [int]$RUM_infra_Dimension.count;`
                                                                        DIM_Version = [int]$RUM_version_Dimension.count;`
                                                                        } | Select-Object Name,Application,Ratio,fav_Metric,fav_Metric_name,PC_count,Phone_count,Tablet_count,PageGroup_count,DIM_Business,DIM_Custom,DIM_Infra,DIM_Version
                    
                    $Column = 1
                    $row++
                    $Range = $finalWorkSheet.Range("A"+$row).Select()
                    $finalWorkSheet.Cells.Item($row,$Column) = [string]$tracker.trackerName                               # Add data
                    $Column++                                                                                             # next column
                    $finalWorkSheet.Cells.Item($row,$Column) = [string]$tracker.applicationName                           # Add data    
                    $Column++                                                                                             # next column
                    $finalWorkSheet.Cells.Item($row,$Column) = [string]$tracker.ratio+" %"                                # Add data                                                
                    $Column++                                                                                             # next column
                    $finalWorkSheet.Cells.Item($row,$Column) = [string]$favMetric.label                                   # Add data
                    $Column++                                                                                             # next column
                    $finalWorkSheet.Cells.Item($row,$Column) = [int]$PC.count                                             # Add data
                    $Column++                                                                                             # next column
                    $finalWorkSheet.Cells.Item($row,$Column) = [int]$phone.count                                          # Add data
                    $Column++                                                                                             # next column
                    $finalWorkSheet.Cells.Item($row,$Column) = [int]$tablet.count                                         # Add data
                    $Column++                                                                                             # next column
                    $finalWorkSheet.Cells.Item($row,$Column) = [int]$RUM_pageGroup.count                                  # Add data
                    $Column++                                                                                             # next column
                    $finalWorkSheet.Cells.Item($row,$Column) = [int]$RUM_business_Dimension.count                         # Add data
                    $Column++                                                                                             # next column
                    $finalWorkSheet.Cells.Item($row,$Column) = [int]$RUM_custom_Dimension.count                           # Add data
                    $Column++                                                                                             # next column
                    $finalWorkSheet.Cells.Item($row,$Column) = [int]$RUM_infra_Dimension.count                            # Add data
                    $Column++                                                                                             # next column
                    $finalWorkSheet.Cells.Item($row,$Column) = [int]$RUM_version_Dimension.count                          # Add data
                    
                    # Mise en page du document
                    $range = $finalWorkSheet.Range(("A{0}" -f $row),("L{0}" -f $row))                                     # Selectionne la ligne en cours
                    $range.Select() | Out-Null                                                                            # Select range
                    $range.Interior.ColorIndex = $color                                                                   # Change la couleur des cellule
                    $range.Font.Size = 12                                                                                 # Fixe la taille de la police
                    $range.EntireColumn.AutoFit() | Out-Null                                                              # Auto ajustement de la taille de la colonne
                    $range.VerticalAlignment = -4160                                                                      # Aligne vertivalement le contenu des cellules
                                        
                    if($color -EQ $FirstRowColor){$color = $SecondtRowColor} ELSE {$color = $FirstRowColor}               # Change color line

                    # ANALYZE _____________________________________________________________________________________
                    if($RUM_pageGroup.count -lt 1){$global:Analyze_Rum += New-Object -TypeName psobject -Property @{Comment = [string]"Les pageGroup ne sont pas configurees pour le tracker ["+$tracker.trackerName+"].";`
                                                                                                                                } | Select-Object Comment}
                    
                    if(($PC.count + $phone.count + $tablet.count) -lt 1){$global:Analyze_Rum += New-Object -TypeName psobject -Property @{Comment = [string]"Aucune metrique pour le tracker ["+$tracker.trackerName+"] depuis $global:Rum_period jours.";`
                                                                                                                            } | Select-Object Comment}
                    # _____________________________________________________________________________________________
                }

                # Show Formated results
                if($Debug -eq $true){$Result1 | Out-GridView -Title "RUM Inventory"}                                      # Display results into GridView

            }else{
                Write-Host "--> Nb Trackers : " $trackers.count -ForegroundColor Blue
                $Column = 1
                $row++
                $Range = $finalWorkSheet.Range("A"+$row).Select()
                $finalWorkSheet.Cells.Item($row,$Column) = [string]"Aucun tracker"                                        # Add data
            }
            
            $Range = $finalWorkSheet.Range("A1").Select()
            $finalWorkSheet.application.activewindow.splitcolumn = 1                                                  # Select Second column for freeze the shutters
            $finalWorkSheet.Application.ActiveWindow.SplitRow = 1                                                     # Select First row for freeze the shutters
            $finalWorkSheet.Application.ActiveWindow.FreezePanes = $true                                              # Freeze the shutters
            Write-Host ""
            Write-Host "--> END RUM inventory for Client Name [$clientName], Client id [$clientID]" -ForegroundColor "Green"
            Write-Host "------------------------------------------------------------------------" -ForegroundColor White

            #endregion ___________________________________________________________________________________________________________________

            #region Scenario_Inventaire __________________________________________________________________________________________________
            # ADD nouvelle feuille -----------------------------------------------------------------
            Write-Host "------------------------------------------------------------------------" -ForegroundColor White
            Write-Host "--> Start SCENARIOS inventory for Client Name [$clientName], Client id [$clientID]" -ForegroundColor "Green"
            $NewWorkSheet = $workbook.Worksheets.Add()                                                  # Création d'un nouvel onglet  
            $finalWorkSheet = $workbook.Worksheets.Item(1)                                              # Sélection du nouvel onglet                    
            $finalWorkSheet.Name = "Scenarios"                                                          # Nom de l'onglet

            # Formatage des cellules ---------------------------------------------------------------------------
            $row = 1
            $Column = 1
            $color = $FirstRowColor

            # Création des entête de colonnes
            if($Debug -eq $true){Write-Host "Create header" -ForegroundColor Yellow}                       
            $finalWorkSheet.Cells.Item($row,$Column) = "Nom" 
            $Column++
            $finalWorkSheet.Cells.Item($row,$Column) = "Type" 
            $Column++
            $finalWorkSheet.Cells.Item($row,$Column) = "Status"
            $Column++ 
            $finalWorkSheet.Cells.Item($row,$Column) = "Zone"
            $Column++ 
            $finalWorkSheet.Cells.Item($row,$Column) = "Site"
            $Column++ 
            $finalWorkSheet.Cells.Item($row,$Column) = "Retry"
            $Column++ 
            $finalWorkSheet.Cells.Item($row,$Column) = "Application"
            $Column++ 
            $finalWorkSheet.Cells.Item($row,$Column) = "Parcours_Name"
            $Column++
            $finalWorkSheet.Cells.Item($row,$Column) = "Version"
            $Column++
            $finalWorkSheet.Cells.Item($row,$Column) = "Etape"
            $Column++
            $finalWorkSheet.Cells.Item($row,$Column) = "Planning"
            $Column++ 
            $finalWorkSheet.Cells.Item($row,$Column) = "Freq."
            $Column++ 
            $finalWorkSheet.Cells.Item($row,$Column) = "Time_out"
            $Column++ 
            $finalWorkSheet.Cells.Item($row,$Column) = "Parameters"
            $Column++ 
            $finalWorkSheet.Cells.Item($row,$Column) = "SLA_Perf"
            $Column++ 
            $finalWorkSheet.Cells.Item($row,$Column) = "SLA_Dispo"
            $Column++ 
            $finalWorkSheet.Cells.Item($row,$Column) = "Alerte"
            
            #$Range = $finalWorkSheet.Range("A"+$row).Select()

            # Format first line
            $range = $finalWorkSheet.UsedRange                                        # Sélectionne les cellules utilisées
            $range.Interior.ColorIndex = 14                                           # Change la couleur des cellule
            $range.Font.Name = "Arial"                                                # Fixe la police
            $range.Font.Size = 18                                                     # Fixe la taille de la police
            $range.Font.Bold = $True                                                  # Met en carractères gras
            $range.Font.ColorIndex = 2                                                # Fixe la couleur de la police
            $range.AutoFilter() | Out-Null                                            # Ajout du filtre automatique
            $range.EntireColumn.AutoFit() | Out-Null                                  # Auto ajustement de la taille de la colonne
            $range.VerticalAlignment = -4160                                          # Aligne vertivalement le contenu des cellules
    
            #========================== scenarios =============================
            # Call WS : Gets all the scenarios of the current tenant
            $uri ="$API/adm-api/scenarios?clientId=$clientId"
            $scenarios = Invoke-RestMethod -Uri $uri -Method POST -Headers $headers
            Write-Host "--> Nb scenarios : " $scenarios.count -ForegroundColor Blue
            $Result1 = @()
            $global:Analyze_scenarios = @()
		    $ListAlert = ""
            $scenario_inactif = 0
            $scenario_actif = 0
            $NB_scenario_retry = 0
            $GT_10_Step = 0
            $NB_scenario_sans_alert = 0
            $NB_scenario_version = 0

            Foreach ($scenario in $scenarios)
            {
                Write-Host "." -NoNewline -ForegroundColor Blue
                if($Debug -eq $true){Write-Host "Scenario Name : " $scenario.name -ForegroundColor Yellow}  
                #if($scenario.active -eq 0){Write-Host -NoNewline "Scenarios active : ";Write-Host "Inactive " -BackgroundColor Red;$Status="Inactive";$scenario_inactif++}else{Write-Host -NoNewline "Scenarios active : ";Write-Host "Active" -BackgroundColor Green;$Status="Active";$scenario_actif++}
                if($scenario.active -eq 0){$Status="Inactif";$scenario_inactif++}else{$Status="Actif";$scenario_actif++}
                if($scenario.Retry -eq 0){$Retry="Inactif"}else{$Retry="Actif";$NB_scenario_retry++}

                if($Debug -eq $true){Write-Host "Scenario zoneName : " $scenario.zoneName -ForegroundColor Yellow}
                if($Debug -eq $true){Write-Host "Scenario applicationName : " $scenario.applicationName -ForegroundColor Yellow}
                if($Debug -eq $true){Write-Host "Scenario startDate : " $scenario.startDate -ForegroundColor Yellow} 
                if($Debug -eq $true){Write-Host "Scenario scriptName : " $scenario.scriptName -ForegroundColor Yellow}
                if($Debug -eq $true){Write-Host "Scenario Retry : " $Retry -ForegroundColor Yellow}
			    if($Debug -eq $true){Write-Host "Parcours scriptID : " $scenario.scriptId -ForegroundColor Yellow}
                if($Debug -eq $true){Write-Host "Parcours versionID : " $scenario.scriptVersion -ForegroundColor Yellow}                

                #Search Parcours info -------------------------------
                $Parcours = get_script($scenario.scriptId)
                $Parcours_detail = $Parcours.message.scriptversions | Where-Object { $_.versionId -eq $scenario.scriptVersion}
                if($Debug -eq $true){Write-Host "Parcours version num : " $Parcours_detail.version -ForegroundColor Yellow}

                # Verifier si le scénario utilise la dernière version du parcours
                $scenario_detail_version = ""
                if($Parcours.message.scriptversions.version[0] -ne $Parcours_detail.version){
                    $NB_scenario_version++                  # Compte le nobre de scénario qui n'utilisent pas la dernière version de parcours pour ANALYZE
                    $scenario_detail_version = $detail_version + "Scénario version = " +$Parcours_detail.version + "`r`nLast parcours version = "+ $Parcours.message.scriptversions.version[0]
                    if($Debug -eq $true){Write-Host ("Scénario = " +$Parcours_detail.version + " / parcours = "+ $Parcours.message.scriptversions.version[0]) -ForegroundColor Yellow}
                }

                #Search Nb step for parcours version
                $uri ="$API/script-api/script/"+$scenario.scriptId+"/version/"+$Parcours_detail.versionId+"/steps?clientId=$clientId"
                $script_info = Invoke-RestMethod -Uri $uri -Method GET -Headers $headers 
                if($Debug -eq $true){Write-Host "NB Parcours Step : " $script_info.message.steps.count -ForegroundColor Yellow}
                $script_NB_Step = $script_info.message.steps.count

                # ANALYZE _____________________________________________________________________________________
                if(($scenario.active -ne 0) -and ($script_NB_Step -gt 10)){$GT_10_Step ++}  # Compte si le scénario est actif et a plus de 10 étape pour ANALYZE
                # _____________________________________________________________________________________________

			    $plugins = $scenario.plugins
                $parameters = $scenario.parameters

                if($Debug -eq $true){Write-Host "Scenario plugins : " $plugins.Name[1] -ForegroundColor Yellow}
                if($Debug -eq $true){Write-Host "Scenario parameters : " (($parameters.Name.count)) -ForegroundColor Yellow}
                if($Debug -eq $true){Write-Host ("Parcours SLA Perf : MAX : " + $scenario.slaThreshold.performance.max + " `n`rMIN : " + $scenario.slaThreshold.performance.min) -ForegroundColor Yellow}
                if($Debug -eq $true){Write-Host ("Parcours SLA Dispo : MAX : " + $scenario.slaThreshold.availability.max + " `n`rMIN : " + $scenario.slaThreshold.availability.min) -ForegroundColor Yellow}

			    # Search Planning Name by ID -------------------------------
			    $Planning = Search_planning($scenario.planningId)
                if($Debug -eq $true){Write-Host "Scenario planning : " $Planning -ForegroundColor Yellow}
                $Freq = $Plannings | Where-Object {$_.id -eq $scenario.planningId}

                $regex = [regex] '(?im)X-EKA-PERIOD:PT(.*M)'                      # Règle de l'expression regulière pour récupérer les fréquences de planning
                $ListFreq = $regex.Matches($Freq.planning).value 
                $regex = [regex] '(?<=EKA-PERIOD:PT)(.*?)(?=M)'                   # Règle de l'expression regulière pour récupérer les fréquences de planning
                $ListFreq = $regex.Matches($ListFreq)
                #$ListFreq.value
                $Frequence = ""
                foreach ($value in $ListFreq.value){
                    $Frequence += "["+$value + "mn] " 
                }

                if($Debug -eq $true){Write-Host "Scenario Freq. : " $ListFreq.value -ForegroundColor Yellow}

                Foreach($parameter in $parameters){
                    #If($parameter.Name -eq "timeout"){Write-Host "Timeout Global : " $parameter.Value; $timeout = $parameter.Value}
                    If($parameter.Name -eq "timeout"){$timeout = $parameter.Value}
                }
			
			    #Search Alert Name by ID ----------------------------------
			    $ListAlert = ""
			    if($Debug -eq $true){Write-Host "alert count : " $scenario.alerts.id.count -ForegroundColor Yellow}
                if($scenario.alerts.id.count -gt 0){
                    Foreach($alert_ID in $scenario.alerts.id){
                        if($Debug -eq $true){Write-Host ("alert ID : " + $alert_ID) -ForegroundColor Yellow}
                        if(($alert_ID -notlike "926") -and ($alert_ID -notlike "") -and ($alert_ID -notlike $Null)){
                            $alert_name = Search_alert($alert_ID)
                            $ListAlert += $alert_name       
                        }
                    }
                    if($Debug -eq $true){Write-Host "Scenario Alerte : " $ListAlert -ForegroundColor Yellow}
                }else{
                    if($Debug -eq $true){Write-Host ("Aucune Alerte") -ForegroundColor Yellow}
                    $ListAlert = "Aucune Alerte"
                    if($scenario.active -ne 0){$NB_scenario_sans_alert++}          # Compte le nombre de scénario actif sans alerte pour ANALYZE
                }

                #Search Site Name  ----------------------------------
                if($Debug -eq $true){Write-Host "Site : " $scenario.sites -ForegroundColor Yellow}
                $site_name = ""
                Foreach($site in $scenario.sites.name){
                    if($Debug -eq $true){Write-Host "Site : " $site -ForegroundColor Yellow}
                    $site_name += $site + " `n`r"
                }


			    $Result1 += New-Object -TypeName psobject -Property @{Scenario_Name = [string]$scenario.name;`
                                                                     Type = [string]$plugins.Name[1];`
                                                                     Status = [string]$Status;`
                                                                     Zone = [string]$scenario.zoneName;`
																     Site = [string]$site_name;`
                                                                     Retry = [string]$Retry;`
                                                                     Application = [string]$scenario.applicationName;`
																     Parcours_Name = [string]$scenario.scriptName;`
                                                                     Version = $Parcours_detail.version;`
                                                                     Etape = $script_NB_Step;`
																     Planning = [string]$Planning;`
                                                                     Freq = [string]$Frequence;`
																     Time_out = $timeout;`
																     Parameters = ($parameters.Name).count;`
                                                                     SLA_Perf = [string]"MAX : " + $scenario.slaThreshold.performance.max + " `n`rMIN : " + $scenario.slaThreshold.performance.min;`
                                                                     SLA_Dispo = [string]"MAX : " + $scenario.slaThreshold.availability.max + " `n`rMIN : " + $scenario.slaThreshold.availability.min;`
                                                                     Alerte = [string]$ListAlert;`
                                                                     } | Select-Object Scenario_Name,Type,Status,Zone,Site,Retry,Application,Parcours_Name,Version,Nb_Step,Planning,Freq,Time_out,Parameters,SLA_Perf,SLA_Dispo,Alerte
            
                $Column = 1
                $row++
                $Range = $finalWorkSheet.Range("A"+$row).Select()
                $finalWorkSheet.Cells.Item($row,$Column) = [string]$scenario.name                       # Add data
                $Column++                                                                               # next column
                $finalWorkSheet.Cells.Item($row,$Column) = [string]$plugins.Name[1]                     # Add data
                $Column++                                                                               # next column
                $finalWorkSheet.Cells.Item($row,$Column) = [string]$Status                              # Add data
                $Column++                                                                               # next column
                $finalWorkSheet.Cells.Item($row,$Column) = [string]$scenario.zoneName                   # Add data
                $Column++                                                                               # next column
                $finalWorkSheet.Cells.Item($row,$Column) = [string]$site_name                           # Add data
                $Column++                                                                               # next column
                $finalWorkSheet.Cells.Item($row,$Column) = [string]$Retry                               # Add data
                $Column++                                                                               # next column
                $finalWorkSheet.Cells.Item($row,$Column) = [string]$scenario.applicationName            # Add data
                $Column++                                                                               # next column
                $finalWorkSheet.Cells.Item($row,$Column) = [string]$scenario.scriptName                 # Add data
                $Column++                                                                               # next column
                $finalWorkSheet.Cells.Item($row,$Column) = $Parcours_detail.version                     # Add data
                if($scenario_detail_version -ne ""){
                    $finalWorkSheet.Cells.Item($row,$Column).AddComment(""+$scenario_detail_version+"")                      # Add data into comment
                    $finalWorkSheet.Cells.Item($row,$Column).Comment.Shape.TextFrame.Characters().Font.Size = 8              # Format comment
                    $finalWorkSheet.Cells.Item($row,$Column).Comment.Shape.TextFrame.Characters().Font.Bold = $False         # Format comment
                }
                
                $Column++                                                                               # next column
                $finalWorkSheet.Cells.Item($row,$Column) = $script_NB_Step                              # Add data
                $Column++                                                                               # next column
                $finalWorkSheet.Cells.Item($row,$Column) = [string]$Planning                            # Add data
                $Column++                                                                               # next column
                $finalWorkSheet.Cells.Item($row,$Column) = [string]$Frequence                           # Add data                                                     
                $Column++                                                                               # next column
                $finalWorkSheet.Cells.Item($row,$Column) = $timeout                                     # Add data
                $Column++                                                                               # next column
                $finalWorkSheet.Cells.Item($row,$Column) = ($parameters.Name).count                     # Add 
                $Column++                                                                               # next column
                $finalWorkSheet.Cells.Item($row,$Column) = [string]"MAX : " + $scenario.slaThreshold.performance.max + " `n`rMIN : " + $scenario.slaThreshold.performance.min                     # Add data
                $Column++                                                                               # next column
                $finalWorkSheet.Cells.Item($row,$Column) = [string]"MAX : " + $scenario.slaThreshold.availability.max + " `n`rMIN : " + $scenario.slaThreshold.availability.min                     # Add data
                $Column++                                                                               # next column
                $finalWorkSheet.Cells.Item($row,$Column) = [string]$ListAlert                           # Add data

                # Mise en page du document
                $range = $finalWorkSheet.Range(("A{0}" -f $row),("Q{0}" -f $row))                                     # Selectionne la ligne en cours
                $range.Select() | Out-Null                                                                            # Select range
                $range.Interior.ColorIndex = $color                                                                   # Change la couleur des cellule
                $range.Font.Size = 12                                                                                 # Fixe la taille de la police
                $range.EntireColumn.AutoFit() | Out-Null                                                              # Auto ajustement de la taille de la colonne
                $range.VerticalAlignment = -4160                                                                      # Aligne vertivalement le contenu des cellules
                                    
                if($color -EQ $FirstRowColor){$color = $SecondtRowColor} ELSE {$color = $FirstRowColor}               # Change color line

            }
            
            # ANALYZE _____________________________________________________________________________________
            $scenario_desactive = ($scenarios.active | Where-Object {$_ -eq 0}).count
            if($scenario_desactive -gt 0){$global:Analyze_scenarios += New-Object -TypeName psobject -Property @{Comment = [string]"Il y a $scenario_desactive scenario(s) desactive(s).";`
                                                                                                            } | Select-Object Comment}

            if($GT_10_Step -gt 0){$global:Analyze_scenarios += New-Object -TypeName psobject -Property @{Comment = [string]"Il y a $GT_10_Step scenario(s) actif(s) avec plus de 10 etapes.";`
                                                                                                            } | Select-Object Comment}

            if($NB_scenario_sans_alert -gt 0){$global:Analyze_scenarios += New-Object -TypeName psobject -Property @{Comment = [string]"Il y a $NB_scenario_sans_alert scenario(s) actif(s) sans alerte de configuree.";`
                                                                                                            } | Select-Object Comment}
            
            if($NB_scenario_version -gt 0){$global:Analyze_scenarios += New-Object -TypeName psobject -Property @{Comment = [string]"Il y a $NB_scenario_version scenario(s) actif(s) qui n'utilisent pas la dernière version de parcours.";`
                                                                                                            } | Select-Object Comment}

            # _____________________________________________________________________________________________

		    # Show Formated results
		    if($Debug -eq $true){$Result1 | Out-GridView -Title "Scenarios Inventory"}                                # Display results into GridView
        
            $Range = $finalWorkSheet.Range("A1").Select()
            $finalWorkSheet.application.activewindow.splitcolumn = 2                                                  # Select Second column for freeze the shutters
            $finalWorkSheet.Application.ActiveWindow.SplitRow = 1                                                     # Select First row for freeze the shutters
            $finalWorkSheet.Application.ActiveWindow.FreezePanes = $true                                              # Freeze the shutters
            Write-Host ""
            Write-Host "--> End SCENARIOS inventory for Client Name [$clientName], Client id [$clientID]" -ForegroundColor "Green"
            Write-Host "------------------------------------------------------------------------" -ForegroundColor White
            #endregion ___________________________________________________________________________________________________________________
            
            #region Users_Inventaire _____________________________________________________________________________________________________
            # ADD nouvelle feuille -----------------------------------------------------------------
            Write-Host "--> Start USERS inventory for Client Name [$clientName], Client id [$clientID]" -ForegroundColor "Green"
            $row = 1
            $Column = 1
            $color = $FirstRowColor
            $NewWorkSheet = $workbook.Worksheets.Add()                                                  # Création d'un nouvel onglet  
            $finalWorkSheet = $workbook.Worksheets.Item(1)                                              # Sélection du nouvel onglet
            $finalWorkSheet.Name = "Utilisateurs"                                                       # Nom de l'onglet

            $finalWorkSheet.Cells.Item($row,$Column) = "Prenom" 
            $Column++
            $finalWorkSheet.Cells.Item($row,$Column) = "Nom" 
            $Column++
            $finalWorkSheet.Cells.Item($row,$Column) = "Profil"
            $Column++ 
            $finalWorkSheet.Cells.Item($row,$Column) = "Email"
            $Column++ 
            $finalWorkSheet.Cells.Item($row,$Column) = "Phone"
            $Column++ 
            $finalWorkSheet.Cells.Item($row,$Column) = "Time_zone"
            $Column++ 
            $finalWorkSheet.Cells.Item($row,$Column) = "Language"
            $Column++ 
            $finalWorkSheet.Cells.Item($row,$Column) = "Workspace"
            $Column++ 
            $finalWorkSheet.Cells.Item($row,$Column) = "AUTHENTICATION"
            $Range = $finalWorkSheet.Range("A"+$row).Select()
            
            # Format first line
            $range = $finalWorkSheet.UsedRange                                        # Sélectionne les cellules utilisées
            $range.Interior.ColorIndex = 14                                           # Change la couleur des cellule
            $range.Font.Name = "Arial"                                                # Fixe la police
            $range.Font.Size = 18                                                     # Fixe la taille de la police
            $range.Font.Bold = $True                                                  # Met en carractères gras
            $range.Font.ColorIndex = 2                                                # Fixe la couleur de la police
            $range.AutoFilter() | Out-Null                                            # Ajout du filtre automatique
            $range.EntireColumn.AutoFit() | Out-Null                                  # Auto ajustement de la taille de la colonne
            $range.VerticalAlignment = -4160                                          # Aligne vertivalement le contenu des cellules

            #========================== Users =============================
            # Call WS : Gets all the users of the current tenant
            Write-Host "--> Nb users : " $users.count -ForegroundColor Blue
            $NB_SSO_Provider_Name = 0
            $Result1 = @()
            Foreach ($user in $Users)
            {
                Write-Host "." -NoNewline -ForegroundColor Blue 
                if($Debug -eq $true){Write-Host "User firstname : " $user.firstname -ForegroundColor Yellow}
                if($Debug -eq $true){Write-Host "User lastname : " $user.lastname -ForegroundColor Yellow}
                if($Debug -eq $true){Write-Host "User Profil : " $user.roleName -ForegroundColor Yellow}
                if($Debug -eq $true){Write-Host "User email : " $user.email -ForegroundColor Yellow}
                if($Debug -eq $true){Write-Host "User phone : " $user.phone.nationalNumber -ForegroundColor Yellow}
                if($Debug -eq $true){Write-Host "User timezone : " $user.timezone.timeZoneLabel -ForegroundColor Yellow}
                if($Debug -eq $true){Write-Host "User language : " $user.language -ForegroundColor Yellow}
                $workspace_name = ""                                           # Reinitialize variable
                Foreach($workspace in $user.workspaces){
                    if($Debug -eq $true){Write-Host "User Workspace : " $workspace.name -ForegroundColor Yellow}
                    $workspace_name += $workspace.name + " `n`r"
                }

                if($user.identityProviderId -ne $null){
                    # Check Proviser
                    #$uri ="$API/adm-api/identityproviders?linked_client=$clientId"      
                    #$identityProvider = Invoke-RestMethod -Uri $uri -Method GET -Headers $headers
                    $SSO_Provider = $SSO_Providers | Where-Object {$_.id -eq $user.identityProviderId}
                    $SSO_Provider_Name = $SSO_Provider.name
                    $NB_SSO_Provider_Name++
                }else{
                    $SSO_Provider_Name = "Standard"
                }
                if($Debug -eq $true){Write-Host "AUTHENTICATION : " $identityProviderName -ForegroundColor Yellow}

                $Result1 += New-Object -TypeName psobject -Property @{Prenom = [string]$user.firstname;`
                                                                    Nom = [string]$user.lastname;`
                                                                    Profil = [string]$user.roleName;`
                                                                    email = [string]$user.email;`
                                                                    phone = [string]$user.phone.nationalNumber;`
                                                                    timezone = [string]$user.timezone.timeZoneLabel;`
                                                                    language = [string]$user.language;`
                                                                    Workspace = [string]$workspace_name;`
                                                                    Authentication = [string]$SSO_Provider_Name;`
                                                                    } | Select-Object First_Name,Last_Name,Profil,email,phone,timezone,language,Workspace,AUTHENTICATION
                
                $Column = 1
                $row++
                $Range = $finalWorkSheet.Range("A"+$row).Select()
                $finalWorkSheet.Cells.Item($row,$Column) = [string]$user.firstname                      # Add data
                $Column++                                                                               # next column
                $finalWorkSheet.Cells.Item($row,$Column) = [string]$user.lastname                       # Add data
                $Column++                                                                               # next column
                $finalWorkSheet.Cells.Item($row,$Column) = [string]$user.roleName                       # Add data
                $Column++                                                                               # next column
                $finalWorkSheet.Cells.Item($row,$Column) = [string]$user.email                          # Add data
                $Column++                                                                               # next column
                $finalWorkSheet.Cells.Item($row,$Column) = [string]$user.phone.nationalNumber           # Add data
                $Column++                                                                               # next column
                $finalWorkSheet.Cells.Item($row,$Column) = [string]$user.timezone.timeZoneLabel         # Add data
                $Column++                                                                               # next column
                $finalWorkSheet.Cells.Item($row,$Column) = [string]$user.language                       # Add data
                $Column++                                                                               # next column
                $finalWorkSheet.Cells.Item($row,$Column) = [string]$workspace_name                      # Add data
                $Column++                                                                               # next column
                $finalWorkSheet.Cells.Item($row,$Column) = [string]$SSO_Provider_Name                   # Add data
                                                
                # Mise en page du document
                $range = $finalWorkSheet.Range(("A{0}" -f $row),("I{0}" -f $row))                                     # Selectionne la ligne en cours
                $range.Select() | Out-Null                                                                            # Select range
                $range.Interior.ColorIndex = $color                                                                   # Change la couleur des cellule
                $range.Font.Size = 12                                                                                 # Fixe la taille de la police
                $range.EntireColumn.AutoFit() | Out-Null                                                              # Auto ajustement de la taille de la colonne
                $range.VerticalAlignment = -4160                                                                      # Aligne vertivalement le contenu des cellules
                                    
                if($color -EQ $FirstRowColor){$color = $SecondtRowColor} ELSE {$color = $FirstRowColor}               # Change color line
            }

		    # Show Formated results
		    if($Debug -eq $true){$Result1 | Out-GridView -Title "Users Inventory"}                                    # Display results into GridView
        
            $Range = $finalWorkSheet.Range("A1").Select()
            $finalWorkSheet.application.activewindow.splitcolumn = 2                                                  # Select Second column for freeze the shutters
            $finalWorkSheet.Application.ActiveWindow.SplitRow = 1                                                     # Select First row for freeze the shutters
            $finalWorkSheet.Application.ActiveWindow.FreezePanes = $true                                              # Freeze the shutters
            Write-Host ""
            Write-Host "--> END USERS inventory for Client Name [$clientName], Client id [$clientID]" -ForegroundColor "Green"
            Write-Host "------------------------------------------------------------------------" -ForegroundColor White
            #endregion ___________________________________________________________________________________________________________________

            #region Applications_Inventaire ______________________________________________________________________________________________
            # ADD nouvelle feuille -----------------------------------------------------------------
            Write-Host "--> Start Applications inventory for Client Name [$clientName], Client id [$clientID]" -ForegroundColor "Green"
            $row = 1
            $Column = 1
            $color = $FirstRowColor
            $NewWorkSheet = $workbook.Worksheets.Add()                                                  # Création d'un nouvel onglet  
            $finalWorkSheet = $workbook.Worksheets.Item(1)                                              # Sélection du nouvel onglet
            $finalWorkSheet.Name = "Applications"                                                       # Nom de l'onglet

            $finalWorkSheet.Cells.Item($row,$Column) = "Nom" 
            $Column++
            $finalWorkSheet.Cells.Item($row,$Column) = "Description" 
            $Column++
            $finalWorkSheet.Cells.Item($row,$Column) = "SLA_Perf"
            $Column++ 
            $finalWorkSheet.Cells.Item($row,$Column) = "SLA_Dispo"
            $Column++ 
            $finalWorkSheet.Cells.Item($row,$Column) = "Workspaces"
            $Range = $finalWorkSheet.Range("A"+$row).Select()
            
            # Format first line
            $range = $finalWorkSheet.UsedRange                                        # Sélectionne les cellules utilisées
            $range.Interior.ColorIndex = 14                                           # Change la couleur des cellule
            $range.Font.Name = "Arial"                                                # Fixe la police
            $range.Font.Size = 18                                                     # Fixe la taille de la police
            $range.Font.Bold = $True                                                  # Met en carractères gras
            $range.Font.ColorIndex = 2                                                # Fixe la couleur de la police
            $range.AutoFilter() | Out-Null                                            # Ajout du filtre automatique
            $range.EntireColumn.AutoFit() | Out-Null                                  # Auto ajustement de la taille de la colonne
            $range.VerticalAlignment = -4160                                          # Aligne vertivalement le contenu des cellules

            #========================== Applications =============================
            # Call WS : Gets all the Applications of the current tenant
            Write-Host ("--> Nb Applications : " + [int]$applications.id.count) -ForegroundColor Blue
            $Result1 = @()
            Foreach ($application in $applications)
            {
                Write-Host "." -NoNewline -ForegroundColor Blue
                if($Debug -eq $true){Write-Host "Applications Name : " $application.name -ForegroundColor Yellow}
                if($Debug -eq $true){Write-Host "Applications Description : " $application.description -ForegroundColor Yellow}
                if($Debug -eq $true){Write-Host ("Applications SLA_Perf : MAX : " + $application.slaThresholds.performance.max + " `n`rMIN : " + $application.slaThresholds.performance.min) -ForegroundColor Yellow}
                if($Debug -eq $true){Write-Host ("Applications SLA_Dispo : MAX : " + $application.slaThresholds.availability.max + " `n`rMIN : " + $application.slaThresholds.availability.min) -ForegroundColor Yellow}
                
                $workspace_name = ""                                           # Reinitialize variable
                Foreach($workspaceID in $application.workspaces){
                    #Recherche l'ID dans la liste des Workspaces
                    $w_name = $workspaces | Where-Object {$_.id -eq $workspaceID}
                    if($Debug -eq $true){Write-Host "User Workspace : " $w_name.name}
                    $workspace_name += $w_name.name + " `n`r"
                }
                if($Debug -eq $true){Write-Host "Applications Workspaces : " $workspace_name -ForegroundColor Yellow}

                $Result1 += New-Object -TypeName psobject -Property @{Nom = [string]$application.name;`
                                                                    Description = [string]$application.description;`
                                                                    SLA_Perf = [string]"MAX : " + $application.slaThresholds.performance.max + " `n`rMIN : " + $application.slaThresholds.performance.min;`
                                                                    SLA_Dispo = [string]"MAX : " + $application.slaThresholds.availability.max + " `n`rMIN : " + $application.slaThresholds.availability.min;`
                                                                    Workspace = [string]$workspace_name;`
                                                                    } | Select-Object Name,Description,SLA_Perf,SLA_Dispo,phone,Workspace
                
                $Column = 1
                $row++
                $Range = $finalWorkSheet.Range("A"+$row).Select()
                $finalWorkSheet.Cells.Item($row,$Column) = [string]$application.name                                  # Add data
                $Column++                                                                                             # next column
                $finalWorkSheet.Cells.Item($row,$Column) = [string]$application.description                           # Add data
                $Column++                                                                                             # next column
                $finalWorkSheet.Cells.Item($row,$Column) = [string]"MAX : " + $application.slaThresholds.performance.max + " `n`rMIN : " + $application.slaThresholds.performance.min                       # Add data
                $Column++                                                                                             # next column
                $finalWorkSheet.Cells.Item($row,$Column) = [string]"MAX : " + $application.slaThresholds.availability.max + " `n`rMIN : " + $application.slaThresholds.availability.min                          # Add data
                $Column++                                                                                             # next column
                $finalWorkSheet.Cells.Item($row,$Column) = [string]$workspace_name                                    # Add data
                                                
                # Mise en page du document
                $range = $finalWorkSheet.Range(("A{0}" -f $row),("E{0}" -f $row))                                     # Selectionne la ligne en cours
                $range.Select() | Out-Null                                                                            # Select range
                $range.Interior.ColorIndex = $color                                                                   # Change la couleur des cellule
                $range.Font.Size = 12                                                                                 # Fixe la taille de la police
                $range.EntireColumn.AutoFit() | Out-Null                                                              # Auto ajustement de la taille de la colonne
                $range.VerticalAlignment = -4160                                                                      # Aligne vertivalement le contenu des cellules
                                    
                if($color -EQ $FirstRowColor){$color = $SecondtRowColor} ELSE {$color = $FirstRowColor}               # Change color line
            }

		    # Show Formated results
		    if($Debug -eq $true){$Result1 | Out-GridView -Title "Applications Inventory"}                             # Display results into GridView
        
            $Range = $finalWorkSheet.Range("A1").Select()
            $finalWorkSheet.application.activewindow.splitcolumn = 1                                                  # Select Second column for freeze the shutters
            $finalWorkSheet.Application.ActiveWindow.SplitRow = 1                                                     # Select First row for freeze the shutters
            $finalWorkSheet.Application.ActiveWindow.FreezePanes = $true                                              # Freeze the shutters
            Write-Host ""
            Write-Host "--> END Applications inventory for Client Name [$clientName], Client id [$clientID]" -ForegroundColor "Green"
            Write-Host "------------------------------------------------------------------------" -ForegroundColor White
            #endregion ___________________________________________________________________________________________________________________

            #region Workspaces_Inventaire ________________________________________________________________________________________________
            # ADD nouvelle feuille -----------------------------------------------------------------
            Write-Host "--> Start Workspaces inventory for Client Name [$clientName], Client id [$clientID]" -ForegroundColor "Green"
            $row = 1
            $Column = 1
            $color = $FirstRowColor
            $NewWorkSheet = $workbook.Worksheets.Add()                                                  # Création d'un nouvel onglet  
            $finalWorkSheet = $workbook.Worksheets.Item(1)                                              # Sélection du nouvel onglet
            $finalWorkSheet.Name = "Workspaces"                                                         # Nom de l'onglet

            $finalWorkSheet.Cells.Item($row,$Column) = "Nom" 
            $Column++
            $finalWorkSheet.Cells.Item($row,$Column) = "Description" 
            $Column++
            $finalWorkSheet.Cells.Item($row,$Column) = "Applications"
            $Column++ 
            $finalWorkSheet.Cells.Item($row,$Column) = "Utilisateurs"
            $Column++ 
            $finalWorkSheet.Cells.Item($row,$Column) = "Rapports"
            $Range = $finalWorkSheet.Range("A"+$row).Select()
            
            # Format first line
            $range = $finalWorkSheet.UsedRange                                        # Sélectionne les cellules utilisées
            $range.Interior.ColorIndex = 14                                           # Change la couleur des cellule
            $range.Font.Name = "Arial"                                                # Fixe la police
            $range.Font.Size = 18                                                     # Fixe la taille de la police
            $range.Font.Bold = $True                                                  # Met en carractères gras
            $range.Font.ColorIndex = 2                                                # Fixe la couleur de la police
            $range.AutoFilter() | Out-Null                                            # Ajout du filtre automatique
            $range.EntireColumn.AutoFit() | Out-Null                                  # Auto ajustement de la taille de la colonne
            $range.VerticalAlignment = -4160                                          # Aligne vertivalement le contenu des cellules

            #========================== Workspaces =============================
            # Call WS : Gets all the Workspaces of the current tenant
            Write-Host ("--> Nb Workspaces : " + [int]$workspaces.id.count) -ForegroundColor Blue
            $Result1 = @()

            Foreach ($workspace in $workspaces)
            {
                Write-Host "." -NoNewline -ForegroundColor Blue
                if($Debug -eq $true){Write-Host "Workspace Name : " $workspace.name -ForegroundColor Yellow}
                if($Debug -eq $true){Write-Host "Workspace Description : " $workspace.description -ForegroundColor Yellow}
                
                $workspace_applications = ""                                               # Reinitialize variable
                Foreach($applicationName in $workspace.applications.name){
                    #$app_name = $applications | Where-Object {$_.id -eq $applicationID}
                    if($Debug -eq $true){Write-Host "User Workspace : " $applicationName -ForegroundColor Yellow}
                    $workspace_applications += $applicationName + " `n`r"
                }
                if($Debug -eq $true){Write-Host "Workspace Applications : " $workspace_applications -ForegroundColor Yellow}

                $workspace_users = ""                                           # Reinitialize variable
                Foreach($userID in $workspace.users){
                    #Recherche l'ID dans la liste des Workspaces
                    $user_name = $Users | Where-Object {$_.id -eq $userID}
                    if($Debug -eq $true){Write-Host ("User Workspace : " + $user_name.firstname + " " + $user_name.lastname) -ForegroundColor Yellow}
                    $workspace_users += $user_name.firstname + " " + $user_name.lastname + " `n`r"
                }
                if($Debug -eq $true){Write-Host "Workspace Users : " $workspace_users -ForegroundColor Yellow}

                $workspace_Reports = ""
                Foreach($ReportName in $workspace.views.name){
                    #$report_name = $applications | Where-Object {$_.id -eq $applicationID}
                    if($Debug -eq $true){Write-Host "Report Workspace : " $ReportName -ForegroundColor Yellow}
                    $workspace_Reports += $ReportName + " `n`r"
                }
                if($Debug -eq $true){Write-Host "Workspace Rapports : " $workspace_Reports -ForegroundColor Yellow}

                $Result1 += New-Object -TypeName psobject -Property @{Nom = [string]$workspace.name;`
                                                                    Description = [string]$workspace.description;`
                                                                    Applications = [string]$workspace_applications;`
                                                                    Utilisateurs = [string]$workspace_users;`
                                                                    Rapports = [string]$workspace_Reports;`
                                                                    } | Select-Object Name,Description,Applications,Users,Reports
                
                $Column = 1
                $row++
                $Range = $finalWorkSheet.Range("A"+$row).Select()
                $finalWorkSheet.Cells.Item($row,$Column) = [string]$workspace.name                                      # Add data
                $Column++                                                                                               # next column
                $finalWorkSheet.Cells.Item($row,$Column) = [string]$workspace.description                               # Add data
                $Column++                                                                                               # next column
                $finalWorkSheet.Cells.Item($row,$Column) = [string]$workspace_applications                              # Add data
                $Column++                                                                                               # next column
                $finalWorkSheet.Cells.Item($row,$Column) = [string]$workspace_users                                     # Add data
                $Column++                                                                                               # next column
                $finalWorkSheet.Cells.Item($row,$Column) = [string]$workspace_Reports                                   # Add data
                                                
                # Mise en page du document
                $range = $finalWorkSheet.Range(("A{0}" -f $row),("E{0}" -f $row))                                     # Selectionne la ligne en cours
                $range.Select() | Out-Null                                                                            # Select range
                $range.Interior.ColorIndex = $color                                                                   # Change la couleur des cellule
                $range.Font.Size = 12                                                                                 # Fixe la taille de la police
                $range.EntireColumn.AutoFit() | Out-Null                                                              # Auto ajustement de la taille de la colonne
                $range.VerticalAlignment = -4160                                                                      # Aligne vertivalement le contenu des cellules
                                    
                if($color -EQ $FirstRowColor){$color = $SecondtRowColor} ELSE {$color = $FirstRowColor}               # Change color line
            }

		    # Show Formated results
		    if($Debug -eq $true){$Result1 | Out-GridView -Title "Workspaces Inventory"}                               # Display results into GridView
        
            $Range = $finalWorkSheet.Range("A1").Select()
            $finalWorkSheet.application.activewindow.splitcolumn = 1                                                  # Select Second column for freeze the shutters
            $finalWorkSheet.Application.ActiveWindow.SplitRow = 1                                                     # Select First row for freeze the shutters
            $finalWorkSheet.Application.ActiveWindow.FreezePanes = $true                                              # Freeze the shutters
            Write-Host ""
            Write-Host "--> END Workspaces inventory for Client Name [$clientName], Client id [$clientID]" -ForegroundColor "Green"
            Write-Host "------------------------------------------------------------------------" -ForegroundColor White
            #endregion ___________________________________________________________________________________________________________________

            #region Zone_Inventaire ______________________________________________________________________________________________________
            # ADD nouvelle feuille -----------------------------------------------------------------
            Write-Host "--> Start Zones inventory for Client Name [$clientName], Client id [$clientID]" -ForegroundColor "Green"
            $row = 1
            $Column = 1
            $color = $FirstRowColor
            $global:Analyze_Zones = @()
            $NewWorkSheet = $workbook.Worksheets.Add()                                                  # Création d'un nouvel onglet  
            $finalWorkSheet = $workbook.Worksheets.Item(1)                                              # Sélection du nouvel onglet
            $finalWorkSheet.Name = "Zones"                                                              # Nom de l'onglet

            $finalWorkSheet.Cells.Item($row,$Column) = "Nom" 
            $Column++
            $finalWorkSheet.Cells.Item($row,$Column) = "Description" 
            $Column++
            $finalWorkSheet.Cells.Item($row,$Column) = "Failstatuscount"
            $Column++ 
            $finalWorkSheet.Cells.Item($row,$Column) = "Scenarios Nb"
            $Column++ 
            $finalWorkSheet.Cells.Item($row,$Column) = "Sites"
            $Column++ 
            $finalWorkSheet.Cells.Item($row,$Column) = "Type"
            $Column++ 
            $finalWorkSheet.Cells.Item($row,$Column) = "Application"

            $Range = $finalWorkSheet.Range("A"+$row).Select()
            
            # Format first line
            $range = $finalWorkSheet.UsedRange                                        # Sélectionne les cellules utilisées
            $range.Interior.ColorIndex = 14                                           # Change la couleur des cellule
            $range.Font.Name = "Arial"                                                # Fixe la police
            $range.Font.Size = 18                                                     # Fixe la taille de la police
            $range.Font.Bold = $True                                                  # Met en carractères gras
            $range.Font.ColorIndex = 2                                                # Fixe la couleur de la police
            $range.AutoFilter() | Out-Null                                            # Ajout du filtre automatique
            $range.EntireColumn.AutoFit() | Out-Null                                  # Auto ajustement de la taille de la colonne
            $range.VerticalAlignment = -4160                                          # Aligne vertivalement le contenu des cellules

            #========================== Zones =============================
            # Call WS : Gets all the Zones of the current tenant
            Write-Host "--> Nb Zones : " $zones.count -ForegroundColor Blue
            $Result1 = @()                                                     # Reinitialize variable
            $site_private_list = @()                                           # Reinitialize variable
            $site_public_list = @()                                            # Reinitialize variable
            $site_private = "NO"                                               # Reinitialize variable
            $site_public = "NO"                                                # Reinitialize variable
            

            Foreach ($zone in $zones)
            {
                Write-Host "." -NoNewline -ForegroundColor Blue
                if($Debug -eq $true){Write-Host "Zone Name : " $zone.name -ForegroundColor Yellow}
                if($Debug -eq $true){Write-Host "Zone Description : " $zone.description -ForegroundColor Yellow}
                if($Debug -eq $true){Write-Host "Zone Failstatuscount : " $zone.Failstatuscount -ForegroundColor Yellow}
                if($Debug -eq $true){Write-Host "Zone ScenariosNb : " $zone.ScenariosNb -ForegroundColor Yellow}
                
                $zone_sites = ""                                                # Reinitialize variable
                $site_type = ""                                                 # Reinitialize variable
                $zone_Application = @()                                         # Reinitialize variable

                Foreach($sites in $zone.sites){
                    if($Debug -eq $true){Write-Host "Site name : " $sites.name -ForegroundColor Yellow}
                    $zone_sites += $sites.name + " `n`r"
                
                    # Rercherche si le site est privé.
                    $Site_to_sites_list = ($sites_list | Where-object {$_.siteId -eq $sites.id})
                    if($Site_to_sites_list.isPrivate -eq $True){
                        $site_private = "YES"
                        $site_type+= "Private `n`r"
                        $site_private_list+=$Site_to_sites_list.name
                    }else{
                        $site_public = "YES"
                        $site_type+= "Public `n`r"
                        $site_public_list+=$Site_to_sites_list.name
                    }
                }
                
                # Check if Zone is Global
                if($zone.applications.name.count -gt 0){
                    foreach($name in $zone.applications.name){
                        $zone_Application+= $name+"`r`n"
                    }
                }else{
                    $zone_Application = "Global"
                }

                $unic_privat_site = $site_private_list | Select-Object -Unique                  # List unique private site name
                $unic_public_site = $site_public_list | Select-Object -Unique                   # List unique public site name
                
                $Result1 += New-Object -TypeName psobject -Property @{Nom = [string]$zone.name;`
                                                                    Description = [string]$zone.description;`
                                                                    Failstatuscount = [string]$zone.Failstatuscount;`
                                                                    Scenarios_Nb = [string]$zone.ScenariosNb;`
                                                                    Sites = [string]$zone_sites;`
                                                                    Type = [string]$site_type;`
                                                                    Applications = [string]$zone_Application;`
                                                                    } | Select-Object Name,Description,Failstatuscount,Scenarios_Nb,Sites,Type,Applications
                
                $Column = 1
                $row++
                $Range = $finalWorkSheet.Range("A"+$row).Select()
                $finalWorkSheet.Cells.Item($row,$Column) = [string]$zone.name                                           # Add data
                $Column++                                                                                               # next column
                $finalWorkSheet.Cells.Item($row,$Column) = [string]$zone.description                                    # Add data
                $Column++                                                                                               # next column
                $finalWorkSheet.Cells.Item($row,$Column) = [string]$zone.Failstatuscount                                # Add data
                $Column++                                                                                               # next column
                $finalWorkSheet.Cells.Item($row,$Column) = [string]$zone.ScenariosNb                                    # Add data
                $Column++                                                                                               # next column
                $finalWorkSheet.Cells.Item($row,$Column) = [string]$zone_sites                                          # Add data
                $Column++                                                                                               # next column
                $finalWorkSheet.Cells.Item($row,$Column) = [string]$site_type                                           # Add data
                $Column++                                                                                               # next column
                $finalWorkSheet.Cells.Item($row,$Column) = [string]$zone_Application                                    # Add data
                                                
                # Mise en page du document
                $range = $finalWorkSheet.Range(("A{0}" -f $row),("G{0}" -f $row))                                     # Selectionne la ligne en cours
                $range.Select() | Out-Null                                                                            # Select range
                $range.Interior.ColorIndex = $color                                                                   # Change la couleur des cellule
                $range.Font.Size = 12                                                                                 # Fixe la taille de la police
                $range.EntireColumn.AutoFit() | Out-Null                                                              # Auto ajustement de la taille de la colonne
                $range.VerticalAlignment = -4160                                                                      # Aligne vertivalement le contenu des cellules
                                    
                if($color -EQ $FirstRowColor){$color = $SecondtRowColor} ELSE {$color = $FirstRowColor}               # Change color line
            }

            foreach($site in $unic_privat_site){
                $unic_privat_site_list += $site+" `r`n"
            }

            foreach($site in $unic_public_site){
                $unic_public_site_list += $site+" `r`n"
            }

            # ANALYZE _____________________________________________________________________________________
            $Zone_sans_scenario = ($zones.scenariosNb | Where-object {$_ -eq 0}).count
            if($Zone_sans_scenario -gt 0){$global:Analyze_Zones += New-Object -TypeName psobject -Property @{Comment = [string]"Il y a $Zone_sans_scenario zone(s) sans aucun parcours affecte.";`
                                                                                                            } | Select-Object Comment}

            # _____________________________________________________________________________________________

		    # Show Formated results
		    if($Debug -eq $true){$Result1 | Out-GridView -Title "Zones Inventory"}                                    # Display results into GridView
        
            $Range = $finalWorkSheet.Range("A1").Select()
            $finalWorkSheet.application.activewindow.splitcolumn = 1                                                  # Select Second column for freeze the shutters
            $finalWorkSheet.Application.ActiveWindow.SplitRow = 1                                                     # Select First row for freeze the shutters
            $finalWorkSheet.Application.ActiveWindow.FreezePanes = $true                                              # Freeze the shutters
            Write-Host ""
            Write-Host "--> END Zones inventory for Client Name [$clientName], Client id [$clientID]" -ForegroundColor "Green"
            Write-Host "------------------------------------------------------------------------" -ForegroundColor White
            #endregion ___________________________________________________________________________________________________________________

            #region Rapport_Inventaire ______________________________________________________________________________________________________
            # ADD nouvelle feuille -----------------------------------------------------------------
            Write-Host "--> Start Report inventory for Client Name [$clientName], Client id [$clientID]" -ForegroundColor "Green"                                                                                                
            
            # Création du contenu
            $NewWorkSheet = $workbook.Worksheets.Add()                                                  # Création d'un nouvel onglet  
            $finalWorkSheet = $workbook.Worksheets.Item(1)                                              # Sélection du nouvel onglet
            $finalWorkSheet.Name = "Rapports"                                                           # Nom de l'onglet
            
            # Création du premier titre (première ligne)
            $Column = 1
            $row = 1
            $color = $FirstRowColor
            $global:Analyze_Synthese = @()                                                                                         
            $finalWorkSheet.Cells.Item($row,$Column) = "Rapports Partager"                                                                                        
            $range = $finalWorkSheet.Range("A"+$row,"E"+$row)                                           # Défini la zone
            $range.select() | Out-Null                                                                  # Sélectionne la zone
            $range.Interior.ColorIndex = 14                                                             # Change la couleur des cellule
            $range.Font.Name = "Arial"                                                                  # Fixe la police
            $range.Font.Size = 18                                                                       # Fixe la taille de la police
            $range.Font.Bold = $True                                                                    # Met en carractères gras
            $range.Font.ColorIndex = 2                                                                  # Fixe la couleur de la police
            $range.EntireColumn.AutoFit() | Out-Null                                                    # Auto ajustement de la taille de la colonne
            $range.HorizontalAlignment = $xlCenter                                                      # Aligne horizontale le contenu des cellules
            $range.VerticalAlignment = $xlBottom                                                        # Aligne vertivalement le contenu des cellules
            $range.columnwidth = 30                                                                     # Largeur des colonnes
            $range.MergeCells = $True                                                                   # Fusionne les celulles
            
            # Création des entêtes                                                                                                
            $row = $row + 2
            $finalWorkSheet.Cells.Item($row,$Column) = "Nom" 
            $Column++
            $finalWorkSheet.Cells.Item($row,$Column) = "Type" 
            $Column++
            $finalWorkSheet.Cells.Item($row,$Column) = "Frequence"
            $Column++ 
            $finalWorkSheet.Cells.Item($row,$Column) = "Date Modif."
            $Column++ 
            $finalWorkSheet.Cells.Item($row,$Column) = "Destinataires"

            # Mise en forme des entêtes
            $range = $finalWorkSheet.Range("A"+$row,"E"+$row)                                           # Défini la zone
            $range.select() | Out-Null                                                                  # Sélectionne la zone
            $range.Interior.ColorIndex = 14                                                             # Change la couleur des cellule
            $range.Font.Name = "Arial"                                                                  # Fixe la police
            $range.Font.Size = 18                                                                       # Fixe la taille de la police
            $range.Font.Bold = $True                                                                    # Met en carractères gras
            $range.Font.ColorIndex = 2                                                                  # Fixe la couleur de la police
            $range.HorizontalAlignment = $xlCenter                                                      # Aligne horizontale le contenu des cellules
            $range.VerticalAlignment = $xlBottom                                                        # Aligne vertivalement le contenu des cellules
            #$range.columnwidth = 50                                                                     # Largeur des colonnes
            $range.AutoFilter() | Out-Null                                                              # Ajout du filtre automatique                                                                                                

            Write-Host ("--> Nb Share Reports : " + [int]($ShareReports.id).count) -ForegroundColor Blue                                                                                                
            Foreach ($Share in $ShareReports){
                Write-Host "." -NoNewline -ForegroundColor Blue
                if($Debug -eq $true){Write-Host "Repport Name : " $Share.name -ForegroundColor Yellow}
                if($Debug -eq $true){Write-Host "Repport Type : " $Share.type -ForegroundColor Yellow}
                switch ($Share.frequency) {
                    0 { $frequency = "Quotidien" }
                    1 { $frequency = "Hebdomadaire" }
                    2 { $frequency = "Mensuel" }
                    Default {$frequency = "??"}
                }
                
                if($Debug -eq $true){Write-Host "Repport Frequency : " $frequency -ForegroundColor Yellow}
                if($Debug -eq $true){Write-Host "Repport LastModified : " $Share.lastModified -ForegroundColor Yellow}
                $listRecipient = ""
                foreach($recipient in $Share.recipients){
                    $listRecipient = $listRecipient + ($recipient.firstname +" "+ $recipient.lastname + " `n`r")
                }

                if($Debug -eq $true){Write-Host ("Repport Recipient : " + $listRecipient) -ForegroundColor Yellow}

                $Column = 1
                $row++
                $Range = $finalWorkSheet.Range("A"+$row).Select()
                $finalWorkSheet.Cells.Item($row,$Column) = [string]$Share.name                                          # Add data
                $Column++                                                                                               # next column
                $finalWorkSheet.Cells.Item($row,$Column) = [string]$Share.type                                          # Add data
                $Column++                                                                                               # next column
                $finalWorkSheet.Cells.Item($row,$Column) = [string]$frequency                                           # Add data
                $Column++                                                                                               # next column
                $finalWorkSheet.Cells.Item($row,$Column) = [string](get-date($Share.lastModified) -format "MM/dd/yyyy")           # Add data
                $Column++                                                                                               # next column
                $finalWorkSheet.Cells.Item($row,$Column) = [string]$listRecipient                                       # Add data

                # Mise en page du document
                $range = $finalWorkSheet.Range(("A{0}" -f $row),("E{0}" -f $row))                                     # Selectionne la ligne en cours
                $range.Select() | Out-Null                                                                            # Select range
                $range.Interior.ColorIndex = $color                                                                   # Change la couleur des cellule
                $range.Font.Size = 12                                                                                 # Fixe la taille de la police
                $range.EntireColumn.AutoFit() | Out-Null                                                              # Auto ajustement de la taille de la colonne
                $range.VerticalAlignment = -4160                                                                      # Aligne vertivalement le contenu des cellules
                                    
                if($color -EQ $FirstRowColor){$color = $SecondtRowColor} ELSE {$color = $FirstRowColor}               # Change color line
            }
            
            # Création du second titre (première ligne)
            $Column = 1
            $row = $row + 4
            $color = $FirstRowColor
            $global:Analyze_Synthese = @() 
            $finalWorkSheet.Cells.Item($row,$Column) = "Rapports Publier"
            $range = $finalWorkSheet.Range("A"+$row,"E"+$row)                                           # Défini la zone
            $range.select() | Out-Null                                                                  # Sélectionne la zone
            $range.Interior.ColorIndex = 14                                                             # Change la couleur des cellule
            $range.Font.Name = "Arial"                                                                  # Fixe la police
            $range.Font.Size = 18                                                                       # Fixe la taille de la police
            $range.Font.Bold = $True                                                                    # Met en carractères gras
            $range.Font.ColorIndex = 2                                                                  # Fixe la couleur de la police
            #$range.EntireColumn.AutoFit() | Out-Null                                                    # Auto ajustement de la taille de la colonne
            $range.HorizontalAlignment = $xlCenter                                                      # Aligne horizontale le contenu des cellules
            $range.VerticalAlignment = $xlBottom                                                        # Aligne vertivalement le contenu des cellules
            $range.columnwidth = 30                                                                     # Largeur des colonnes
            $range.MergeCells = $True                                                                   # Fusionne les celulles                                                                                               
            
            # Création des entêtes                                                                                                
            $row = $row + 2
            $finalWorkSheet.Cells.Item($row,$Column) = "Nom" 
            $Column++
            $finalWorkSheet.Cells.Item($row,$Column) = "Classeur" 
            $Column++
            $finalWorkSheet.Cells.Item($row,$Column) = "Tableau de bord"
            $Column++ 
            $finalWorkSheet.Cells.Item($row,$Column) = "Date Modif."
            $Column++ 
            $finalWorkSheet.Cells.Item($row,$Column) = "Workspaces"

            # Mise en forme des entêtes
            $range = $finalWorkSheet.Range("A"+$row,"E"+$row)                                           # Défini la zone
            $range.select() | Out-Null                                                                  # Sélectionne la zone
            $range.Interior.ColorIndex = 14                                                             # Change la couleur des cellule
            $range.Font.Name = "Arial"                                                                  # Fixe la police
            $range.Font.Size = 18                                                                       # Fixe la taille de la police
            $range.Font.Bold = $True                                                                    # Met en carractères gras
            $range.Font.ColorIndex = 2                                                                  # Fixe la couleur de la police
            $range.HorizontalAlignment = $xlCenter                                                      # Aligne horizontale le contenu des cellules
            $range.VerticalAlignment = $xlBottom                                                        # Aligne vertivalement le contenu des cellules
            #$range.columnwidth = 50                                                                    # Largeur des colonnes
            #$range.AutoFilter() | Out-Null                                                              # Ajout du filtre automatique
            
            Write-Host ""
            Write-Host ("--> Nb Publish Reports : " + [int]($PublishReports.id).count) -ForegroundColor Blue 
            Foreach ($Publish in $PublishReports){
                Write-Host "." -NoNewline -ForegroundColor Blue
                if($Debug -eq $true){Write-Host "Repport Name : " $Publish.name -ForegroundColor Yellow}
                if($Debug -eq $true){Write-Host "Repport workbookName : " $Publish.workbookName -ForegroundColor Yellow}
                if($Debug -eq $true){Write-Host "Repport dashboardName : " $Publish.dashboardName -ForegroundColor Yellow}
                if($Debug -eq $true){Write-Host "Repport lastUpdated : " get-date($Publish.lastUpdated) -format "dd/MM/yyyy" -ForegroundColor Yellow}
                $listworkspaces = ""
                foreach($PublishWorkspace in $Publish.workspaces){
                    $WorkspaceName = $workspaces | Select-Object -Property id, name | Where-Object {$_.id -eq $PublishWorkspace} | select-object -Property name
                    $listworkspaces = $listworkspaces + ($WorkspaceName.name + " `n`r")
                }
                
                if($Debug -eq $true){Write-Host "Repport workspaces : " $listworkspaces -ForegroundColor Yellow}

                $Column = 1
                $row++
                $Range = $finalWorkSheet.Range("A"+$row).Select()
                $finalWorkSheet.Cells.Item($row,$Column) = [string]$Publish.name                                        # Add data
                $Column++                                                                                               # next column
                $finalWorkSheet.Cells.Item($row,$Column) = [string]$Publish.workbookName                                # Add data
                $Column++                                                                                               # next column
                $finalWorkSheet.Cells.Item($row,$Column) = [string]$Publish.dashboardName                               # Add data
                $Column++                                                                                               # next column
                $finalWorkSheet.Cells.Item($row,$Column) = [string](get-date($Publish.lastUpdated) -format "dd/MM/yyyy")          # Add data
                $Column++                                                                                               # next column
                $finalWorkSheet.Cells.Item($row,$Column) = [string]$listworkspaces                                      # Add data

                # Mise en page du document
                $range = $finalWorkSheet.Range(("A{0}" -f $row),("E{0}" -f $row))                                     # Selectionne la ligne en cours
                $range.Select() | Out-Null                                                                            # Select range
                $range.Interior.ColorIndex = $color                                                                   # Change la couleur des cellule
                $range.Font.Size = 12                                                                                 # Fixe la taille de la police
                $range.EntireColumn.AutoFit() | Out-Null                                                              # Auto ajustement de la taille de la colonne
                $range.VerticalAlignment = -4160                                                                      # Aligne vertivalement le contenu des cellules
                                        
                if($color -EQ $FirstRowColor){$color = $SecondtRowColor} ELSE {$color = $FirstRowColor}               # Change color line
            }    

            Write-Host ""
            Write-Host "--> END Zones inventory for Client Name [$clientName], Client id [$clientID]" -ForegroundColor "Green"
            Write-Host "------------------------------------------------------------------------" -ForegroundColor White                                                                                                
            #endregion ___________________________________________________________________________________________________________________                                                                                                

            #region Conso_Inventaire _____________________________________________________________________________________________________
            # ADD nouvelle feuille -----------------------------------------------------------------
            Write-Host "--> Start Conso for Client Name [$clientName], Client id [$clientID]" -ForegroundColor "Green"
            $NewWorkSheet = $workbook.Worksheets.Add()                                                  # Création d'un nouvel onglet  
            $finalWorkSheet = $workbook.Worksheets.Item(1)                                              # Sélection du nouvel onglet     
            $finalWorkSheet.Name = "Conso"                                                              # Nom de l'onglet

            # Création du titre (première ligne)
            $Column = 1
            $row = 1
            $color = $FirstRowColor
            $global:Analyze_conso = @()
            $Range = $finalWorkSheet.Range("A"+$row).Select()
            $finalWorkSheet.Cells.Item($row,$Column) = "Actions par utilisateur sur $global:Action_period jours"
            $range = $finalWorkSheet.Range(("A"+$row),("E"+$row))
            $range.Select() | Out-Null
            $range.Font.Size = 18
            $Range.Font.Bold = $True
            $Range.Font.ColorIndex = 2
            $range.MergeCells = $true
            $range.HorizontalAlignment = -4108
            $range.BorderAround($xlContinuous,$xlThick,$xlColorIndexBlue)
            $range.Interior.ColorIndex = 14                                                             # Change la couleur des cellule

            # Création du contenu
            $Column = 2
            $row = 3
            $finalWorkSheet.Cells.Item($row,$Column) = "Prenom" 
            $Column++
            $finalWorkSheet.Cells.Item($row,$Column) = "Nom"
            $Column++
            $finalWorkSheet.Cells.Item($row,$Column) = "Nb_Action"

            # Format first line
            $range = $finalWorkSheet.Range(("B"+$row),("D"+$row))                       # Sélectionne les cellules
            $range.Interior.ColorIndex = 14                                             # Change la couleur des cellule
            $range.Font.Name = "Arial"                                                  # Fixe la police
            $range.Font.Size = 18                                                       # Fixe la taille de la police
            $range.Font.Bold = $True                                                    # Met en carractères gras
            $range.Font.ColorIndex = 2                                                  # Fixe la couleur de la police
            $range.EntireColumn.AutoFit() | Out-Null                                    # Auto ajustement de la taille de la colonne
            $range.VerticalAlignment = -4160                                            # Aligne vertivalement le contenu des cellules

            Write-Host "--> Total Nb actions : " $actions.count -ForegroundColor Blue
            #List of Unique users
            $UserList = ($actions | Sort-Object -Unique -Property @{Expression = "userLastname";},
                                                                  @{Expression = "userFirstname";})
            $NbDifUser = ($UserList).count
            Write-Host "--> Total Unique User : " $NbDifUser -ForegroundColor Blue
            
            
            # Formated aray with all actions
            $Result1 = @()
            Foreach ($action in $actions)
            {
                #Formatting the array
                $Result1 += New-Object -TypeName psobject -Property @{Firstname = [string]$action.userFirstname;`
                                                                    Lastname = [string]$action.userLastname;`
                                                                    Type = [string]$action.type;`
                                                                    Action = [string]$action.action;`
                                                                    Source = [string]$action.source;`
                                                                    Date = [string]$action.timestamp;
                                                                    } | Select-Object Firstname,Lastname,Type,Action,Source,Date
            }

            # Show Formated results
            if($Debug -eq $true){$Result1 | Out-GridView -Title "all actions"}  
            
            #Write-Host "------------ DETAIL --------------"
            # Formated aray with unique user
            $Result2 = @()
            $more_500_action = 0
            foreach ($user in $UserList){
                Write-Host "." -NoNewline -ForegroundColor Blue
                #Count number actions by user
                [int]$nbAction = ($actions | Where-Object -FilterScript {($_.userLastname -match $user.userLastname) -and ($_.userFirstname -match $user.userFirstname)}).count
                if($Debug -eq $true){Write-Host "Nb actions for user : "$user.userFirstname $user.userLastname " = " [int]$nbAction -ForegroundColor Yellow}
                #Formatting the array
                $Result2 += New-Object -TypeName psobject -Property @{Prenom = [string]$user.userFirstname;`
                                                                    Nom = [string]$user.userLastname;`
                                                                    Nb_Action = [int]$nbAction;
                                                                    } | Select-Object Firstname,Lastname,Nb_Action
            
                $Column = 2
                $row++
                $Range = $finalWorkSheet.Range("B"+$row).Select()
                $finalWorkSheet.Cells.Item($row,$Column) = [string]$user.userFirstname                                  # Add data
                $Column++                                                                                               # next column
                $finalWorkSheet.Cells.Item($row,$Column) = [string]$user.userLastname                                   # Add data
                $Column++                                                                                               # next column
                $finalWorkSheet.Cells.Item($row,$Column) = [int]$nbAction                                               # Add data
                
                if($nbAction -gt 500){$more_500_action++}          # Compte le nombre d'action pour l'ALANYZE

                # Mise en page du document
                $range = $finalWorkSheet.Range(("B{0}" -f $row),("D{0}" -f $row))                                     # Selectionne la ligne en cours
                $range.Select() | Out-Null                                                                            # Select range
                $range.Interior.ColorIndex = $color                                                                   # Change la couleur des cellule
                $range.Font.Size = 12                                                                                 # Fixe la taille de la police
                $range.EntireColumn.AutoFit() | Out-Null                                                              # Auto ajustement de la taille de la colonne
                $range.VerticalAlignment = -4160                                                                      # Aligne vertivalement le contenu des cellules
                                                                                            
                if($color -EQ $FirstRowColor){$color = $SecondtRowColor} ELSE {$color = $FirstRowColor}               # Change color line
             
            }
            
            # Show Formated results
            if($Debug -eq $true){$Result2 | Out-GridView -Title "unique user"}   
            
            # ANALYZE _____________________________________________________________________________________
            if($more_500_action -gt 0){$global:Analyze_conso += New-Object -TypeName psobject -Property @{Comment = [string]"Il y a $more_500_action utilisateur(s) qui ont effectue plus de 500 actions en $global:Action_period jours.";`
                                                                                                    } | Select-Object Comment}
            # _____________________________________________________________________________________________

            $Range = $finalWorkSheet.Range("A1").Select()
            $finalWorkSheet.Application.ActiveWindow.SplitRow = 1                                                     # Select First row for freeze the shutters
            $finalWorkSheet.Application.ActiveWindow.FreezePanes = $true                                              # Freeze the shutters
            Write-Host ""
            Write-Host "--> END conso inventory for Client Name [$clientName], Client id [$clientID]" -ForegroundColor "Green"
            Write-Host "------------------------------------------------------------------------" -ForegroundColor White
            #endregion ___________________________________________________________________________________________________________________

            #region Stats_Inventaire _____________________________________________________________________________________________________
            # ADD nouvelle feuille -----------------------------------------------------------------
            Write-Host "--> Start STATS for Client Name [$clientName], Client id [$clientID]" -ForegroundColor "Green"
            # Création du contenu
            $NewWorkSheet = $workbook.Worksheets.Add()                                                  # Création d'un nouvel onglet  
            $finalWorkSheet = $workbook.Worksheets.Item(1)                                              # Sélection du nouvel onglet
            $finalWorkSheet.Name = "STATS"                                                              # Nom de l'onglet
            
            # Récupération des information du compte client
            $client = $clients | Where-Object -FilterScript {$_.id -eq $clientID}  
            
            # Création du premier titre (première ligne)
            $Column = 1
            $row = 1
            
            $global:Analyze_Synthese = @()                                                                                         
            $finalWorkSheet.Cells.Item($row,$Column) = "Statistiques globaux du perimetre"                                                                                        
            $range = $finalWorkSheet.Range("A"+$row,"D"+$row)                                           # Défini la zone
            $range.select() | Out-Null                                                                  # Sélectionne la zone
            $range.Interior.ColorIndex = 14                                                             # Change la couleur des cellule
            $range.Font.Name = "Arial"                                                                  # Fixe la police
            $range.Font.Size = 18                                                                       # Fixe la taille de la police
            $range.Font.Bold = $True                                                                    # Met en carractères gras
            $range.Font.ColorIndex = 2                                                                  # Fixe la couleur de la police
            $range.EntireColumn.AutoFit() | Out-Null                                                    # Auto ajustement de la taille de la colonne
            $range.HorizontalAlignment = $xlCenter                                                      # Aligne horizontale le contenu des cellules
            $range.VerticalAlignment = $xlBottom                                                        # Aligne vertivalement le contenu des cellules
            $range.columnwidth = 30                                                                     # Largeur des colonnes
            $range.MergeCells = $True                                                                   # Fusionne les celulles
            
            # Création des entêtes
            $Column = 2
            $row = 3
            $finalWorkSheet.Cells.Item($row,$Column) = "Type de scénario"
            $Column++
            $finalWorkSheet.Cells.Item($row,$Column) = "NB"

            # Mise en forme des entêtes
            $range = $finalWorkSheet.Range("B"+$row,"C"+$row)                                           # Défini la zone
            $range.select() | Out-Null                                                                  # Sélectionne la zone
            $range.Interior.ColorIndex = 14                                                             # Change la couleur des cellule
            $range.Font.Name = "Arial"                                                                  # Fixe la police
            $range.Font.Size = 18                                                                       # Fixe la taille de la police
            $range.Font.Bold = $True                                                                    # Met en carractères gras
            $range.Font.ColorIndex = 2                                                                  # Fixe la couleur de la police
            $range.HorizontalAlignment = $xlCenter                                                      # Aligne horizontale le contenu des cellules
            $range.VerticalAlignment = $xlBottom                                                        # Aligne vertivalement le contenu des cellules
            $range.columnwidth = 50                                                                     # Largeur des colonnes
            
            # Création des 2 colonnes
            $Column = 2
            $row = 4
            $color = $SecondtRowColor

            # Compte le nombre de scénarios
            $uri ="$API/adm-api/scenarios?clientId=$clientId"
            $ListScenarios = Invoke-RestMethod -Uri $uri -Method POST -Headers $headers
            $ActiveScenarios = $ListScenarios | where-object {$_.active -eq 1}
            Write-Host "--> Nb scenarios : " $ListScenarios.count -ForegroundColor Blue                                                                                        
            
            # Compte le nombre de type de scénario actif
            $nb_HTTPR = 0
            $nb_BPL = 0
            $nb_WEB = 0
            $nb_DESKTOP = 0
            $nb_Android = 0
            $nb_IOS = 0
            $nb_API = 0
            $nb_EXPERT = 0
            $nb_other = 0

            Foreach ($scenario in $ActiveScenarios)
            {
                $plugins = $scenario.plugins
                if($Debug -eq $true){Write-Host "Scenario plugins : " $plugins.Name[1] -ForegroundColor Yellow}
                switch ($plugins.Name[1]) {
                    "HTTP REQUEST" { $nb_HTTPR++;break }
                    "BROWSER PAGE LOAD" { $nb_BPL++;break }
                    "WEB" { $nb_WEB++;break }
                    "DESKTOP" { $nb_DESKTOP++ }
                    "MOBILE-ANDROID" { $nb_Android++;break }
                    "MOBILE-IPHONE" { $nb_IOS++;break }
                    "API"{ $nb_API++;break }
                    "NEWTEST" { $nb_NWT++;break}
                    "EXPERT" { $nb_EXPERT++;break}
                    Default {$nb_other++;break}
                }
            }

            # Création du contenu des 2 colonnes                                                                                       
            if($color -EQ $FirstRowColor){$color = $SecondtRowColor} ELSE {$color = $FirstRowColor}
            $finalWorkSheet.Cells.Item($row,$Column) = "HTTP Request" 
            $finalWorkSheet.Cells.Item($row,$Column).Interior.ColorIndex = $color                                           # Add data
            $finalWorkSheet.Cells.Item($row,$Column + 1) = [int]$nb_HTTPR                                                   # Add data
            $finalWorkSheet.Cells.Item($row,$Column + 1).Interior.ColorIndex = $color

            $row++
            if($color -EQ $FirstRowColor){$color = $SecondtRowColor} ELSE {$color = $FirstRowColor}                         # Change color line
            $finalWorkSheet.Cells.Item($row,$Column) = "Browser Page Load"                                                  # Add data
            $finalWorkSheet.Cells.Item($row,$Column).Interior.ColorIndex = $color 
            $finalWorkSheet.Cells.Item($row,$Column + 1) = [int]$nb_BPL                                                     # Add data
            $finalWorkSheet.Cells.Item($row,$Column + 1).Interior.ColorIndex = $color 
            
            $row++
            if($color -EQ $FirstRowColor){$color = $SecondtRowColor} ELSE {$color = $FirstRowColor}                         # Change color line
            $finalWorkSheet.Cells.Item($row,$Column) = "Web"                                                                # Add data
            $finalWorkSheet.Cells.Item($row,$Column).Interior.ColorIndex = $color 
            $finalWorkSheet.Cells.Item($row,$Column + 1) = [int]$nb_WEB                                                     # Add data
            $finalWorkSheet.Cells.Item($row,$Column + 1).Interior.ColorIndex = $color 

            $row++
            if($color -EQ $FirstRowColor){$color = $SecondtRowColor} ELSE {$color = $FirstRowColor}                         # Change color line
            $finalWorkSheet.Cells.Item($row,$Column) = "Business APP"                                                       # Add data
            $finalWorkSheet.Cells.Item($row,$Column).Interior.ColorIndex = $color 
            $finalWorkSheet.Cells.Item($row,$Column + 1) = [int]$nb_DESKTOP                                                 # Add data
            $finalWorkSheet.Cells.Item($row,$Column + 1).Interior.ColorIndex = $color 

            $row++
            if($color -EQ $FirstRowColor){$color = $SecondtRowColor} ELSE {$color = $FirstRowColor}                         # Change color line
            $finalWorkSheet.Cells.Item($row,$Column) = "Mobile Android"                                                     # Add data
            $finalWorkSheet.Cells.Item($row,$Column).Interior.ColorIndex = $color 
            $finalWorkSheet.Cells.Item($row,$Column + 1) = [int]$nb_Android                                                 # Add data
            $finalWorkSheet.Cells.Item($row,$Column + 1).Interior.ColorIndex = $color

            $row++
            if($color -EQ $FirstRowColor){$color = $SecondtRowColor} ELSE {$color = $FirstRowColor}                         # Change color line
            $finalWorkSheet.Cells.Item($row,$Column) = "Mobile IOS"                                                         # Add data
            $finalWorkSheet.Cells.Item($row,$Column).Interior.ColorIndex = $color 
            $finalWorkSheet.Cells.Item($row,$Column + 1) = [int]$nb_IOS                                                     # Add data
            $finalWorkSheet.Cells.Item($row,$Column + 1).Interior.ColorIndex = $color

            $row++
            if($color -EQ $FirstRowColor){$color = $SecondtRowColor} ELSE {$color = $FirstRowColor}                         # Change color line
            $finalWorkSheet.Cells.Item($row,$Column) = "API"                                                                # Add data
            $finalWorkSheet.Cells.Item($row,$Column).Interior.ColorIndex = $color 
            $finalWorkSheet.Cells.Item($row,$Column + 1) = [int]$nb_API                                                     # Add data
            $finalWorkSheet.Cells.Item($row,$Column + 1).Interior.ColorIndex = $color

            $row++
            if($color -EQ $FirstRowColor){$color = $SecondtRowColor} ELSE {$color = $FirstRowColor}                         # Change color line
            $finalWorkSheet.Cells.Item($row,$Column) = "NEWTEST"                                                            # Add data
            $finalWorkSheet.Cells.Item($row,$Column).Interior.ColorIndex = $color 
            $finalWorkSheet.Cells.Item($row,$Column + 1) = [int]$nb_NWT                                                     # Add data
            $finalWorkSheet.Cells.Item($row,$Column + 1).Interior.ColorIndex = $color
            
            $row++
            if($color -EQ $FirstRowColor){$color = $SecondtRowColor} ELSE {$color = $FirstRowColor}                         # Change color line
            $finalWorkSheet.Cells.Item($row,$Column) = "EXPERT"                                                             # Add data
            $finalWorkSheet.Cells.Item($row,$Column).Interior.ColorIndex = $color 
            $finalWorkSheet.Cells.Item($row,$Column + 1) = [int]$nb_EXPERT                                                  # Add data
            $finalWorkSheet.Cells.Item($row,$Column + 1).Interior.ColorIndex = $color

            if($nb_other -gt 0){
                $row++
                if($color -EQ $FirstRowColor){$color = $SecondtRowColor} ELSE {$color = $FirstRowColor}                     # Change color line
                $finalWorkSheet.Cells.Item($row,$Column) = "Autre"                                                          # Add data
                $finalWorkSheet.Cells.Item($row,$Column).Interior.ColorIndex = $color 
                $finalWorkSheet.Cells.Item($row,$Column + 1) = [int]$nb_other                                               # Add data
                $finalWorkSheet.Cells.Item($row,$Column + 1).Interior.ColorIndex = $color
                $range = $finalWorkSheet.Range(("B"+$row),("C"+$row))                                                       # Sélectionne les cellules
                $range.select() | Out-Null                                                                                  # Sélectionne la zone
                $range.Font.ColorIndex = 3
            }

            # Mise en forme des colonnes
            $range = $finalWorkSheet.Range("B4","C"+$row)                                               # Défini la zone
            $range.select() | Out-Null                                                                  # Sélectionne la zone
            $range.Font.Name = "Arial"                                                                  # Fixe la police
            $range.Font.Size = 12                                                                       # Fixe la taille de la police
            $range.Font.Bold = $False                                                                   # Met en carractères gras
            $range.Font.ColorIndex = 55                                                                 # Fixe la couleur de la police
            $range.HorizontalAlignment = $xlCenter                                                      # Aligne horizontale le contenu des cellules
            $range.VerticalAlignment = $xlBottom                                                        # Aligne vertivalement le contenu des cellules
            $range.EntireColumn.AutoFit() | Out-Null                                                    # Auto ajustement de la taille de la colonne
            $range.WrapText = $false
            $range.Orientation = 0
            $range.ShrinkToFit = $false
            $range.ReadingOrder = $xlContext
            $range.MergeCells = $false

            # Mise en forme des cadres
            $range = $finalWorkSheet.Range("B4","C"+($row+1))                                           # Défini la zone
            $range.select() | Out-Null                                                                  # Sélectionne la zone
            $range.Borders.Item($xlInsideHorizontal).Weight = $xlThin
            $range.Borders.Item($xlInsideHorizontal).LineStyle = $xlContinuous
            $range.Borders.Item($xlInsideHorizontal).ColorIndex = 11                                                                                         
            #endregion ___________________________________________________________________________________________________________________

            #region Synthese _____________________________________________________________________________________________________________
            # ADD nouvelle feuille -----------------------------------------------------------------
            Write-Host "--> Start Synthese for Client Name [$clientName], Client id [$clientID]" -ForegroundColor "Green"
            # Création du contenu
            $NewWorkSheet = $workbook.Worksheets.Add()                                                  # Création d'un nouvel onglet  
            $finalWorkSheet = $workbook.Worksheets.Item(1)                                              # Sélection du nouvel onglet
            $finalWorkSheet.Name = "Synthese"                                                           # Nom de l'onglet
            
            # Récupération des information du compte client
            $client = $clients | Where-Object -FilterScript {$_.id -eq $clientID}  
            
            # Création du premier titre (première ligne)
            $Column = 1
            $row = 1
            
            $global:Analyze_Synthese = @()                                                                                         
            $finalWorkSheet.Cells.Item($row,$Column) = "Synthese compte"                                                                                        
            $range = $finalWorkSheet.Range("A"+$row,"D"+$row)                                           # Défini la zone
            $range.select() | Out-Null                                                                  # Sélectionne la zone
            $range.Interior.ColorIndex = 14                                                             # Change la couleur des cellule
            $range.Font.Name = "Arial"                                                                  # Fixe la police
            $range.Font.Size = 18                                                                       # Fixe la taille de la police
            $range.Font.Bold = $True                                                                    # Met en carractères gras
            $range.Font.ColorIndex = 2                                                                  # Fixe la couleur de la police
            $range.EntireColumn.AutoFit() | Out-Null                                                    # Auto ajustement de la taille de la colonne
            $range.HorizontalAlignment = $xlCenter                                                      # Aligne horizontale le contenu des cellules
            $range.VerticalAlignment = $xlBottom                                                        # Aligne vertivalement le contenu des cellules
            $range.columnwidth = 30                                                                     # Largeur des colonnes
            $range.MergeCells = $True                                                                   # Fusionne les celulles
            
            # Création des entêtes
            $Column = 2
            $row = 3
            $finalWorkSheet.Cells.Item($row,$Column) = "Themes"
            $Column++
            $finalWorkSheet.Cells.Item($row,$Column) = "Valeurs"

            # Mise en forme des entêtes
            $range = $finalWorkSheet.Range("B"+$row,"C"+$row)                                           # Défini la zone
            $range.select() | Out-Null                                                                  # Sélectionne la zone
            $range.Interior.ColorIndex = 14                                                             # Change la couleur des cellule
            $range.Font.Name = "Arial"                                                                  # Fixe la police
            $range.Font.Size = 18                                                                       # Fixe la taille de la police
            $range.Font.Bold = $True                                                                    # Met en carractères gras
            $range.Font.ColorIndex = 2                                                                  # Fixe la couleur de la police
            $range.HorizontalAlignment = $xlCenter                                                      # Aligne horizontale le contenu des cellules
            $range.VerticalAlignment = $xlBottom                                                        # Aligne vertivalement le contenu des cellules
            $range.columnwidth = 50                                                                     # Largeur des colonnes
            
            # Création des 2 colonnes
            $Column = 2
            $row = 4
            $color = $SecondtRowColor

            # Création du contenu des 2 colonnes                                                                                       
            if($color -EQ $FirstRowColor){$color = $SecondtRowColor} ELSE {$color = $FirstRowColor}
            $finalWorkSheet.Cells.Item($row,$Column) = "id" 
            $finalWorkSheet.Cells.Item($row,$Column).Interior.ColorIndex = $color                                           # Add data
            $finalWorkSheet.Cells.Item($row,$Column + 1) = $client.id                                                       # Add data
            $finalWorkSheet.Cells.Item($row,$Column + 1).Interior.ColorIndex = $color
             
            $row++
            if($color -EQ $FirstRowColor){$color = $SecondtRowColor} ELSE {$color = $FirstRowColor}                         # Change color line
            $finalWorkSheet.Cells.Item($row,$Column) = "Nom"                                                                # Add data
            $finalWorkSheet.Cells.Item($row,$Column).Interior.ColorIndex = $color 
            $finalWorkSheet.Cells.Item($row,$Column + 1) = $client.name                                                     # Add data
            $finalWorkSheet.Cells.Item($row,$Column + 1).Interior.ColorIndex = $color                                                                                         

            $row++
            if($color -EQ $FirstRowColor){$color = $SecondtRowColor} ELSE {$color = $FirstRowColor}                         # Change color line
            $finalWorkSheet.Cells.Item($row,$Column) = "Email"                                                              # Add data
            $finalWorkSheet.Cells.Item($row,$Column).Interior.ColorIndex = $color 
            $finalWorkSheet.Cells.Item($row,$Column + 1) = $client.email                                                    # Add data
            $finalWorkSheet.Cells.Item($row,$Column + 1).Interior.ColorIndex = $color                                                                                        

            $row++
            if($color -EQ $FirstRowColor){$color = $SecondtRowColor} ELSE {$color = $FirstRowColor}                         # Change color line
            $finalWorkSheet.Cells.Item($row,$Column) = "Date de Creation"                                                   # Add data
            $finalWorkSheet.Cells.Item($row,$Column).Interior.ColorIndex = $color 
            $finalWorkSheet.Cells.Item($row,$Column + 1) = get-date($client.dateCreated) -Format "dd/MM/yyyy"               # Add data
            $finalWorkSheet.Cells.Item($row,$Column + 1).Interior.ColorIndex = $color 
            
            $row++
            if($color -EQ $FirstRowColor){$color = $SecondtRowColor} ELSE {$color = $FirstRowColor}                         # Change color line
            $finalWorkSheet.Cells.Item($row,$Column) = "Partner"                                                            # Add data
            $finalWorkSheet.Cells.Item($row,$Column).Interior.ColorIndex = $color 
            $finalWorkSheet.Cells.Item($row,$Column + 1) = [string]$client.isPartner                                        # Add data
            $finalWorkSheet.Cells.Item($row,$Column + 1).Interior.ColorIndex = $color 

            $row++
            if($color -EQ $FirstRowColor){$color = $SecondtRowColor} ELSE {$color = $FirstRowColor}                         # Change color line
            $finalWorkSheet.Cells.Item($row,$Column) = "ID du compte parent"                                                # Add data
            $finalWorkSheet.Cells.Item($row,$Column).Interior.ColorIndex = $color 
            $finalWorkSheet.Cells.Item($row,$Column + 1) = $client.parentId                                                 # Add data
            $finalWorkSheet.Cells.Item($row,$Column + 1).Interior.ColorIndex = $color                                                                                         

            $row++
            if($color -EQ $FirstRowColor){$color = $SecondtRowColor} ELSE {$color = $FirstRowColor}                         # Change color line
            $finalWorkSheet.Cells.Item($row,$Column) = "Nom du compte parent"                                               # Add data
            $finalWorkSheet.Cells.Item($row,$Column).Interior.ColorIndex = $color 
            $finalWorkSheet.Cells.Item($row,$Column + 1) = $client.parentName                                               # Add data
            $finalWorkSheet.Cells.Item($row,$Column + 1).Interior.ColorIndex = $color                                                                                          

            $row++
            if($color -EQ $FirstRowColor){$color = $SecondtRowColor} ELSE {$color = $FirstRowColor}                         # Change color line
            $finalWorkSheet.Cells.Item($row,$Column) = "Actif"                                                              # Add data
            $finalWorkSheet.Cells.Item($row,$Column).Interior.ColorIndex = $color 
            $finalWorkSheet.Cells.Item($row,$Column + 1) = $client.isActif                                                  # Add data
            $finalWorkSheet.Cells.Item($row,$Column + 1).Interior.ColorIndex = $color                                                                                         

            $row++
            if($color -EQ $FirstRowColor){$color = $SecondtRowColor} ELSE {$color = $FirstRowColor}                         # Change color line
            $finalWorkSheet.Cells.Item($row,$Column) = "Quota"                                                              # Add data
            $finalWorkSheet.Cells.Item($row,$Column).Interior.ColorIndex = $color 
            $finalWorkSheet.Cells.Item($row,$Column + 1) = $client.quota                                                    # Add data
            $finalWorkSheet.Cells.Item($row,$Column + 1).Interior.ColorIndex = $color         
            
            $row++
            if($color -EQ $FirstRowColor){$color = $SecondtRowColor} ELSE {$color = $FirstRowColor}                         # Change color line
            $finalWorkSheet.Cells.Item($row,$Column) = "Nom entreprise"                                                     # Add data
            $finalWorkSheet.Cells.Item($row,$Column).Interior.ColorIndex = $color 
            $finalWorkSheet.Cells.Item($row,$Column + 1) = $client.entrepriseName                                           # Add data
            $finalWorkSheet.Cells.Item($row,$Column + 1).Interior.ColorIndex = $color

            $row++
            if($color -EQ $FirstRowColor){$color = $SecondtRowColor} ELSE {$color = $FirstRowColor}                         # Change color line
            $finalWorkSheet.Cells.Item($row,$Column) = "Duree du token"                                                     # Add data
            $finalWorkSheet.Cells.Item($row,$Column).Interior.ColorIndex = $color 
            $finalWorkSheet.Cells.Item($row,$Column + 1) = $client.tokenduration                                            # Add data
            $finalWorkSheet.Cells.Item($row,$Column + 1).Interior.ColorIndex = $color                                                                                        
            
            # Mise en forme des colonnes
            $range = $finalWorkSheet.Range("B4","C"+$row)                                               # Défini la zone
            $range.select() | Out-Null                                                                  # Sélectionne la zone
            $range.Font.Name = "Arial"                                                                  # Fixe la police
            $range.Font.Size = 12                                                                       # Fixe la taille de la police
            $range.Font.Bold = $False                                                                   # Met en carractères gras
            $range.Font.ColorIndex = 55                                                                 # Fixe la couleur de la police
            $range.HorizontalAlignment = $xlCenter                                                      # Aligne horizontale le contenu des cellules
            $range.VerticalAlignment = $xlBottom                                                        # Aligne vertivalement le contenu des cellules
            $range.EntireColumn.AutoFit() | Out-Null                                                    # Auto ajustement de la taille de la colonne
            $range.WrapText = $false
            $range.Orientation = 0
            $range.ShrinkToFit = $false
            $range.ReadingOrder = $xlContext
            $range.MergeCells = $false

            # Mise en forme des cadres
            $range = $finalWorkSheet.Range("B4","C"+($row+1))                                           # Défini la zone
            $range.select() | Out-Null                                                                  # Sélectionne la zone
            $range.Borders.Item($xlInsideHorizontal).Weight = $xlThin
            $range.Borders.Item($xlInsideHorizontal).LineStyle = $xlContinuous
            $range.Borders.Item($xlInsideHorizontal).ColorIndex = 11                                                                                        


            # Création du second titre (première ligne)
            $Column = 1
            $row = $row + 2
            
            $global:Analyze_Synthese = @() 
            $finalWorkSheet.Cells.Item($row,$Column) = "Synthese perimetre"
            $range = $finalWorkSheet.Range("A"+$row,"D"+$row)                                           # Défini la zone
            $range.select() | Out-Null                                                                  # Sélectionne la zone
            $range.Interior.ColorIndex = 14                                                             # Change la couleur des cellule
            $range.Font.Name = "Arial"                                                                  # Fixe la police
            $range.Font.Size = 18                                                                       # Fixe la taille de la police
            $range.Font.Bold = $True                                                                    # Met en carractères gras
            $range.Font.ColorIndex = 2                                                                  # Fixe la couleur de la police
            $range.HorizontalAlignment = $xlCenter                                                      # Aligne horizontale le contenu des cellules
            $range.VerticalAlignment = $xlBottom                                                        # Aligne vertivalement le contenu des cellules
            $range.MergeCells = $True                                                                   # Fusionne les celulles

            # Création des entêtes
            $Column = 2
            $row = $row + 2
            $finalWorkSheet.Cells.Item($row,$Column) = "Themes"
            $Column++
            $finalWorkSheet.Cells.Item($row,$Column) = "Quantites"
            
            # Mise en forme des entêtes
            $range = $finalWorkSheet.Range("B"+$row,"C"+$row)                                           # Défini la zone
            $range.select() | Out-Null                                                                  # Sélectionne la zone
            $range.Interior.ColorIndex = 14                                                             # Change la couleur des cellule
            $range.Font.Name = "Arial"                                                                  # Fixe la police
            $range.Font.Size = 18                                                                       # Fixe la taille de la police
            $range.Font.Bold = $True                                                                    # Met en carractères gras
            $range.Font.ColorIndex = 2                                                                  # Fixe la couleur de la police
            $range.HorizontalAlignment = $xlCenter                                                      # Aligne horizontale le contenu des cellules
            $range.VerticalAlignment = $xlBottom                                                        # Aligne vertivalement le contenu des cellules
            
            # Création des 2 colonnes
            $Column = 2
            $row++
            $color = $SecondtRowColor

            # Création du contenu des 2 colonnes                                                                                       
            if($color -EQ $FirstRowColor){$color = $SecondtRowColor} ELSE {$color = $FirstRowColor}
            $finalWorkSheet.Cells.Item($row,$Column) = "Scenario" 
            $finalWorkSheet.Cells.Item($row,$Column).Interior.ColorIndex = $color                                            # Add data
            $finalWorkSheet.Cells.Item($row,$Column + 1) = [int]$scenario_actif                                              # Add data
            $finalWorkSheet.Cells.Item($row,$Column + 1).Interior.ColorIndex = $color
            $finalWorkSheet.Cells.Item($row,$Column + 1).AddComment("Nombre de scénario ACTIF")                              # Add data into comment
            $finalWorkSheet.Cells.Item($row,$Column + 1).Comment.Shape.TextFrame.Characters().Font.Size = 8                  # Format comment
            $finalWorkSheet.Cells.Item($row,$Column + 1).Comment.Shape.TextFrame.Characters().Font.Bold = $False             # Format comment
             
            $row++
            if($color -EQ $FirstRowColor){$color = $SecondtRowColor} ELSE {$color = $FirstRowColor}                          # Change color line
            $finalWorkSheet.Cells.Item($row,$Column) = "Parcours"                                                            # Add data
            $finalWorkSheet.Cells.Item($row,$Column).Interior.ColorIndex = $color 
            $finalWorkSheet.Cells.Item($row,$Column + 1) = [int]$Scripts.message.name.count                                  # Add data
            $finalWorkSheet.Cells.Item($row,$Column + 1).Interior.ColorIndex = $color 
            
            $row++
            if($color -EQ $FirstRowColor){$color = $SecondtRowColor} ELSE {$color = $FirstRowColor}                          # Change color line
            $finalWorkSheet.Cells.Item($row,$Column) = "Planning"                                                            # Add data
            $finalWorkSheet.Cells.Item($row,$Column).Interior.ColorIndex = $color 
            $finalWorkSheet.Cells.Item($row,$Column + 1) = [int]$Plannings.id.count                                          # Add data
            $finalWorkSheet.Cells.Item($row,$Column + 1).Interior.ColorIndex = $color 
            
            $row++ 
            if($color -EQ $FirstRowColor){$color = $SecondtRowColor} ELSE {$color = $FirstRowColor}                          # Change color line                                                                                       
            $finalWorkSheet.Cells.Item($row,$Column) = "Alert"                                                               # Add data     
            $finalWorkSheet.Cells.Item($row,$Column).Interior.ColorIndex = $color                      
            $finalWorkSheet.Cells.Item($row,$Column + 1) = [int]$Alerts.id.count                                             # Add data
            $finalWorkSheet.Cells.Item($row,$Column + 1).Interior.ColorIndex = $color 
            
            $row++
            if($color -EQ $FirstRowColor){$color = $SecondtRowColor} ELSE {$color = $FirstRowColor}                          # Change color line
            $finalWorkSheet.Cells.Item($row,$Column) = "Utilisateurs"                                                        # Add data                           
            $finalWorkSheet.Cells.Item($row,$Column).Interior.ColorIndex = $color 
            $finalWorkSheet.Cells.Item($row,$Column + 1) = [int]$Users.id.count                                              # Add data  
            $finalWorkSheet.Cells.Item($row,$Column + 1).Interior.ColorIndex = $color                                                                                       
            
            $row++
            if($color -EQ $FirstRowColor){$color = $SecondtRowColor} ELSE {$color = $FirstRowColor}                         # Change color line
            $finalWorkSheet.Cells.Item($row,$Column) = "SSO"                                                                # Add data                           
            $finalWorkSheet.Cells.Item($row,$Column).Interior.ColorIndex = $color 
            $finalWorkSheet.Cells.Item($row,$Column + 1) = [int]$NB_SSO_Provider_Name                                       # Add data  
            $finalWorkSheet.Cells.Item($row,$Column + 1).Interior.ColorIndex = $color    

            $row++                                                                                        
            if($color -EQ $FirstRowColor){$color = $SecondtRowColor} ELSE {$color = $FirstRowColor}                          # Change color line
            $finalWorkSheet.Cells.Item($row,$Column) = "Applications"                                                        # Add data
            $finalWorkSheet.Cells.Item($row,$Column).Interior.ColorIndex = $color 
            $finalWorkSheet.Cells.Item($row,$Column + 1) = [int]$applications.id.count                                       # Add data 
            $finalWorkSheet.Cells.Item($row,$Column + 1).Interior.ColorIndex = $color                                                                                        
            
            $row++ 
            if($color -EQ $FirstRowColor){$color = $SecondtRowColor} ELSE {$color = $FirstRowColor}                          # Change color line
            $finalWorkSheet.Cells.Item($row,$Column) = "Workspaces"                                                          # Add data
            $finalWorkSheet.Cells.Item($row,$Column).Interior.ColorIndex = $color 
            $finalWorkSheet.Cells.Item($row,$Column + 1) = [int]$workspaces.id.count                                         # Add data
            $finalWorkSheet.Cells.Item($row,$Column + 1).Interior.ColorIndex = $color 
            
            $row++
            if($color -EQ $FirstRowColor){$color = $SecondtRowColor} ELSE {$color = $FirstRowColor}                          # Change color line
            $finalWorkSheet.Cells.Item($row,$Column) = "Zones"                                                               # Add data
            $finalWorkSheet.Cells.Item($row,$Column).Interior.ColorIndex = $color 
            $finalWorkSheet.Cells.Item($row,$Column + 1) = [int]$zones.id.count                                              # Add data
            $finalWorkSheet.Cells.Item($row,$Column + 1).Interior.ColorIndex = $color 

            $row++
            if($color -EQ $FirstRowColor){$color = $SecondtRowColor} ELSE {$color = $FirstRowColor}                          # Change color line
            $finalWorkSheet.Cells.Item($row,$Column) = "Robots Prives"                                                       # Add data
            $finalWorkSheet.Cells.Item($row,$Column).Interior.ColorIndex = $color 
            $finalWorkSheet.Cells.Item($row,$Column + 1) = [int]$unic_privat_site.count                                      # Add data
            $finalWorkSheet.Cells.Item($row,$Column + 1).Interior.ColorIndex = $color 
            if([int]$unic_privat_site.count -gt 0){
                $finalWorkSheet.Cells.Item($row,$Column + 1).AddComment(""+$unic_privat_site_list+"")                        # Add data into comment
                $finalWorkSheet.Cells.Item($row,$Column + 1).Comment.Shape.TextFrame.Characters().Font.Size = 8              # Format comment
                $finalWorkSheet.Cells.Item($row,$Column + 1).Comment.Shape.TextFrame.Characters().Font.Bold = $False         # Format comment
            }

            $row++  
            if($color -EQ $FirstRowColor){$color = $SecondtRowColor} ELSE {$color = $FirstRowColor}                          # Change color line                                                                                      
            $finalWorkSheet.Cells.Item($row,$Column) = "Robots Publics"   
            $finalWorkSheet.Cells.Item($row,$Column).Interior.ColorIndex = $color 
            $finalWorkSheet.Cells.Item($row,$Column + 1) = [int]$unic_public_site.count                                      # Add data
            $finalWorkSheet.Cells.Item($row,$Column + 1).Interior.ColorIndex = $color
            if([int]$unic_public_site.count -gt 0){                                                                                     
                $finalWorkSheet.Cells.Item($row,$Column + 1).AddComment(""+$unic_public_site_list+"")                        # Add data into comment
                $finalWorkSheet.Cells.Item($row,$Column + 1).Comment.Shape.TextFrame.Characters().Font.Size = 8              # Format comment
                $finalWorkSheet.Cells.Item($row,$Column + 1).Comment.Shape.TextFrame.Characters().Font.Bold = $False         # Format comment                                                                                        
            }
            
            $row++  
            if($color -EQ $FirstRowColor){$color = $SecondtRowColor} ELSE {$color = $FirstRowColor}                          # Change color line                                                                                      
            $finalWorkSheet.Cells.Item($row,$Column) = "Donnees Partagees"   
            $finalWorkSheet.Cells.Item($row,$Column).Interior.ColorIndex = $color 
            $finalWorkSheet.Cells.Item($row,$Column + 1) = [int]$shared_data.data.name.count                                 # Add data
            $finalWorkSheet.Cells.Item($row,$Column + 1).Interior.ColorIndex = $color
            
            $row++  
            if($color -EQ $FirstRowColor){$color = $SecondtRowColor} ELSE {$color = $FirstRowColor}                          # Change color line                                                                                      
            $finalWorkSheet.Cells.Item($row,$Column) = "Retry configuré"   
            $finalWorkSheet.Cells.Item($row,$Column).Interior.ColorIndex = $color 
            $finalWorkSheet.Cells.Item($row,$Column + 1) = [int]$NB_scenario_retry                                          # Add data
            $finalWorkSheet.Cells.Item($row,$Column + 1).Interior.ColorIndex = $color

            $row++ 
            if($color -EQ $FirstRowColor){$color = $SecondtRowColor} ELSE {$color = $FirstRowColor}                          # Change color line
            $finalWorkSheet.Cells.Item($row,$Column) = "Rum Tracker"                                                         # Add data
            $finalWorkSheet.Cells.Item($row,$Column).Interior.ColorIndex = $color 
            $finalWorkSheet.Cells.Item($row,$Column + 1) = [int]$trackers.trackerId.count                                    # Add data
            $finalWorkSheet.Cells.Item($row,$Column + 1).Interior.ColorIndex = $color 
            
            $row++ 
            if($color -EQ $FirstRowColor){$color = $SecondtRowColor} ELSE {$color = $FirstRowColor}                          # Change color line
            $finalWorkSheet.Cells.Item($row,$Column) = "Raports publies"                                                     # Add data
            $finalWorkSheet.Cells.Item($row,$Column).Interior.ColorIndex = $color 
            $finalWorkSheet.Cells.Item($row,$Column + 1) = [int]$PublishReports.id.count                                     # Add data
            $finalWorkSheet.Cells.Item($row,$Column + 1).Interior.ColorIndex = $color 
            
            $row++
            if($color -EQ $FirstRowColor){$color = $SecondtRowColor} ELSE {$color = $FirstRowColor}                          # Change color line
            $finalWorkSheet.Cells.Item($row,$Column) = "Reports Partages"                                                    # Add data
            $finalWorkSheet.Cells.Item($row,$Column).Interior.ColorIndex = $color 
            $finalWorkSheet.Cells.Item($row,$Column + 1) = [int]$ShareReports.id.count                                       # Add data 
            $finalWorkSheet.Cells.Item($row,$Column + 1).Interior.ColorIndex = $color                                                                                        
            
            # Mise en forme des colonnes B et C
            $range = $finalWorkSheet.Range("B19","C"+$row)                                              # Défini la zone
            $range.select() | Out-Null                                                                  # Sélectionne la zone
            $range.Font.Name = "Arial"                                                                  # Fixe la police
            $range.Font.Size = 12                                                                       # Fixe la taille de la police
            $range.Font.Bold = $False                                                                   # Met en carractères gras
            $range.Font.ColorIndex = 55                                                                 # Fixe la couleur de la police
            $range.HorizontalAlignment = $xlCenter                                                      # Aligne horizontale le contenu des cellules
            $range.VerticalAlignment = $xlBottom                                                        # Aligne vertivalement le contenu des cellules
            #$range.EntireColumn.AutoFit() | Out-Null                                                    # Auto ajustement de la taille de la colonne
            $range.WrapText = $false
            $range.Orientation = 0
            $range.ShrinkToFit = $false
            $range.ReadingOrder = $xlContext
            $range.MergeCells = $false

            # Mise en forme des cadres
            $range = $finalWorkSheet.Range("B19","C"+($row+1))                                          # Défini la zone
            $range.select() | Out-Null                                                                  # Sélectionne la zone
            $range.Borders.Item($xlInsideHorizontal).Weight = $xlThin
            $range.Borders.Item($xlInsideHorizontal).LineStyle = $xlContinuous
            $range.Borders.Item($xlInsideHorizontal).ColorIndex = 11
            
            # ANALYZE _____________________________________________________________________________________                                                    
            if($scripts.message.name.count -gt $Scenarios.count){$ecart = ($scripts.message.name.count - $Scenarios.count);$Analyze_Synthese += New-Object -TypeName psobject -Property @{Comment = [string]"Il y a $ecart parcours non configure(s).";`
                                                                                                                            } | Select-Object Comment}
            # _____________________________________________________________________________________________
            
            # AJOUTE L'ANALYZE DANS LE FICHIER XLS _______________________________________________________________________
            $row = $row + 3
            $Column = 2
            if(($global:Analyze_scenarios.comment).count -gt 0){
                foreach($comment in $global:Analyze_scenarios.comment){
                    $finalWorkSheet.Cells.Item($row,$Column) = [string]$comment
                    $range = $finalWorkSheet.Range(("B"+$row),("C"+$row))                                       # Sélectionne les cellules
                    $range.select() | Out-Null                                                                  # Sélectionne la zone
                    $range.Font.ColorIndex = 3
                    $row++
                }
            }
            if(($global:Analyze_Zones.comment).count -gt 0){
                foreach($comment in $global:Analyze_Zones.comment){
                    $finalWorkSheet.Cells.Item($row,$Column) = [string]$comment
                    $range = $finalWorkSheet.Range(("B"+$row),("C"+$row))                                       # Sélectionne les cellules
                    $range.select() | Out-Null                                                                  # Sélectionne la zone
                    $range.Font.ColorIndex = 3
                    $row++
                }
            }
            if(($global:Analyze_conso.comment).count -gt 0){
                foreach($comment in $global:Analyze_conso.comment){
                    $finalWorkSheet.Cells.Item($row,$Column) = [string]$comment
                    $range = $finalWorkSheet.Range(("B"+$row),("C"+$row))                                       # Sélectionne les cellules
                    $range.select() | Out-Null                                                                  # Sélectionne la zone
                    $range.Font.ColorIndex = 3
                    $row++
                }
            }

            if(($global:shared_data.data.name).count -eq 0){
                $finalWorkSheet.Cells.Item($row,$Column) = [string]"Le donnees partagees ne sont pas utilisees (" + ($shared_data.data.name).count + ")"
                $range = $finalWorkSheet.Range(("B"+$row),("C"+$row))                                       # Sélectionne les cellules
                $range.select() | Out-Null                                                                  # Sélectionne la zone
                $range.Font.ColorIndex = 3
                $row++
            }

            if(($global:Analyze_Rum.comment).count -gt 0){
                foreach($comment in $global:Analyze_Rum.comment){
                    $finalWorkSheet.Cells.Item($row,$Column) = [string]$comment
                    $range = $finalWorkSheet.Range(("B"+$row),("C"+$row))                                       # Sélectionne les cellules
                    $range.select() | Out-Null                                                                  # Sélectionne la zone
                    $range.Font.ColorIndex = 3
                    $row++
                }
            }
            
            $Result1 = @()
            $Result1 += New-Object -TypeName psobject -Property @{Scenario = [int]$scenario_actif;`
                                                                Parcours = [int]$scripts.message.name.count;`
                                                                Planning = [int]$Plannings.id.count;`
                                                                Alert = [int]$Alerts.id.count;`
                                                                User = [int]$Users.id.count;`
                                                                Application = [int]$applications.id.count;`
                                                                Workspace = [int]$workspaces.id.count;`
                                                                Zone = [int]$Zones.id.count;`
                                                                Rum_Tracker = [int]$trackers.trackerId.count;`
                                                                Publish_reports = [int]$PublishReports.id.count;`
                                                                Share_reports = [int]$ShareReports.id.count;`
                                                                } | Select-Object Scenario,Parcours,Planning,Alert,User,Application,Workspace,Zone,Rum_Tracker,Publish_reports,Share_reports

		    # Show Formated results
		    if($Debug -eq $true){$Result1 | Out-GridView -Title "Synthese"}                                          # Display results into GridView
            
            # AFFICHE L'ANALYZE
            if($Debug -eq $true){$global:Analyze_scenarios.comment + $global:Analyze_Zones.comment + $global:Analyze_conso.comment + $global:Analyze_Rum.comment | Out-GridView -Title "Analyze"}

            Write-Host "--> END Synthese for Client Name [$clientName], Client id [$clientID]" -ForegroundColor "Green"
            Write-Host "------------------------------------------------------------------------" -ForegroundColor White
            #endregion ___________________________________________________________________________________________________________________
            
            # Enregistrement du fichier ------------------------------------------------------------------------
            Write-Host ("Saving data and closing Excel file. [" + $Path + "\" + $NewFile +  "]") -ForegroundColor Green
            $workbook.SaveAs($ExcelPath)                                                                         # Save file
            # Fermeture du fichier 
            $Range = $workbook.Close()                                                                           # close file
            [System.Runtime.Interopservices.Marshal]::ReleaseComObject($objExcel) | Out-Null                     # close Excel

            # Ouvre le fichier XLS
            Write-Host "Open Excel file." -ForegroundColor Green
            Invoke-Item -Path "$Path\$NewFile" 
            return $Result_OK
        }catch{
            Error_popup("Erreur pendant l'inventaire")

            Write-Host -message "-------------------------------------------------------------" -ForegroundColor Red
            Write-Host -message "Erreur pendant l'inventaire ...." -BackgroundColor "Red"
            Write-Host -message $Error.exception.Message[0]
            Write-Host -message $Error[0]
            Write-Host -message $error[0].ScriptStackTrace
            Write-Host -message "-------------------------------------------------------------" -ForegroundColor red
        
            [System.Windows.Forms.MessageBox]::Show(`
                    "------------------------------------`n`r Erreur pendant l'inventaire`n`r------------------------------------`n`r",`
                    "Resultat",[System.Windows.Forms.MessageBoxButtons]::OKCancel,[System.Windows.Forms.MessageBoxIcon]::Warning)

            [System.Reflection.Assembly]::LoadWithPartialName('System.Windows.Forms')
            [System.Windows.Forms.MessageBox]::Show("Erreur pendant l'inventaire",'WARNING')

            #exit
        }finally{
           $Error.Clear()
        }
        $global:Result_OK = $global:Result_OK+1
    }else{
        Write-Host "--> Client Name or Client id not selected !" -ForegroundColor Red
        $global:Result_OK = $global:Result_OK+1
    }         
    return $Result_OK
}


#--------------------------------------------------------------------------------------------------------
function get_plannings(){
    #List all Plannings for customer
    if($Debug -eq $true){Write-Host ("Search Planning for client : "+ $clientID) -ForegroundColor Yellow}
    $uri1 ="$API/adm-api/plannings?clientId=$clientId"
    $global:Plannings = Invoke-RestMethod -Uri $uri1 -Method GET -Headers $headers
    if($Debug -eq $true){Write-Host ("Nb Plannings finded :" + $Plannings.count) -ForegroundColor Yellow}
}

function get_alerts(){
    #List all Alerts for customer
    if($Debug -eq $true){Write-Host ("Search Alerts for client : "+ $clientID) -ForegroundColor Yellow}
    $uri2 ="$API/adm-api/alerts?clientId=$clientId"
    $global:Alerts = Invoke-RestMethod -Uri $uri2 -Method GET -Headers $headers
    if($Debug -eq $true){Write-Host ("Nb alerts finded :" + $Alerts.count) -ForegroundColor Yellow}
    return $Alerts
}

function get_users(){
    #List all Users for customer
    if($Debug -eq $true){Write-Host ("Search Users for client : "+ $clientID) -ForegroundColor Yellow}
    $uri3 ="$API/adm-api/client/users?clientId=$clientId"
    $global:Users = Invoke-RestMethod -Uri $uri3 -Method GET -Headers $headers
    if($Debug -eq $true){Write-Host ("Nb users finded :" + $Users.count) -ForegroundColor Yellow}
    return $Users
}

function get_script($id_script){
    #List script for ID for customer
    if($Debug -eq $true){Write-Host ("Search parcours for client : "+ $clientID) -ForegroundColor Yellow}
    $uri4 =($API+"/script-api/script/"+$id_script+"?clientId="+$clientId)
    $global:Script = Invoke-RestMethod -Uri $uri4 -Method GET -Headers $headers
    if($Debug -eq $true){Write-Host ("Nb Scripts finded :" + $Script.message.name.count) -ForegroundColor Yellow}
    Return $Script
}

function get_scripts{
    #List all scripts for customer
    if($Debug -eq $true){Write-Host ("Search all parcours for client : "+ $clientID) -ForegroundColor Yellow}
    $uri4 =($API+"/script-api/scripts?clientId="+$clientId)
    $global:Scripts = Invoke-RestMethod -Uri $uri4 -Method POST -Headers $headers
    if($Debug -eq $true){Write-Host ("Nb Scripts finded :" + $Scripts.message.name.Count) -ForegroundColor Yellow}
    Return $Scripts
}

function get_applications(){
    #List all applications for customer
    if($Debug -eq $true){Write-Host ("Search applications for client : "+ $clientID) -ForegroundColor Yellow}
    $uri5 ="$API/adm-api/applications?clientId=$clientId"
    $global:applications = Invoke-RestMethod -Uri $uri5 -Method GET -Headers $headers
    if($Debug -eq $true){Write-Host ("Nb applications finded :" + $applications.count) -ForegroundColor Yellow}
    return $applications
}

function get_zones(){
    #List all zones for customer
    if($Debug -eq $true){Write-Host ("Search zones for client : "+ $clientID) -ForegroundColor Yellow}
    $uri5 ="$API/adm-api/zones?clientId=$clientId"
    $global:zones = Invoke-RestMethod -Uri $uri5 -Method GET -Headers $headers
    if($Debug -eq $true){Write-Host ("Nb zones finded :" + $zones.count) -ForegroundColor Yellow}
    return $zones
}

function get_sites(){
    #List all sites for customer
    if($Debug -eq $true){Write-Host ("Search zones for client : "+ $clientID) -ForegroundColor Yellow}
    $uri5 ="$API/adm-api/sites?clientId=$clientId"
    $global:sites = Invoke-RestMethod -Uri $uri5 -Method GET -Headers $headers
    if($Debug -eq $true){Write-Host ("Nb sites finded :" + $sites.count) -ForegroundColor Yellow}
    return $sites
}

function get_workspaces(){
    #List all workspaces for customer
    if($Debug -eq $true){Write-Host ("Search workspaces for client : "+ $clientID) -ForegroundColor Yellow}
    $uri5 ="$API/adm-api/workspaces?clientId=$clientId"
    $global:workspaces = Invoke-RestMethod -Uri $uri5 -Method GET -Headers $headers
    if($Debug -eq $true){Write-Host ("Nb workspaces finded :" + $workspaces.count) -ForegroundColor Yellow}
    return $workspaces
}

function get_actions(){
    #List all actions for customer
    $global:Action_period = 31
    if($Debug -eq $true){Write-Host ("Search all actions for client : "+ $clientID) -ForegroundColor Yellow}
    $global:DefaultDebut = ((Get-Date).adddays(-$global:Action_period)).tostring("yyyy-MM-dd") + " 00:00:00"          # Set the start date
    $global:DefaultFin = (Get-Date).tostring("yyyy-MM-dd") + " 00:00:00"                          # Set the End Date
    if($Debug -eq $true){Write-Host "Periode de $global:Action_period jours : du $DefaultDebut au $DefaultFin" -ForegroundColor Yellow}
    $uri ="$API/adm-api/actions?startTime=$DefaultDebut&endTime=$DefaultFin&clientId=$clientId"
    $global:actions = Invoke-RestMethod -Uri $uri -Method POST -Headers $headers
    if($Debug -eq $true){Write-Host ("--> Total Nb actions : " + $actions.count) -ForegroundColor Yellow}
    return $actions
}

function get_SSO_provider(){
    # List provider
    $uri ="$API/adm-api/identityproviders?linked_client=$clientId"      
    $global:SSO_Providers = Invoke-RestMethod -Uri $uri -Method GET -Headers $headers
    Return $SSO_Providers
}

function get_shared_data(){
    #List all shared data for customer
    if($Debug -eq $true){Write-Host ("Search all shared data for client : "+ $clientID) -ForegroundColor Yellow}
    $uri ="$API/adm-api/shareddata?clientId=$clientId"
    $global:shared_data = Invoke-RestMethod -Uri $uri -Method GET -Headers $headers
    return $shared_data
}

function get_rum_tracker(){
    #List all tracker for customer
    if($Debug -eq $true){Write-Host ("Search Trackers for client : "+ $clientID) -ForegroundColor Yellow}
    $uri5 ="$API/rum-restit/trk?clientId=$clientId"
    $global:trackers = Invoke-RestMethod -Uri $uri5 -Method GET -Headers $headers
    if($Debug -eq $true){Write-Host ("Nb trackers finded : " + $trackers.count) -ForegroundColor Yellow}
    return $trackers
}

function get_RUM_metrics(){
    #List all metrics
    if($Debug -eq $true){Write-Host ("Search All RUM metrics") -ForegroundColor Yellow}
    $uri5 ="$API/rum-restit/metrics?clientId=$clientId"
    $global:RUMmetrics = Invoke-RestMethod -Uri $uri5 -Method GET -Headers $headers
    if($Debug -eq $true){Write-Host ("Nb RUM Metrics finded :" + $RUMmetrics.count) -ForegroundColor Yellow}
    return $RUMmetrics
}

function get_RUM_overview([int]$trackerID, [int]$metricID){
    #List RUM overview for tracker ID
    if($Debug -eq $true){Write-Host ("List RUM overview for tracker ID : "+ $trackerID) -ForegroundColor Yellow}
    $global:Rum_period = 30
    $DefaultDebut = ((Get-Date).adddays(-$global:Rum_period)).tostring("yyyy-MM-dd") + " 00:00:00"          # Set the start date
    $DefaultFin = (Get-Date).tostring("yyyy-MM-dd") + " 00:00:00"                                # Set the End Date
    if($Debug -eq $true){Write-Host "Periode de $global:Rum_period jours : du $DefaultDebut au $DefaultFin" -ForegroundColor Yellow}    
    $uri5 ="$API/rum-restit/trk/$trackerID/results/$metricID/overview?clientId=$clientId&from=$DefaultDebut&to=$DefaultFin"
    $global:RUMoverview = Invoke-RestMethod -Uri $uri5 -Method POST -Headers $headers
    if($Debug -eq $true){Write-Host ("Nb RUM overview finded :" + $RUMoverview.count) -ForegroundColor Yellow}
    return $RUMoverview
}

function get_RUM_pageGroup($trackerID){
    #GET /rum-restit/trk/{trackerId}/urlgroups
    if($Debug -eq $true){Write-Host ("List RUM Parge Group for tracker ID : "+ $trackerID) -ForegroundColor Yellow}
    $uri5 ="$API/rum-restit/trk/$trackerID/urlgroups?clientId=$clientId"
    $global:RUMpageGroup = Invoke-RestMethod -Uri $uri5 -Method GET -Headers $headers
    if($Debug -eq $true){Write-Host ("Nb RUM overview finded :" + $RUMpageGroup.count) -ForegroundColor Yellow}
    return $RUMpageGroup
}

function get_RUM_business_Dimension($trackerID){
    #POST /rum-restit/customdim/business
    if($Debug -eq $true){Write-Host ("List RUM business Dimension for tracker ID : "+ $trackerID) -ForegroundColor Yellow}
    $uri5 ="$API/rum-restit/customdim/business?clientId=$clientId"
    
    $body = @{
        'trackers' = @($trackerID)
    }
     
    $json = $Body | ConvertTo-Json -Depth 5
    $global:RUMbusinessDimension = Invoke-RestMethod -Uri $uri5 -Method POST -Headers $headers -Body $json -ContentType 'application/json'
    if($Debug -eq $true){Write-Host ("Nb RUM business Dimension finded : " + $RUMbusinessDimension.count) -ForegroundColor Yellow}
    return $RUMbusinessDimension
}

function get_RUM_custom_Dimension($trackerID){
    #POST /rum-restit/customdim/custom
    if($Debug -eq $true){Write-Host ("List RUM custom Dimension for tracker ID : "+ $trackerID) -ForegroundColor Yellow}
    $uri5 ="$API/rum-restit/customdim/custom?clientId=$clientId"
    $body = @{
        'trackers' = @($trackerID)
    }
     
    $json = $Body | ConvertTo-Json -Depth 5
    $global:RUMCustomDimension = Invoke-RestMethod -Uri $uri5 -Method POST -Headers $headers -Body $json -ContentType 'application/json'
    if($Debug -eq $true){Write-Host ("Nb RUM custom Dimension finded : " + $RUMCustomDimension.count) -ForegroundColor Yellow}
    return $RUMCustomDimension
}

function get_RUM_infra_Dimension($trackerID){
    #POST /rum-restit/customdim/infra
    if($Debug -eq $true){Write-Host ("List RUM infra Dimension for tracker ID : "+ $trackerID) -ForegroundColor Yellow}
    $uri5 ="$API/rum-restit/customdim/infra?clientId=$clientId"
    $body = @{
        'trackers' = @($trackerID)
    }
     
    $json = $Body | ConvertTo-Json -Depth 5
    $global:RUMinfraDimension = Invoke-RestMethod -Uri $uri5 -Method POST -Headers $headers -Body $json -ContentType 'application/json'
    if($Debug -eq $true){Write-Host ("Nb RUM infra Dimension finded : " + $RUMinfraDimension.count) -ForegroundColor Yellow}
    return $RUMinfraDimension
}

function get_RUM_version_Dimension($trackerID){
    #POST /rum-restit/customdim/version
    if($Debug -eq $true){Write-Host ("List RUM version Dimension for tracker ID : "+ $trackerID) -ForegroundColor Yellow}
    $uri5 ="$API/rum-restit/customdim/version?clientId=$clientId"
    $body = @{
        'trackers' = @($trackerID)
    }
     
    $json = $Body | ConvertTo-Json -Depth 5
    $global:RUMversionDimension = Invoke-RestMethod -Uri $uri5 -Method POST -Headers $headers -Body $json -ContentType 'application/json'
    if($Debug -eq $true){Write-Host ("Nb RUM version Dimension finded : " + $RUMversionDimension.count) -ForegroundColor Yellow}
    return $RUMversionDimension
}

function get_Publish_reports(){
    #GET /adm-api/reports/views
    if($Debug -eq $true){Write-Host ("List All published reports for tracker ID : "+ $trackerID) -ForegroundColor Yellow}
    $uri5 ="$API/adm-api/reports/views?clientId=$clientId"
    $Publish_reports = Invoke-RestMethod -Uri $uri5 -Method GET -Headers $headers
    if($Debug -eq $true){Write-Host ("Nb published reports finded : " + $Publish_reports.count) -ForegroundColor Yellow}
    return $Publish_reports
}

function get_Share_reports(){
    #GET /adm-api/reports/schedules
    if($Debug -eq $true){Write-Host ("List All shared reports for tracker ID : "+ $trackerID) -ForegroundColor Yellow}
    $uri5 ="$API/adm-api/reports/schedules?clientId=$clientId"
    $Share_reports = Invoke-RestMethod -Uri $uri5 -Method GET -Headers $headers
    if($Debug -eq $true){Write-Host ("Nb shared reports finded : " + [int]$Share_reports.count) -ForegroundColor Yellow}
    return $Share_reports
}

function Search_planning($id_planning){
    #List planning
    if($Debug -eq $true){Write-Host "List planning" -ForegroundColor Yellow}
    $Planning = $Plannings | Where-Object { $_.id -eq $id_planning}
    $result_planning = $Planning.name
    Return $result_planning
}

function Search_alert($alert_ID){
    #List Alerte
    if($Debug -eq $true){Write-Host "List Alerte" -ForegroundColor Yellow}
    $User_name = ""
    if($Debug -eq $true){Write-Host "Search alert_ID = " $alert_ID -ForegroundColor Yellow}
    $Alert = $Alerts | Where-Object { $_.id -eq $alert_ID}
    $Alert_name += $Alert.name
    if($Debug -eq $true){Write-Host "Alert_Name = " $Alert_name -ForegroundColor Yellow}
    if($Debug -eq $true){Write-Host "NB Destinataires alerte : " $Alert.recipients.count -ForegroundColor Yellow}

    foreach ($recipient in $Alert.recipients){
        if($Debug -eq $true){Write-Host ("Destinataire alerte : " + $recipient.firstname + " " + $recipient.lastname) -ForegroundColor Yellow}
        #$User = $Users | Where-Object { $_.id -eq $recipient}
        $User_name += $recipient.firstname + " " + $recipient.lastname + " / "
        if($Debug -eq $true){Write-Host "User_name = " $User_name -ForegroundColor Yellow}
    }
    if($User_name -ne ""){
        $result_alert = "["+$Alert.name + "] : (" + $User_name + ") `n`r"
    }else{
        $result_alert = " - "
    }
    
    Return $result_alert
}

function Search_user($id_user){
    #List User
    if($Debug -eq $true){Write-Host "List User" -ForegroundColor Yellow}
    $User = $Users | Where-Object { $_.id -eq $id_user}
    $User_name += $User.firstname + " " + $User.lastname + " "
    Return $User_name
}

Function Error_popup($error_message){
    [reflection.assembly]::loadwithpartialname('System.Windows.Forms')
    [reflection.assembly]::loadwithpartialname('System.Drawing')
    $notify = new-object system.windows.forms.notifyicon
    $notify.icon = [System.Drawing.SystemIcons]::Warning
    $notify.visible = $true
    $notify.showballoontip(20,'WARNING',$error_message,[system.windows.forms.tooltipicon]::None)
}
#--------------------------------------------------------------------------------------------------------
#endregion

#region Main
#========================== START SCRIPT ======================================
Hide-Console
Authentication
List_Clients
#endregion