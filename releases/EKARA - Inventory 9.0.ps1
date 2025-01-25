#####################################################################################################
#                           Example of use of the EKARA API                                         #
#####################################################################################################
# Swagger interface : https://api.ekara.ip-label.net                                                #
# Personalized parameters : username / password / API-Key / Action_period / Debug mode              #
# Purpose of the script : Full inventory                                                            #
#####################################################################################################
# Author : Guy Sacilotto
# Last Update : 15/01/2025
[String]$global:Version = "9.0"

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
    adm-api/webhooks
    adm-api/apikeys
    adm-api/identityproviders
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
    adm-api/client/users
    adm-api/roles
#>

Clear-Host

#region VARIABLES ______________________________________________________________________________________________________
#========================== SETTING THE VARIABLES ===============================
$error.clear()
add-type -assemblyName "Microsoft.VisualBasic"
$PSDefaultParameterValues['*:Encoding'] = 'utf8'
$global:API = "https://api.ekara.ip-label.net"                                                # Webservice URL
$global:UserName = ""                                                                         # EKARA Account
$global:PlainPassword = ""                                                                    # EKARA Password
$global:API_KEY = ""                                                                          # EKARA Key account

$global:Result_OK = 0
$global:Result_KO = 0
[String]$global:date = [DateTime]::Now.ToString("yyyy-MM-dd HH-mm-ss")                        # Recupere la date du jour
$global:Action_period = 10                                                                    # Duree de l'historique des actions
$global:Debug = $False                                                                        # Value : True ou False

#Image au format Base64String.(correspond à une image PNG 96*96)
$base64PNG = "iVBORw0KGgoAAAANSUhEUgAAAGAAAABgCAMAAADVRocKAAADAFBMVEVdWVcsKSY4NDE7NzXtdAE3MzDs7OwQDwtbWFbe3d3f3t/1mCnyk0gtKifu7u41Mi9DPz3c29s6NjMvLCkxLiswLCfh4OA9Ojc0MS7b2dorKyuysLDNzMzudQMmJCLn5+dWU1EXFRPBv78xLithXFhGQ0BCPT3/14ZZVlNMSUc4NDHswKDzkEC8uro2MzNUUE5PTEqKiIfi39traGasq6ssKSbW1NQzLy0nGg3f3t45NTJIRUMgHRswLSpFOTn8lTcsKSZwbWo8OTXk5OTq6uppY1/f3t7e3NzFw8M3Iw2tbR7h2dTzlD8SEQ8wLSvwjTvvfRbh3Nzg398KCQjY19f0lTs0MC0vLCk4NDFAPDnzkig4NDH0kjtiYF2cmplzcW9ILhEsKSYsKSY1MS6VlJPveAvKyMg4MzDe3d3f3t7e3d3e3d2ioKCYYB4uKiddWVc4MzAsKSZfVU84NDHqlCzyj0DtlT2Rj47e3d3QZgIbGRWGhIPe3d3f3t57e3tXMw3e3d3f3t5dRigsKSVbWFZfXFphXVvufRcyLiyMbUnn3dbwgyQ6NjPzlkrT0dJRTUs5NTLf3t7xiTF4dXM6NTNhWlbrbAA2MS/zk0goJSPzk0Lf3t7zk0jibwLuexGWSwPvkzzf3t6NWhuCgH6fZR2lpKQ1Mi/f3t7xijPf3t51TBmlaB3zlEjg39+sfUcvLChcWFddWVbvfRX0lEHUiDc4NDDueQ/Hxsbe3d25t7drZ2b0lzrc3NzzhyXpcwR9URpuSB74tW1dWVf0mSpdWFcsKCb1mClfW1iWlZT0ljRdWVdaVVSvrazyk0i5WgHtdQSgnp3e3d3f3d1dUkRMSEZbTTu4cCzSfDb5v4r0n1H1nzhZVlOGakspKSmSkI7ueAvZ2NiBQASvraxfPBhBODGTj46tqqp9eXj1pkv1rl28YhHRpXrDw8H5v4j5vYGTj4/0mCk1GgHIxsZpZWMWFhfwhCVKRkORjY1cWFjvgyT/2V4mkgB5AFcAvFBHcExHcEyXR9h+AAABAHRSTlP///////////////7///////////81/////wH//////////yz//wQB///8Ax//DP///wX//9z////1////SQcH9f8V////Igz///8awv8h//8SZP//d4JZ9/9f6gv/////6vxq////3/o3t6f//6j51r3UPv//Lv/C////cvH//9RS/9QlYEC4nv8r/8sU/f6TRf//wMATtPPJVYH4/5n/quv/////UYjS2////8n/cZblfJP/m6linPrkhJE5////8nHoMo5K62LZgqSd5f/51eSv/+L///9MaP/v/f9B1fb/8v9g5M7k/9bnx758xn3v3NWB9PPwlEvJncn/ZAAAgroDtAAACqtJREFUaN7smXdQG1cex5cVK2l3UASSQBISRQIhiS5QEAjRRbewDTLFNsUGU0wJmLGxaTa4MYDpOfeCG+c2TpxzTVzmxpnLXezk4pa5yTjVl5u5uZmb63d/cW9XElqJXQkS5798/wL98f28X3v79i00/xML+hnwM+CnAUR3FpH/DYv2f60A/5Dr7SnkH4pKU6JfIyAsZTS8/BX5l5DdG9rKQqJT/V8LYPXedhiGS/0dAB4e6bs7PrxaloLLl6SQoszosGUBhh99Afz1Q+ScFB32wLVxY/oGXG+StPtwR1tpmW+m/1IBw5s1DACARzPJNW/zcKWN6W8evFrkvyRAzhN5DIID1nWSq16b7uFG6R0pYUsAxB1trCMCgMOvzjsVwZ0Oh7gHKJsuJGRwCQCvjVyE1A/dA9LL3APi5rB6S4ZguIG8oMyhjTS2gQaZicVimboD01+tj3AHKNyORSmsgPJaUobaaGoQKCspKUmsqclnlbASt5x6f+3dA2muAE2fYRJvKwA+uNBHvh006zeUGF7m3y+4dLagdacHN/HBr449XvHekcgIOoByoI8EWOgj3w6a9ctMqurBgtaaxMT8iwUPKvs5NZ9f9vI6tGLNtjSaCIobSQB4ItUyZUMdG6j9Zf1RT8/e7zbA3SzTzheDVVqYIHgdW7EmUukESC0CXawEAHnyAqCdKHN0belhygwZWP0+dRcvJZqmEypBmS+2Zpsluu9W4gSAOJLmAAh7dRB3K07C6oMXAOGlIAR/37I2Sv/AkmAzGt/6gGWqloYCQH7Bc3lQlaHVy6LHaw6QAZ3t4aUghMk+zDYHthDCQq5SV0BmuG1MgFp3sljdwSIASLw0KMnyg/I/n9pKEA69d8AOyByyuDVdQFHNgr8lBF/qIQ4sqczKSgjFAbKMZCvAj6k1nNy0amQlkSYrAQekrANuE9HzhZtRTGuvMtzuC2aYukUNpnPMLHP16fv4kMkAoKbgqRwzipN33gu4d2/VFB7DmjQrIHWCaMvzYJIbMSlkB4Rfjw5po2mhl1Jmlk/LYEE+i5Ds9OmMBNRorM7fcfzkyeObRkCiHh+xAvaMEnYzRfOTs6iPhmdP0rqylCHqEpiSfQBAWnn6LEHovlgwqBVjRmZd4iff7Dx53DNgFcjTikgCEPdXS2uC9eI5ioq1hwCPlnVQD5mJ7ccM8hFLnp8uuFhT893ZS4OaBHGWkalN/GTHjuOeQIBwbE0EDhj+vXX/WVebWpyEmvv1dkB5O80UmyA/ptFPLJaEPmgtKDh7/3lXgtgniAAEBHh6Wgl4CNB877e2Da7hfM52EILaHgKXQ7OLmtQo0whCEEurMlRPszVRZrHYz2hkViXusPoDwtZDfwCA3GaJDQCPdhb3oeIeexUEdLu/TDDGNBqzxGKxT4JUOiYGKMwI1J9vBwSMeJ1aPw8VXiEteWZfM4rlhdqGTUQXgIdBpgUAYxYwtgqs38hEY3ce91zQppW/3QYAu6TshekNP/jRBRQ7ZwuJNgAwaPFBOCHIz4JAs/D1M28bTgbYAQEjh47gAPM0qfVnHvVhZg0DcVkBPEeBchwAEFlAQcSfzKx4UgmIKrwfAQBYS4WdwJv5HkxbhjVDHBchVGIWgl3Mc4Zv7pH8PQOmTq2HVm/G6iH7Dgfrn33rg8lDiQwJ+PQhdJs0QY4EplxY4xCAp+fIigNQ3Bzq08WwA2BkvMWM3QbbNiIQuQB4mAxdWSQCkylX5L+YWrXJIUe/jITmB2ax+lAumaDoGsMkwYhIh7gCBJpk1VLga7X3a2Hkv7h82WslmbEJBxRuR9EW0h4KCMJqOSpRiUQuARxwphBq6v2Yb7zxRpBZC/3zb3/+19TUyq1bt06N2BAEIKI3CUs4Q9riAEGvOmeOqtTTzgFHB9AckT5Q9lLV31XXEy+UsQ773l177cSJG/uBbpzY/ylRZRwwj+8P9fGwgxBFf17eNIMGIEpW6zx03jGQKpkPnvjg6CULFBBHNf+0X+BanxZ57SucQACUA5+h1r4hEXiqKnkGNQAeZ3tz4Fg2BMSOSRYiIoQhrPj3fx3fY+7eAEH87wD+RMs9moSiUU4EGKmIhygAfA53HBoX8GIgm9iqmOCeFu3tm0rHN7HIE5967icA84W3GsEuGgojjgSqGgu4fAXEriD5E+oS+2FzzmfTyBsBJ9ZbzkU5zYAgj2c4ERYDBHqRSA3FMMYd/dnxURh6K875XLrtq2sR1pMdQZCeUTgQEIFDisDGwdfzRDw1pE6GnFXlgz7JdQakXVtLnCpSwbtA4Qd9KGZuUelJCERHDoEj4nNgBg9hjLPVKmd/9pkEbPPqRWf3bZEEoHOiM3U+txds1GjemWRkAQFmmQTQwXwdg8HjVQSzgxcFwM6QUgHmlQTgVXlDqW90RFMzHoSkWrGAQEiTxgcl0TOAFIuWTxRBTgmwHLxqy+HwhuvnMwuLryRhWIJkWq3nIs6PNBFDpMP9GWqICpDtBoAfKmZKU77sxRHmKE2ot54LIPYUgfRzYdxfGEMFgACgOZcOcL7c+jhbNzPxj0d//xrUQirRZMdW6O0AAYOB8IgMsekiWNymNoBvg/2RWd7w7E5PlTYqr16urToTbEkRB4wvg0cEwIilDADUAD1K+xKYedCh+7mwsEIxzq7MzsiItQAEOo7FncHwVtEA8pKKlXQA/9py530IQbi4dBxiEkQiPs/iL6QOAACkfZP078lffoHAlBLx8T7iwFyBFaBQ0QCmx3bl0AMWjqfO4vK5IAQBD7F0KKOCuoUAQGOmbiILYHWzRkhJQASIwAMUGOZaAGo2DQCqauxV0gOUxWM9FZQEEVcAArB1EF2CIHalZLbJ1V1FzvaELtImRK62jo+bE5uEdzDd+sFWRD3HNoByYFaszRZyKRAiEcPWobT+oARf02TI9qaf29uHSeugCmRRGAjX5h9D6w+FancNu7nxyj06i/lIW6pjK2AuYhesBw8Ad+uH2NX1cxHurtRyi3ehYCvNO9eTAakV3rgUsWwFKDAxYMku/CGobleT+0vBiKZb4PiCYeCVJS9Kgksu1ygsA+Y9znZhD57IR+OWcuuYOzl3pa8Rh1hU36Ow2KtdLh+cKbYPL/FaMy5nYO7JlQuzfUlJSWOgrfD1V7izZ8ffLlYu/eY3IjenaXJg4OZ/7iTzeAwhWD3btT8UWvfB6mXevqcWnR9qCAebtjpGBbkTe/ovw8u4Wk7NLPJ9NTHq7Z0cG+N27ZZn2feTy7m7Lvp4y7Px8RgVeynmRIL+eDNuOYDUve9Cy9KfXPsvrkHYsgihdz6KW+4HitS9W5YOcOtP1UVh+95+a2n27Lf3RfyQj0T+e975zVL83/r4ofuvLpRzoIzeu4Qgtvw684d/R/Pf87strhHvvvMw7Ed9qAt7+P/htoJD5fiGNjuKZwLVWg5N3YhhBzAHqtxZ97qMOlONTm2HThcdzwebCwIMKiquneu2lNw6+FabKhYAI0Ovpa311Iapd4Cgs3Pdui1bSriB4MU7b2pOlqqVLSgBAW44uHXwBjUtKPPhRgdUtUBnwYeTPptDo0AgFGrXyy7qWaB+bPessl57aK147DDI/MM3jahoQQzynEzMbqANJ6PsaTcjHnMsNHQv0eaPLhoYtWAoWAAAdrOQ1ZQbSUMAAAAASUVORK5CYII="
$PNG1 = [System.Convert]::FromBase64String($base64PNG)

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
        # 3 = With API-KEY
    $global:Auth = 2
#endregion

#region Functions
function Authentication{
    try{
        Switch($Auth){
            1{
                # Without asking for an account and password
                if(($null -ne $UserName -and $null -ne $PlainPassword) -and ($UserName -ne '' -and $PlainPassword -ne '')){
                    Write-Host "--- Automatic AUTHENTICATION (account) ---------------------------" -BackgroundColor Green
                    $uri = "$API/auth/login"                                                                                                    # Webservice Methode
                    $response = Invoke-RestMethod -Uri $uri -Method POST -Body @{ email = "$UserName"; password = "$PlainPassword"}             # Call WebService method
                    $global:Token = $response.token                                                                                             # Register the TOKEN
                    $global:headers.Add("authorization","Bearer $Token")                                                                        # Adding the TOKEN into header
                }Else{
                    Write-Host "--- Account and Password not set ! ---------------------------" -BackgroundColor Red
                    Write-Host "--- To use this connection mode, you must configure the account and password in this script." -ForegroundColor Red
                    Break Script
                }
            }
            2{
                # Requests the entry of an account and a password (default) 
                Write-Host "------------------------------ AUTHENTICATION with account entry ---------------------------" -ForegroundColor Green
                $MyAccount = $Null
                $MyAccount = Get-credential -Message "EKARA login account" -ErrorAction Stop                                                    # Request entry of the EKARA Account
                if(($null -ne $MyAccount) -and ($MyAccount.password.Length -gt 0)){
                    $UserName = $MyAccount.GetNetworkCredential().username
                    $PlainPassword = $MyAccount.GetNetworkCredential().Password
                    $uri = "$API/auth/login"
                    $response = Invoke-RestMethod -Uri $uri -Method POST -Body @{ email = "$UserName"; password = "$PlainPassword"}             # Call WebService method
                    $Token = $response.token                                                                                                    # Register the TOKEN
                    $global:headers.Add("authorization","Bearer $Token")
                }Else{
                    Write-Host "--- Account and password not specified ! ---------------------------" -BackgroundColor Red
                    Write-Host "--- To use this connection mode, you must enter Account and password." -ForegroundColor Red
                    Break Script
                }
            }
            3{
                # With API-KEY
                Write-Host "------------------------------ AUTHENTICATION With API-KEY ---------------------------" -ForegroundColor Green
                if(($null -ne $API_KEY) -and ($API_KEY -ne '')){
                    $global:headers.Add("X-API-KEY", $API_KEY)
                }Else{
                    Write-Host "--- API-KEY not specified ! ---------------------------" -BackgroundColor Red
                    Write-Host "--- To use this connection mode, you must configure API-KEY." -ForegroundColor Red
                    Break Script
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
        Break Script
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
    [Console.Window]::ShowWindow($consolePtr, 2) | Out-Null                     #0 hide / 1 normal / 2 réduit 
}

Function List_Clients{
    try{
        #========================== adm-api/clients =============================
        Write-Host "-------------------------------------------------------------" -ForegroundColor green
        Write-Host "------------------- List all customers  ---------------------" -BackgroundColor "White" -ForegroundColor "DarkCyan"
        $uri ="$API/adm-api/clients"
        $clients = Invoke-RestMethod -Uri $uri -Method POST -Headers $headers 
        $count = $clients.count
        Write-Host "--> [$count] customers ---------------------------" -BackgroundColor Green
        Add-Type -AssemblyName System.Windows.Forms
        Add-Type -AssemblyName System.Drawing

        function ListIndexChanged { 
            #$label2.Text = $listbox.SelectedItems.Count
            $okButton.enabled = $True
        }

        $form = New-Object System.Windows.Forms.Form
        $form.Text = 'List all customers'
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
        $label.Text = 'Select the customers to run inventory:'
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
            Write-Host "------------------- Customers selected -------------------" -BackgroundColor "White" -ForegroundColor "DarkCyan"
            $ItemsName = $listBox.SelectedItems.name
            $global:ItemsID = $listBox.SelectedItems.id
            $global:clientId = $ItemsID
            Write-Host "Client name selected :$ItemsName (ID = $clientId)" -ForegroundColor Green
            
            # RUN All requests
            Write-Host "--> Request scenarios list" -ForegroundColor Blue
            $global:scenarios = get_scenarios
            Write-Host "--> Request Plannings list" -ForegroundColor Blue
            $global:plannings = get_plannings
            Write-Host "--> Request Alerts list" -ForegroundColor Blue
            $global:Alerts = get_alerts
            Write-Host "--> Request Users list" -ForegroundColor Blue
            $global:Users = get_users
            Write-Host "--> Request Users Profils" -ForegroundColor Blue
            $global:Roles = get_Profils
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
            Write-Host "--> Request Webhook list" -ForegroundColor Blue
            $global:webhooks = get_Webhook
            Write-Host "--> Request API_KEY list" -ForegroundColor Blue
            $global:apikeys = get_API_KEY
            Write-Host "--> Request SSO Provider list" -ForegroundColor Blue
            $global:privider = get_SSO_provider
            Write-Host "--> Request robots list" -ForegroundColor Blue
            $global:robots = get_robots
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

            Write-Host ("END customers inventory ["+$ItemsName+"]: " + $Result_OK) -BackgroundColor Green
            [System.Windows.Forms.MessageBox]::Show("Client inventory [$ItemsName] : $Result_OK","Resultat",[System.Windows.Forms.MessageBoxButtons]::OKCancel,[System.Windows.Forms.MessageBoxIcon]::Information)

        }else{
            Write-Host "No customer selected !" -ForegroundColor Red
            [System.Windows.Forms.MessageBox]::Show(`
                "------------------------------------`n`r No customer selected !`n`r------------------------------------`n`r",`
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
        #region CREATE XLS File ______________________________________________________________________________________________________   
            # Creation du ficheir XLS ---------------------------------------------------
            Write-Host "------------------------------------------------------------------------" -ForegroundColor White
            Write-Host ("--> Create XLS file") -ForegroundColor Green
            [String]$global:NewFile = 'EKARA_Inventaire_'+$clientName+"_"+$date + '.xlsx'          #  Nom du Nouveau fichier XLS 
            Write-Host "--> File name ["$NewFile"]" -ForegroundColor Blue  
            Write-Host "--> Folder ["$path"]" -ForegroundColor Blue         
            $ExcelPath = "$Path\$NewFile"                                                           # Nom et Chemin du fichier
            $objExcel = new-object -comobject excel.application                                     # Creation du fichier XLS
            $objExcel.Visible = $False                                                              # Affiche Excel

            $FirstRowColor = 20                                                                     # Cellule Couleur 1
            $SecondtRowColor = 2                                                                    # Cellule Couleur 2
            $color = $FirstRowColor
            $VAlignTop = -4160                                                                      # Cellule Vertical Aligne Haut
            $VAlignBottom = -4107                                                                   # Cellule Vertical Aligne bas
            $VAlignCenter = -4108                                                                   # Cellule Vertical Aligne Center
            $HAlignCenter = -4108                                                                   # Cellule Horizontal Aligne Center
            $HAlignLeft = -4131                                                                     # Cellule Horizontal Aligne Gauche
            $HAlignRight = -4152                                                                    # Cellule Horizontal Aligne Droite

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
        #endregion ___________________________________________________________________________________________________________________

        #region RUM_Inventaire _______________________________________________________________________________________________________
            Write-Host "------------------------------------------------------------------------" -ForegroundColor White
            Write-Host "--> Start RUM inventory for Client Name [$clientName], Client id [$clientID]" -ForegroundColor "Green"
            $row = 1
            $Column = 1
            $workbook = $objExcel.Workbooks.add()                                   # Ajout d'une feuille au fichier XLS
            $finalWorkSheet = $workbook.WorkSheets.item(1)                          # Selection du nouvel onglet                    
            $finalWorkSheet.Name = "RUM"                                            # Nom de l'onglet
            
            # Creation des entete de colonnes
            if($Debug -eq $true){Write-Host "Create header" -ForegroundColor Yellow}                       
            $finalWorkSheet.Cells.Item($row,$Column) = "Tracker Name" 
            $Column++
            $finalWorkSheet.Cells.Item($row,$Column) = "Application" 
            $Column++
            $finalWorkSheet.Cells.Item($row,$Column) = "Scope"
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
            $range = $finalWorkSheet.UsedRange                                        # Selectionne les cellules utilisees
            $range.Interior.ColorIndex = 14                                           # Change la couleur des cellule
            $range.Font.Name = "Arial"                                                # Fixe la police
            $range.Font.Size = 18                                                     # Fixe la taille de la police
            $range.Font.Bold = $True                                                  # Met en carracteres gras
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
                $count = 0
                Foreach ($tracker in $trackers)
                {
                    $count++
                    Write-Host ("`r--> Tracker : " + $count + "/" + $trackers.trackerId.count) -NoNewline -ForegroundColor Blue
                    
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
                                                                        Scope = [string]$tracker.scope;`
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
                    $finalWorkSheet.Cells.Item($row,$Column) = [string]$tracker.scope                                     # Add data                                                    
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
                Write-Host "--> Nb Trackers : " $trackers.trackerId.count -ForegroundColor Blue
                $Column = 1
                $row++
                $Range = $finalWorkSheet.Range("A"+$row).Select()
                $finalWorkSheet.Cells.Item($row,$Column) = [string]"Aucun tracker"                                        # Add data
            }
            
            $Range = $finalWorkSheet.Range("A1").Select()
            $finalWorkSheet.application.activewindow.splitcolumn = 1                                                      # Select Second column for freeze the shutters
            $finalWorkSheet.Application.ActiveWindow.SplitRow = 1                                                         # Select First row for freeze the shutters
            $finalWorkSheet.Application.ActiveWindow.FreezePanes = $true                                                  # Freeze the shutters
            Write-Host ""
            Write-Host "--> END RUM inventory for Client Name [$clientName], Client id [$clientID]" -ForegroundColor "Green"
            Write-Host "------------------------------------------------------------------------" -ForegroundColor White

        #endregion ___________________________________________________________________________________________________________________

        #region SCENARIO_Inventaire __________________________________________________________________________________________________
            # ADD nouvelle feuille -----------------------------------------------------------------------------
            Write-Host "--> Start SCENARIOS inventory for Client Name [$clientName], Client id [$clientID]" -ForegroundColor "Green"
            $NewWorkSheet = $workbook.Worksheets.Add()                                                                    # Creation d'un nouvel onglet  
            $finalWorkSheet = $workbook.Worksheets.Item(1)                                                                # Selection du nouvel onglet                    
            $finalWorkSheet.Name = "Scenarios"                                                                            # Nom de l'onglet

            # Formatage des cellules ---------------------------------------------------------------------------
            $row = 1
            $Column = 1
            $color = $FirstRowColor

            # Creation des entete de colonnes
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
            $range = $finalWorkSheet.UsedRange                                        # Selectionne les cellules utilisees
            $range.Interior.ColorIndex = 14                                           # Change la couleur des cellule
            $range.Font.Name = "Arial"                                                # Fixe la police
            $range.Font.Size = 18                                                     # Fixe la taille de la police
            $range.Font.Bold = $True                                                  # Met en carracteres gras
            $range.Font.ColorIndex = 2                                                # Fixe la couleur de la police
            $range.AutoFilter() | Out-Null                                            # Ajout du filtre automatique
            $range.EntireColumn.AutoFit() | Out-Null                                  # Auto ajustement de la taille de la colonne
            $range.VerticalAlignment = -4160                                          # Aligne vertivalement le contenu des cellules
    
            if(($scenarios.id).count -gt 0){
                #========================== scenarios =============================
                Write-Host "--> Nb scenarios : " ($scenarios.id).count -ForegroundColor Blue
                $Result1 = @()
                $global:Analyze_scenarios = @()
                $ListAlert = ""
                $scenario_inactif = 0
                $scenario_actif = 0
                $NB_scenario_retry = 0
                $GT_10_Step = 0
                $NB_scenario_sans_alert = 0
                $NB_scenario_version = 0
                $count = 0

                Foreach ($scenario in $scenarios)
                {
                    $count++
                    Write-Host ("`r--> Scenario : " + $count + "/" + ($scenarios.id).count) -NoNewline -ForegroundColor Blue
                    
                    $uri2 =("$API/adm-api/scenarios/"+$scenario.id+"?clientId="+$clientId)
                    $scenario_detail = Invoke-RestMethod -Uri $uri2 -Method GET -Headers $headers
                    if($scenario_detail.active -eq 0){$Status="Inactif";$scenario_inactif++}else{$Status="Actif";$scenario_actif++}
                    if($scenario_detail.Retry -eq 0){$Retry="Inactif"}else{$Retry="Actif";$NB_scenario_retry++}

                    if($Debug -eq $true){Write-Host "Scenario Name : " $scenario_detail.name -ForegroundColor Yellow}  
                    if($Debug -eq $true){Write-Host "Scenario zoneName : " $scenario_detail.zoneName -ForegroundColor Yellow}
                    if($Debug -eq $true){Write-Host "Scenario applicationName : " $scenario_detail.applicationName -ForegroundColor Yellow}
                    if($Debug -eq $true){Write-Host "Scenario startDate : " $scenario_detail.startDate -ForegroundColor Yellow} 
                    if($Debug -eq $true){Write-Host "Scenario scriptName : " $scenario_detail.scriptName -ForegroundColor Yellow}
                    if($Debug -eq $true){Write-Host "Scenario Retry : " $Retry -ForegroundColor Yellow}
                    if($Debug -eq $true){Write-Host "Parcours scriptID : " $scenario_detail.scriptId -ForegroundColor Yellow}
                    if($Debug -eq $true){Write-Host "Parcours versionID : " $scenario_detail.scriptVersion -ForegroundColor Yellow}                

                    # Search Parcours info -------------------------------
                    $Parcours = get_script($scenario_detail.scriptId)
                    $Parcours_detail = $Parcours.message.scriptversions | Where-Object { $_.versionId -eq $scenario_detail.scriptVersion}
                    if($Debug -eq $true){Write-Host "Parcours version num : " $Parcours_detail.version -ForegroundColor Yellow}

                    # Verifier si le scenario utilise la derniere version du parcours
                    $scenario_detail_version = ""
                    if($Parcours.message.scriptversions.version[0] -ne $Parcours_detail.version){
                        $NB_scenario_version++                  # Compte le nombre de scenario qui n'utilisent pas la derniere version de parcours pour ANALYZE
                        $scenario_detail_version = $detail_version + "Scenario version = " +$Parcours_detail.version + "`r`nLast parcours version = "+ $Parcours.message.scriptversions.version[0]
                        if($Debug -eq $true){Write-Host ("Scenario = " +$Parcours_detail.version + " / parcours = "+ $Parcours.message.scriptversions.version[0]) -ForegroundColor Yellow}
                    }

                    # Search Nb step for parcours version
                    $uri ="$API/script-api/script/"+$scenario_detail.scriptId+"/version/"+$Parcours_detail.versionId+"/steps?clientId=$clientId"
                    $script_info = Invoke-RestMethod -Uri $uri -Method GET -Headers $headers 
                    if($Debug -eq $true){Write-Host "NB Parcours Step : " $script_info.message.steps.count -ForegroundColor Yellow}
                    $script_NB_Step = $script_info.message.steps.count

                    # ANALYZE _____________________________________________________________________________________
                    if(($scenario_detail.active -ne 0) -and ($script_NB_Step -gt 10)){$GT_10_Step ++}  # Compte si le scenario est actif et a plus de 10 etapes pour ANALYZE
                    # _____________________________________________________________________________________________

                    $plugins = $scenario_detail.plugins
                    $parameters = $scenario_detail.parameters

                    if($Debug -eq $true){Write-Host "Scenario plugins : " $plugins.Name[1] -ForegroundColor Yellow}
                    if($Debug -eq $true){Write-Host "Scenario parameters : " (($parameters.Name.count)) -ForegroundColor Yellow}
                    if($Debug -eq $true){Write-Host ("Parcours SLA Perf : MAX : " + $scenario_detail.slaThreshold.performance.max + " `n`rMIN : " + $scenario_detail.slaThreshold.performance.min) -ForegroundColor Yellow}
                    if($Debug -eq $true){Write-Host ("Parcours SLA Dispo : MAX : " + $scenario_detail.slaThreshold.availability.max + " `n`rMIN : " + $scenario_detail.slaThreshold.availability.min) -ForegroundColor Yellow}

                    # Search Planning Name by ID -------------------------------
                    $Planning = Search_planning($scenario_detail.planningId)
                    if($Debug -eq $true){Write-Host "Scenario planning : " $Planning -ForegroundColor Yellow}
                    $Freq = $Plannings | Where-Object {$_.id -eq $scenario_detail.planningId}

                    $regex = [regex] '(?im)X-EKA-PERIOD:PT(.*M)'                      # Regle de l'expression reguliere pour recuperer les frequences de planning
                    $ListFreq = $regex.Matches($Freq.planning).value 
                    $regex = [regex] '(?<=EKA-PERIOD:PT)(.*?)(?=M)'                   # Regle de l'expression reguliere pour recuperer les frequences de planning
                    $ListFreq = $regex.Matches($ListFreq)
                    
                    #$ListFreq.value
                    $Frequence = ""
                    foreach ($value in $ListFreq.value){
                        $Frequence += "["+$value + "mn] " 
                    }

                    if($Debug -eq $true){Write-Host "Scenario Freq. : " $ListFreq.value -ForegroundColor Yellow}

                    Foreach($parameter in $parameters){
                        If($parameter.Name -eq "timeout"){$timeout = $parameter.Value}
                    }
                
                    #Search Alert Name by ID ----------------------------------
                    $ListAlert = ""
                    if($Debug -eq $true){Write-Host "alert count : " $scenario_detail.alerts.id.count -ForegroundColor Yellow}
                    if($scenario_detail.alerts.id.count -gt 0){
                        Foreach($alert_ID in $scenario_detail.alerts.id){
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
                    if($Debug -eq $true){Write-Host "Site : " $scenario_detail.sites -ForegroundColor Yellow}
                    $site_name = ""
                    Foreach($site in $scenario_detail.sites.name){
                        if($Debug -eq $true){Write-Host "Site : " $site -ForegroundColor Yellow}
                        $site_name += $site + " `n`r"
                    }


                    $Result1 += New-Object -TypeName psobject -Property @{Scenario_Name = [string]$scenario_detail.name;`
                                                                        Type = [string]$plugins.Name[1];`
                                                                        Status = [string]$Status;`
                                                                        Zone = [string]$scenario_detail.zoneName;`
                                                                        Site = [string]$site_name;`
                                                                        Retry = [string]$Retry;`
                                                                        Application = [string]$scenario_detail.applicationName;`
                                                                        Parcours_Name = [string]$scenario_detail.scriptName;`
                                                                        Version = $Parcours_detail.version;`
                                                                        Etape = $script_NB_Step;`
                                                                        Planning = [string]$Planning;`
                                                                        Freq = [string]$Frequence;`
                                                                        Time_out = $timeout;`
                                                                        Parameters = ($parameters.Name).count;`
                                                                        SLA_Perf = [string]"MAX : " + $scenario_detail.slaThreshold.performance.max + " `n`rMIN : " + $scenario_detail.slaThreshold.performance.min;`
                                                                        SLA_Dispo = [string]"MAX : " + $scenario_detail.slaThreshold.availability.max + " `n`rMIN : " + $scenario_detail.slaThreshold.availability.min;`
                                                                        Alerte = [string]$ListAlert;`
                                                                        } | Select-Object Scenario_Name,Type,Status,Zone,Site,Retry,Application,Parcours_Name,Version,Nb_Step,Planning,Freq,Time_out,Parameters,SLA_Perf,SLA_Dispo,Alerte
                
                    $Column = 1
                    $row++
                    $Range = $finalWorkSheet.Range("A"+$row).Select()
                    $finalWorkSheet.Cells.Item($row,$Column) = [string]$scenario_detail.name                # Add data
                    $Column++                                                                               # next column
                    $finalWorkSheet.Cells.Item($row,$Column) = [string]$plugins.Name[1]                     # Add data
                    $Column++                                                                               # next column
                    $finalWorkSheet.Cells.Item($row,$Column) = [string]$Status                              # Add data
                    $Column++                                                                               # next column
                    $finalWorkSheet.Cells.Item($row,$Column) = [string]$scenario_detail.zoneName            # Add data
                    $Column++                                                                               # next column
                    $finalWorkSheet.Cells.Item($row,$Column) = [string]$site_name                           # Add data
                    $Column++                                                                               # next column
                    $finalWorkSheet.Cells.Item($row,$Column) = [string]$Retry                               # Add data
                    $Column++                                                                               # next column
                    $finalWorkSheet.Cells.Item($row,$Column) = [string]$scenario_detail.applicationName     # Add data
                    $Column++                                                                               # next column
                    $finalWorkSheet.Cells.Item($row,$Column) = [string]$scenario_detail.scriptName          # Add data
                    $Column++                                                                               # next column
                    $finalWorkSheet.Cells.Item($row,$Column) = $Parcours_detail.version                     # Add data
                    if($scenario_detail_version -ne ""){
                        [void]$finalWorkSheet.Cells.Item($row,$Column).AddComment(""+$scenario_detail_version+"")                   # Add data into comment
                        $finalWorkSheet.Cells.Item($row,$Column).Comment.Shape.TextFrame.Characters().Font.Size = 8                 # Format comment
                        $finalWorkSheet.Cells.Item($row,$Column).Comment.Shape.TextFrame.Characters().Font.Bold = $False            # Format comment
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
                    $finalWorkSheet.Cells.Item($row,$Column) = [string]"MAX : " + $scenario_detail.slaThreshold.performance.max + " `n`rMIN : " + $scenario_detail.slaThreshold.performance.min                     # Add data
                    $Column++                                                                               # next column
                    $finalWorkSheet.Cells.Item($row,$Column) = [string]"MAX : " + $scenario_detail.slaThreshold.availability.max + " `n`rMIN : " + $scenario_detail.slaThreshold.availability.min                     # Add data
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
            }else{
                Write-Host "--> Nb scenario : " $scenarios.count -ForegroundColor Blue
                $Column = 1
                $row++
                $Range = $finalWorkSheet.Range("A"+$row).Select()
                $finalWorkSheet.Cells.Item($row,$Column) = [string]"Aucun scenario"                                        # Add data
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
            
        #region ALERTS_Inventaire ____________________________________________________________________________________________________
            # ADD nouvelle feuille -----------------------------------------------------------------
            Write-Host "--> Start ALERTS inventory for Client Name [$clientName], Client id [$clientID]" -ForegroundColor "Green"
            $NewWorkSheet = $workbook.Worksheets.Add()                                                  # Creation d'un nouvel onglet  
            $finalWorkSheet = $workbook.Worksheets.Item(1)                                              # Selection du nouvel onglet                    
            $finalWorkSheet.Name = "ALERTS"                                                             # Nom de l'onglet

            # Formatage des cellules ---------------------------------------------------------------------------
            $row = 1
            $Column = 1
            $color = $FirstRowColor

            # Creation des entete de colonnes
            if($Debug -eq $true){Write-Host "Create header" -ForegroundColor Yellow}                       
            $finalWorkSheet.Cells.Item($row,$Column) = "Scenario_name" 
            $Column++
            $finalWorkSheet.Cells.Item($row,$Column) = "Scenario_Type" 
            $Column++
            $finalWorkSheet.Cells.Item($row,$Column) = "Scenario_Status"
            $Column++  
            $finalWorkSheet.Cells.Item($row,$Column) = "Alert_name"
            $Column++ 
            $finalWorkSheet.Cells.Item($row,$Column) = "Alert_type"
            $Column++ 
            $finalWorkSheet.Cells.Item($row,$Column) = "Alert_destinataires"
            $Column++
            $finalWorkSheet.Cells.Item($row,$Column) = "Mail"
            $Column++ 
            $finalWorkSheet.Cells.Item($row,$Column) = "SMS"
            $Column++ 
            $finalWorkSheet.Cells.Item($row,$Column) = "Webhook"
            
            # Format first line
            $range = $finalWorkSheet.UsedRange                                        # Selectionne les cellules utilisees
            $range.Interior.ColorIndex = 14                                           # Change la couleur des cellule
            $range.Font.Name = "Arial"                                                # Fixe la police
            $range.Font.Size = 18                                                     # Fixe la taille de la police
            $range.Font.Bold = $True                                                  # Met en carractères gras
            $range.Font.ColorIndex = 2                                                # Fixe la couleur de la police
            $range.AutoFilter() | Out-Null                                            # Ajout du filtre automatique
            $range.EntireColumn.AutoFit() | Out-Null                                  # Auto ajustement de la taille de la colonne
            $range.VerticalAlignment = -4160                                          # Aligne vertivalement le contenu des cellules
            
            if($alerts.id.count -gt 0){
                #========================== ALERTS =============================
                # Call WS : Gets all the alerts of the current tenant
                Write-Host "--> Nb alertes : " $alerts.id.count -ForegroundColor Blue
                $Result1 = @()
                $ListAlert = ""
                $NB_alert = $alerts.id.count
                $count = 0
                if($scenarios.id.count -gt 0){
                    Write-Host "--> Nb scenarios : " $scenarios.id.count -ForegroundColor Blue
                    Foreach ($scenario in $scenarios){
                        # Requette pour rechercher le detail du scenario en cours
                        $count++
                        Write-Host ("`r--> Scenario : " + $count + "/" + $scenarios.id.count) -NoNewline -ForegroundColor Blue

                        $uri2 =("$API/adm-api/scenarios/"+$scenario.id+"?clientId="+$clientId)
                        $scenario_detail = Invoke-RestMethod -Uri $uri2 -Method GET -Headers $headers
                        if($Debug -eq $true){Write-Host "Scenario Name : " $scenario_detail.name -ForegroundColor Yellow}
                        
                        $Column = 1
                        $row++                                                                                  # Add Line
                        $rowstard = $row                                                                        # Memorize the starting row
                        $Range = $finalWorkSheet.Range("A"+$row).Select()
                        $finalWorkSheet.Cells.Item($row,$Column) = [string]$scenario_detail.name                # Add data

                        $plugins = $scenario_detail.plugins 
                        if($Debug -eq $true){Write-Host "Scenario plugins : " $plugins.Name[1] -ForegroundColor Yellow}

                        $Column++                                                                               # next column 2
                        $finalWorkSheet.Cells.Item($row,$Column) = [string]$plugins.Name[1]                     # Add data

                        if($scenario_detail.active -eq 0){$Status="Inactif"}else{$Status="Actif"}

                        $Column++                                                                               # next column 3
                        $finalWorkSheet.Cells.Item($row,$Column) = [string]$Status                              # Add data

                        if(($scenario_detail.alerts.id.count -eq $true) -or ($scenario_detail.alerts.id.count -gt 0)){
                            if($Debug -eq $true){Write-Host "alert count : " $scenario_detail.alerts.id.count -ForegroundColor Yellow}
                            
                            # Inventory all alerts for this scenario
                            Foreach($alert_ID in $scenario_detail.alerts.id){
                                if(($alert_ID -notlike "926") -and ($alert_ID -notlike "") -and ($alert_ID -notlike $Null)){
                                    $alert_detail = $alerts | Where-Object {$_.id -eq $alert_ID}
                                    if($Debug -eq $true){Write-Host "Alert name : " $alert_detail.name -ForegroundColor Yellow}
                                    
                                    $Column = 4                                                                                     # next column 4
                                    $finalWorkSheet.Cells.Item($row,$Column) = [string]$alert_detail.name                           # Add data

                                    # Properties for MAIL alert 
                                    $alert_type = " - "
                                    If($alert_detail.notificationServiceId -eq 2){
                                        $alert_type = "Mail"
                                        $Column++                                                                                   # next column 5
                                        $finalWorkSheet.Cells.Item($row,$Column) = [string]$alert_type                              # Add data
                                        if($Debug -eq $true){Write-Host "Alert type : " $alert_type -ForegroundColor Yellow}

                                        $alert_Mail = $alert_detail.recipients
                                        foreach($recipient in $alert_Mail){
                                            if($Debug -eq $true){Write-Host ("Destinataire lastname : " + $recipient.lastname + " " + $recipient.firstname) -ForegroundColor Yellow}
                                            if($Debug -eq $true){Write-Host "Destinataire mail : " $recipient.email -ForegroundColor Yellow}
                                            if($Debug -eq $true){Write-Host "Destinataire Phone : " $recipient.phone.nationalNumber -ForegroundColor Yellow}
                                            $lastname_firstname = $recipient.lastname + " " + $recipient.firstname
                                            $Column++                                                                               # next column 6
                                            $finalWorkSheet.Cells.Item($row,$Column) = [string]$lastname_firstname                  # Add data
                                            $Column++                                                                               # next column 7
                                            $finalWorkSheet.Cells.Item($row,$Column) = [string]$recipient.email                     # Add data
                                            $Column++                                                                               # next column 8
                                            $finalWorkSheet.Cells.Item($row,$Column) = [string]" - "                                # Add data
                                            $Column++                                                                               # next column 9
                                            $finalWorkSheet.Cells.Item($row,$Column) = [string]" - "                                # Add data
                                            if ($recipient -ne $alert_Mail[-1]){ 
                                                $row++                                                                              # Add Line
                                                $Column = 5                                                                         # Fixed the column number
                                            }
                                        }
                                    }

                                    # Properties for SMS alert
                                    $alert_type = " - "
                                    If($alert_detail.notificationServiceId -eq 1){
                                        $alert_type = "SMS"
                                        $Column++                                                                                   # next column 5
                                        $finalWorkSheet.Cells.Item($row,$Column) = [string]$alert_type                              # Add data
                                        if($Debug -eq $true){Write-Host "Alert type : " $alert_type -ForegroundColor Yellow}

                                        $alert_SMS = $alert_detail.recipients
                                        foreach($recipient in $alert_SMS){
                                            if($Debug -eq $true){Write-Host ("Destinataire lastname : " + $recipient.lastname + " " + $recipient.firstname) -ForegroundColor Yellow}
                                            if($Debug -eq $true){Write-Host "Destinataire mail : " $recipient.email -ForegroundColor Yellow}
                                            if($Debug -eq $true){Write-Host "Destinataire Phone : " $recipient.phone.nationalNumber -ForegroundColor Yellow}
                                            $lastname_firstname = $recipient.lastname + " " + $recipient.firstname
                                            $Column++                                                                               # next column 6
                                            $finalWorkSheet.Cells.Item($row,$Column) = [string]$lastname_firstname                  # Add data
                                            $Column++                                                                               # next column 7
                                            $finalWorkSheet.Cells.Item($row,$Column) = [string]" - "                                # Add data
                                            $Column++                                                                               # next column 8
                                            $finalWorkSheet.Cells.Item($row,$Column) = [string]$recipient.phone.nationalNumber      # Add data
                                            $Column++                                                                               # next column 9
                                            $finalWorkSheet.Cells.Item($row,$Column) = [string]" - "                                # Add data
                                            if ($recipient -ne $alert_Mail[-1]){ 
                                                $row++                                                                              # Add Line
                                                $Column = 5                                                                         # Fixed the column number
                                            }
                                        }
                                    }

                                    # Properties for Webhook alert 
                                    $alert_type = " - "
                                    If($alert_detail.notificationServiceId -eq 4){
                                        $alert_type = "Webhook"
                                        $Column++                                                                                   # next column 5
                                        $finalWorkSheet.Cells.Item($row,$Column) = [string]$alert_type                              # Add data
                                        if($Debug -eq $true){Write-Host "Alert type : " $alert_type -ForegroundColor Yellow}

                                        $alert_Webhook = $alert_detail.webhooks
                                        foreach($recipient in $alert_Webhook){
                                            if($Debug -eq $true){Write-Host "Destinataire mail : " $recipient.email -ForegroundColor Yellow}
                                            if($Debug -eq $true){Write-Host "Destinataire lastname : " $recipient.lastname -ForegroundColor Yellow}
                                            if($Debug -eq $true){Write-Host "Destinataire firstname : " $recipient.firstname -ForegroundColor Yellow}
                                            $webhook_name = $webhooks | Where-Object {$_.id -eq $recipient}
                                            
                                            $lastname_firstname = $recipient.lastname + " " + $recipient.firstname
                                            $Column++                                                                               # next column 6
                                            $finalWorkSheet.Cells.Item($row,$Column) = [string]" - "                                # Add data
                                            $Column++                                                                               # next column 7
                                            $finalWorkSheet.Cells.Item($row,$Column) = [string]" - "                                # Add data
                                            $Column++                                                                               # next column 8
                                            $finalWorkSheet.Cells.Item($row,$Column) = [string]" - "                                # Add data
                                            $Column++                                                                               # next column 9
                                            $finalWorkSheet.Cells.Item($row,$Column) = [string]$webhook_name.name                   # Add data
                                            if ($recipient -ne $alert_Mail[-1]){ 
                                                $row++                                                                              # Add Line
                                                $Column = 5                                                                         # Fixed the column number
                                            }
                                        }
                                    }
                                }
                            }

                            # Mise en page du document
                            $range = $finalWorkSheet.Range(("A{0}" -f $rowstard),("i{0}" -f $row))                                # Selectionne la ligne en cours
                            $range.Select() | Out-Null                                                                            # Select range
                            $range.Interior.ColorIndex = $color                                                                   # Change la couleur des cellule
                            $range.Font.Size = 12                                                                                 # Fixe la taille de la police
                            $range.EntireColumn.AutoFit() | Out-Null                                                              # Auto ajustement de la taille de la colonne
                            $range.VerticalAlignment = -4160                                                                      # Aligne vertivalement le contenu des cellules

                        }else{
                            # Scenario without alert
                            if($Debug -eq $true){Write-Host ("Aucune Alerte") -ForegroundColor Yellow}
                            $Column++                                                                                         # next column
                            $finalWorkSheet.Cells.Item($row,$Column) = [string]" - "                                          # Add data 
                            $Column++                                                                                         # next column
                            $finalWorkSheet.Cells.Item($row,$Column) = [string]" - "                                          # Add data 
                            $Column++                                                                                         # next column
                            $finalWorkSheet.Cells.Item($row,$Column) = [string]" - "                                          # Add data 
                            $Column++                                                                                         # next column
                            $finalWorkSheet.Cells.Item($row,$Column) = [string]" - "                                          # Add data 
                            $Column++                                                                                         # next column 
                            $finalWorkSheet.Cells.Item($row,$Column) = [string]" - "                                          # Add data 
                            $Column++                                                                                         # next column
                            $finalWorkSheet.Cells.Item($row,$Column) = [string]" - "                                          # Add data 
                        }

                        # Mise en page du document
                        $range = $finalWorkSheet.Range(("A{0}" -f $rowstard),("i{0}" -f $row))                                # Selectionne la ligne en cours
                        $range.Select() | Out-Null                                                                            # Select range
                        $range.Interior.ColorIndex = $color                                                                   # Change la couleur des cellule
                        $range.Font.Size = 12                                                                                 # Fixe la taille de la police
                        $range.EntireColumn.AutoFit() | Out-Null                                                              # Auto ajustement de la taille de la colonne
                        $range.VerticalAlignment = -4160                                                                      # Aligne vertivalement le contenu des cellules
                        if($color -EQ $FirstRowColor){$color = $SecondtRowColor} ELSE {$color = $FirstRowColor}               # Change color line
                    }
                }else{
                    Write-Host "--> Nb scenario : " $trackers.count -ForegroundColor Blue
                }
            }else{
                Write-Host "--> Nb alerts : " $alerts.id.count -ForegroundColor Blue
                Write-Host "--> Nb scenario : " $trackers.count -ForegroundColor Blue
                $Column = 1
                $row++
                $Range = $finalWorkSheet.Range("A"+$row).Select()
                $finalWorkSheet.Cells.Item($row,$Column) = [string]"Aucune alerte"                                            # Add data
            } 

            $Range = $finalWorkSheet.Range("A1").Select()
            $finalWorkSheet.application.activewindow.splitcolumn = 1                                                          # Select Second column for freeze the shutters
            $finalWorkSheet.Application.ActiveWindow.SplitRow = 1                                                             # Select First row for freeze the shutters
            $finalWorkSheet.Application.ActiveWindow.FreezePanes = $true                                                      # Freeze the shutters
            Write-Host ""
            Write-Host "--> End ALERTS inventory for Client Name [$clientName], Client id [$clientID]" -ForegroundColor "Green"
            Write-Host "------------------------------------------------------------------------" -ForegroundColor White   
            # _____________________________________________________________________________________________

		    # Show Formated results
		    if($Debug -eq $true){$Result1 | Out-GridView -Title "Alerts Inventory"}                                           # Display results into GridView

        #endregion ___________________________________________________________________________________________________________________                                                                                                

        #region USERS_Inventaire _____________________________________________________________________________________________________
            # ADD nouvelle feuille -----------------------------------------------------------------
            Write-Host "--> Start USERS inventory for Client Name [$clientName], Client id [$clientID]" -ForegroundColor "Green"
            $row = 1
            $Column = 1
            $color = $FirstRowColor
            $NewWorkSheet = $workbook.Worksheets.Add()                                                  # Création d'un nouvel onglet  
            $finalWorkSheet = $workbook.Worksheets.Item(1)                                              # Sélection du nouvel onglet
            $finalWorkSheet.Name = "Utilisateurs"                                                       # Nom de l'onglet

            $finalWorkSheet.Cells.Item($row,$Column) = "Prenom"                                         # Add data
            $Column++                                                                                   # next column
            $finalWorkSheet.Cells.Item($row,$Column) = "Nom"                                            # Add data
            $Column++                                                                                   # next column
            $finalWorkSheet.Cells.Item($row,$Column) = "Profil"                                         # Add data
            $Column++                                                                                   # next column 
            $finalWorkSheet.Cells.Item($row,$Column) = "Status"                                         # Add data
            $Column++                                                                                   # next column                                                                                                
            $finalWorkSheet.Cells.Item($row,$Column) = "Email"                                          # Add data
            $Column++                                                                                   # next column 
            $finalWorkSheet.Cells.Item($row,$Column) = "Phone"                                          # Add data
            $Column++                                                                                   # next column 
            $finalWorkSheet.Cells.Item($row,$Column) = "Time_zone"                                      # Add data
            $Column++                                                                                   # next column 
            $finalWorkSheet.Cells.Item($row,$Column) = "Language"                                       # Add data
            $Column++                                                                                   # next column 
            $finalWorkSheet.Cells.Item($row,$Column) = "Workspace"                                      # Add data
            $Column++                                                                                   # next column 
            $finalWorkSheet.Cells.Item($row,$Column) = "AUTHENTICATION"                                 # Add data
            $Range = $finalWorkSheet.Range("A"+$row).Select()
            
            # Format first line
            $range = $finalWorkSheet.UsedRange                                                          # Sélectionne les cellules utilisées
            $range.Interior.ColorIndex = 14                                                             # Change la couleur des cellule
            $range.Font.Name = "Arial"                                                                  # Fixe la police
            $range.Font.Size = 18                                                                       # Fixe la taille de la police
            $range.Font.Bold = $True                                                                    # Met en carractères gras
            $range.Font.ColorIndex = 2                                                                  # Fixe la couleur de la police
            $range.AutoFilter() | Out-Null                                                              # Ajout du filtre automatique
            $range.EntireColumn.AutoFit() | Out-Null                                                    # Auto ajustement de la taille de la colonne
            $range.VerticalAlignment = -4160                                                            # Aligne vertivalement le contenu des cellules

            #========================== Users =============================
            # Call WS : Gets all the users of the current tenant
            $global:Analyze_users = @()
            Write-Host "--> Nb users : " $users.count -ForegroundColor Blue
            $user_desactive = ($Users.isActive | Where-Object {$_ -eq 0}).count
            if($user_desactive -gt 0){$global:Analyze_users += New-Object -TypeName psobject -Property @{Comment = [string]"Il y a $user_desactive utilisateur(s) desactive(s).";`
                                                                                                            } | Select-Object Comment}

            $count = 0
            $NB_SSO_Provider_Name = 0
            $Result1 = @()
            Foreach ($user in $Users)
            {
                $count++
                Write-Host ("`r--> User : " + $count + "/" + $Users.count) -NoNewline -ForegroundColor Blue
                if($Debug -eq $true){Write-Host "User firstname : " $user.firstname -ForegroundColor Yellow}
                if($Debug -eq $true){Write-Host "User lastname : " $user.lastname -ForegroundColor Yellow}
                if($Debug -eq $true){Write-Host "User Profil : " $user.roleName -ForegroundColor Yellow}
                if($Debug -eq $true){Write-Host "User email : " $user.email -ForegroundColor Yellow}
                if($Debug -eq $true){Write-Host "User phone : " $user.phone.nationalNumber -ForegroundColor Yellow}
                if($Debug -eq $true){Write-Host "User timezone : " $user.timezone.timeZoneLabel -ForegroundColor Yellow}
                if($Debug -eq $true){Write-Host "User language : " $user.language -ForegroundColor Yellow}
                $workspace_name = ""                                                                                      # Reinitialize variable
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

                if($user.isActive -eq $true){$status = "Activé"}else{$status = "Désactivé"}

                $Result1 += New-Object -TypeName psobject -Property @{Prenom = [string]$user.firstname;`
                                                                    Nom = [string]$user.lastname;`
                                                                    Profil = [string]$user.roleName;`
                                                                    Status = [string]$status;`
                                                                    email = [string]$user.email;`
                                                                    phone = [string]$user.phone.nationalNumber;`
                                                                    timezone = [string]$user.timezone.timeZoneLabel;`
                                                                    language = [string]$user.language;`
                                                                    Workspace = [string]$workspace_name;`
                                                                    Authentication = [string]$SSO_Provider_Name;`
                                                                    } | Select-Object First_Name,Last_Name,Profil,Status,email,phone,timezone,language,Workspace,AUTHENTICATION
                
                $Column = 1
                $row++
                $Range = $finalWorkSheet.Range("A"+$row).Select()
                $finalWorkSheet.Cells.Item($row,$Column) = [string]$user.firstname                      # Add data
                $Column++                                                                               # next column
                $finalWorkSheet.Cells.Item($row,$Column) = [string]$user.lastname                       # Add data
                $Column++                                                                               # next column
                $finalWorkSheet.Cells.Item($row,$Column) = [string]$user.roleName                       # Add data
                $Column++                                                                               # next column
                if ($status -eq "Désactivé" ){
                    $finalWorkSheet.Cells.Item($row,$Column).Font.ColorIndex = 3                        # Set font color (3 = red)
                    $finalWorkSheet.Cells.Item($row,$Column) = [string]$status                          # Add data
                    
                }else{
                    #$finalWorkSheet.Cells.Item($row,$Column).Font.ColorIndex = 1                       # Set font color (1 = black)
                    $finalWorkSheet.Cells.Item($row,$Column) = [string]$status                          # Add data
                }
                $Column++                                                                               # next column
                #$finalWorkSheet.Cells.Item($row,$Column).Font.ColorIndex = 1                           # Set font color (1 = black)
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
                $range = $finalWorkSheet.Range(("A{0}" -f $row),("J{0}" -f $row))                                     # Selectionne la ligne en cours
                $range.Select() | Out-Null                                                                            # Select range
                $range.Interior.ColorIndex = $color                                                                   # Change la couleur des cellule
                $range.Font.Size = 12                                                                                 # Fixe la taille de la police
                $range.EntireColumn.AutoFit() | Out-Null                                                              # Auto ajustement de la taille de la colonne
                $range.VerticalAlignment = -4160                                                                      # Aligne vertivalement le contenu des cellules
                                    
                if($color -EQ $FirstRowColor){$color = $SecondtRowColor} ELSE {$color = $FirstRowColor}               # Change color line
            }

            if([int]$Roles.count -gt 0){
                foreach($Role in $Roles){
                    $List_Roles += $Role.name+" `r`n"
                }
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

        #region APPLiCATION_Inventaire _______________________________________________________________________________________________
            # ADD nouvelle feuille -----------------------------------------------------------------
            Write-Host "--> Start Applications inventory for Client Name [$clientName], Client id [$clientID]" -ForegroundColor "Green"
            $row = 1
            $Column = 1
            $color = $FirstRowColor
            $NewWorkSheet = $workbook.Worksheets.Add()                                                  # Creation d'un nouvel onglet  
            $finalWorkSheet = $workbook.Worksheets.Item(1)                                              # Selection du nouvel onglet
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
            $range = $finalWorkSheet.UsedRange                                        # Selectionne les cellules utilisees
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
            $count = 0
            Foreach ($application in $applications)
            {
                $count++
                Write-Host ("`r--> Application : " + $count + "/" + $applications.id.count) -NoNewline -ForegroundColor Blue
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
                $finalWorkSheet.Cells.Item($row,$Column) = [string]"MAX : " + $application.slaThresholds.performance.max + " `n`rMIN : " + $application.slaThresholds.performance.min      # Add data
                $Column++                                                                                             # next column
                $finalWorkSheet.Cells.Item($row,$Column) = [string]"MAX : " + $application.slaThresholds.availability.max + " `n`rMIN : " + $application.slaThresholds.availability.min    # Add data
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

        #region WORKSPACES_Inventaire ________________________________________________________________________________________________
            # ADD nouvelle feuille -----------------------------------------------------------------
            Write-Host "--> Start Workspaces inventory for Client Name [$clientName], Client id [$clientID]" -ForegroundColor "Green"
            $row = 1
            $Column = 1
            $color = $FirstRowColor
            $NewWorkSheet = $workbook.Worksheets.Add()                                                  # Creation d'un nouvel onglet  
            $finalWorkSheet = $workbook.Worksheets.Item(1)                                              # Selection du nouvel onglet
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
            $range = $finalWorkSheet.UsedRange                                        # Selectionne les cellules utilisees
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
            $count=0

            Foreach ($workspace in $workspaces)
            {
                $count++
                Write-Host ("`r--> Workspaces : " + $count + "/" + $workspaces.id.count) -NoNewline -ForegroundColor Blue
                if($Debug -eq $true){Write-Host "Workspace Name : " $workspace.name -ForegroundColor Yellow}
                if($Debug -eq $true){Write-Host "Workspace Description : " $workspace.description -ForegroundColor Yellow}
                
                $workspace_applications = ""                                               # Reinitialize variable
                Foreach($applicationName in $workspace.applications.name){
                    #$app_name = $applications | Where-Object {$_.id -eq $applicationID}
                    if($Debug -eq $true){Write-Host "User Workspace : " $applicationName -ForegroundColor Yellow}
                    $workspace_applications += $applicationName + " `n`r"
                }
                if($Debug -eq $true){Write-Host "Workspace Applications : " $workspace_applications -ForegroundColor Yellow}

                $workspace_users = ""                                                      # Reinitialize variable
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
                $range = $finalWorkSheet.Range(("A{0}" -f $row),("E{0}" -f $row))                                       # Selectionne la ligne en cours
                $range.Select() | Out-Null                                                                              # Select range
                $range.Interior.ColorIndex = $color                                                                     # Change la couleur des cellule
                $range.Font.Size = 12                                                                                   # Fixe la taille de la police
                $range.EntireColumn.AutoFit() | Out-Null                                                                # Auto ajustement de la taille de la colonne
                $range.VerticalAlignment = -4160                                                                        # Aligne vertivalement le contenu des cellules
                                    
                if($color -EQ $FirstRowColor){$color = $SecondtRowColor} ELSE {$color = $FirstRowColor}                 # Change color line
            }

		    # Show Formated results
		    if($Debug -eq $true){$Result1 | Out-GridView -Title "Workspaces Inventory"}                                 # Display results into GridView
        
            $Range = $finalWorkSheet.Range("A1").Select()
            $finalWorkSheet.application.activewindow.splitcolumn = 1                                                    # Select Second column for freeze the shutters
            $finalWorkSheet.Application.ActiveWindow.SplitRow = 1                                                       # Select First row for freeze the shutters
            $finalWorkSheet.Application.ActiveWindow.FreezePanes = $true                                                # Freeze the shutters
            Write-Host ""
            Write-Host "--> END Workspaces inventory for Client Name [$clientName], Client id [$clientID]" -ForegroundColor "Green"
            Write-Host "------------------------------------------------------------------------" -ForegroundColor White
        #endregion ___________________________________________________________________________________________________________________

        #region ZONE_Inventaire ______________________________________________________________________________________________________
            # ADD nouvelle feuille -----------------------------------------------------------------
            Write-Host "--> Start Zones inventory for Client Name [$clientName], Client id [$clientID]" -ForegroundColor "Green"
            $row = 1
            $Column = 1
            $color = $FirstRowColor
            $global:Analyze_Zones = @()
            $NewWorkSheet = $workbook.Worksheets.Add()                                                  # Creation d'un nouvel onglet  
            $finalWorkSheet = $workbook.Worksheets.Item(1)                                              # Selection du nouvel onglet
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
            $range = $finalWorkSheet.UsedRange                                        # Selectionne les cellules utilisées
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
            Write-Host "--> Nb Zones : " $zones.id.count -ForegroundColor Blue
            $Result1 = @()                                                            # Reinitialize variable
            $site_private_list = @()                                                  # Reinitialize variable
            $site_public_list = @()                                                   # Reinitialize variable
            $site_private = "NO"                                                      # Reinitialize variable
            $site_public = "NO"                                                       # Reinitialize variable
            $count = 0

            Foreach ($zone in $zones)
            {
                $count++
                Write-Host ("`r--> Zone : " + $count + "/" + $zones.id.count) -NoNewline -ForegroundColor Blue
                if($Debug -eq $true){Write-Host "Zone Name : " $zone.name -ForegroundColor Yellow}
                if($Debug -eq $true){Write-Host "Zone Description : " $zone.description -ForegroundColor Yellow}
                if($Debug -eq $true){Write-Host "Zone Failstatuscount : " $zone.Failstatuscount -ForegroundColor Yellow}
                if($Debug -eq $true){Write-Host "Zone ScenariosNb : " $zone.ScenariosNb -ForegroundColor Yellow}
                
                $zone_sites = ""                                                      # Reinitialize variable
                $site_type = ""                                                       # Reinitialize variable
                $zone_Application = @()                                               # Reinitialize variable

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
                $range = $finalWorkSheet.Range(("A{0}" -f $row),("G{0}" -f $row))                                       # Selectionne la ligne en cours
                $range.Select() | Out-Null                                                                              # Select range
                $range.Interior.ColorIndex = $color                                                                     # Change la couleur des cellule
                $range.Font.Size = 12                                                                                   # Fixe la taille de la police
                $range.EntireColumn.AutoFit() | Out-Null                                                                # Auto ajustement de la taille de la colonne
                $range.VerticalAlignment = -4160                                                                        # Aligne vertivalement le contenu des cellules
                                    
                if($color -EQ $FirstRowColor){$color = $SecondtRowColor} ELSE {$color = $FirstRowColor}                 # Change color line
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

        #region RAPPORTS_Inventaire __________________________________________________________________________________________________
            # ADD nouvelle feuille -----------------------------------------------------------------
            Write-Host "--> Start Report inventory for Client Name [$clientName], Client id [$clientID]" -ForegroundColor "Green"                                                                                                
            
            # Création du contenu
            $NewWorkSheet = $workbook.Worksheets.Add()                                                  # Creation d'un nouvel onglet  
            $finalWorkSheet = $workbook.Worksheets.Item(1)                                              # Selection du nouvel onglet
            $finalWorkSheet.Name = "Rapports"                                                           # Nom de l'onglet
            
            # Création du premier titre (première ligne)
            $Column = 1
            $row = 1
            $color = $FirstRowColor
            $global:Analyze_Synthese = @()                                                                                         
            $finalWorkSheet.Cells.Item($row,$Column) = "Rapports Partager"                                                                                        
            $range = $finalWorkSheet.Range("A"+$row,"E"+$row)                                           # Defini la zone
            $range.select() | Out-Null                                                                  # Selectionne la zone
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
            $range = $finalWorkSheet.Range("A"+$row,"E"+$row)                                           # Defini la zone
            $range.select() | Out-Null                                                                  # Selectionne la zone
            $range.Interior.ColorIndex = 14                                                             # Change la couleur des cellule
            $range.Font.Name = "Arial"                                                                  # Fixe la police
            $range.Font.Size = 18                                                                       # Fixe la taille de la police
            $range.Font.Bold = $True                                                                    # Met en carractères gras
            $range.Font.ColorIndex = 2                                                                  # Fixe la couleur de la police
            $range.HorizontalAlignment = $xlCenter                                                      # Aligne horizontale le contenu des cellules
            $range.VerticalAlignment = $xlBottom                                                        # Aligne vertivalement le contenu des cellules
            $range.AutoFilter() | Out-Null                                                              # Ajout du filtre automatique                                                                                                

            Write-Host ("--> Nb Share Reports : " + [int]($ShareReports.id).count) -ForegroundColor Blue 
            $count = 0  

            Foreach ($Share in $ShareReports){
                $count++
                Write-Host ("`r--> User : " + $count + "/" + $ShareReports.count) -NoNewline -ForegroundColor Blue
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
                $finalWorkSheet.Cells.Item($row,$Column) = [string](get-date($Share.lastModified) -format "MM/dd/yyyy") # Add data
                $Column++                                                                                               # next column
                $finalWorkSheet.Cells.Item($row,$Column) = [string]$listRecipient                                       # Add data

                # Mise en page du document
                $range = $finalWorkSheet.Range(("A{0}" -f $row),("E{0}" -f $row))                                       # Selectionne la ligne en cours
                $range.Select() | Out-Null                                                                              # Select range
                $range.Interior.ColorIndex = $color                                                                     # Change la couleur des cellule
                $range.Font.Size = 12                                                                                   # Fixe la taille de la police
                $range.EntireColumn.AutoFit() | Out-Null                                                                # Auto ajustement de la taille de la colonne
                $range.VerticalAlignment = -4160                                                                        # Aligne vertivalement le contenu des cellules
                                    
                if($color -EQ $FirstRowColor){$color = $SecondtRowColor} ELSE {$color = $FirstRowColor}                 # Change color line
            }
            
            # Création du second titre (première ligne)
            $Column = 1
            $row = $row + 4
            $color = $FirstRowColor
            $global:Analyze_Synthese = @() 
            $finalWorkSheet.Cells.Item($row,$Column) = "Rapports Publier"
            $range = $finalWorkSheet.Range("A"+$row,"E"+$row)                                           # Defini la zone
            $range.select() | Out-Null                                                                  # Selectionne la zone
            $range.Interior.ColorIndex = 14                                                             # Change la couleur des cellule
            $range.Font.Name = "Arial"                                                                  # Fixe la police
            $range.Font.Size = 18                                                                       # Fixe la taille de la police
            $range.Font.Bold = $True                                                                    # Met en carractères gras
            $range.Font.ColorIndex = 2                                                                  # Fixe la couleur de la police
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
            $range = $finalWorkSheet.Range("A"+$row,"E"+$row)                                           # Defini la zone
            $range.select() | Out-Null                                                                  # Selectionne la zone
            $range.Interior.ColorIndex = 14                                                             # Change la couleur des cellule
            $range.Font.Name = "Arial"                                                                  # Fixe la police
            $range.Font.Size = 18                                                                       # Fixe la taille de la police
            $range.Font.Bold = $True                                                                    # Met en carractères gras
            $range.Font.ColorIndex = 2                                                                  # Fixe la couleur de la police
            $range.HorizontalAlignment = $xlCenter                                                      # Aligne horizontale le contenu des cellules
            $range.VerticalAlignment = $xlBottom                                                        # Aligne vertivalement le contenu des cellules
            #$range.columnwidth = 50                                                                    # Largeur des colonnes
            #$range.AutoFilter() | Out-Null                                                             # Ajout du filtre automatique
            
            Write-Host ""
            Write-Host ("--> Nb Publish Reports : " + [int]($PublishReports.id).count) -ForegroundColor Blue 
            $count = 0
            
            Foreach ($Publish in $PublishReports){
                $count++
                Write-Host ("`r--> User : " + $count + "/" + $PublishReports.count) -NoNewline -ForegroundColor Blue
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
                $finalWorkSheet.Cells.Item($row,$Column) = [string]$Publish.name                                                # Add data
                $Column++                                                                                                       # next column
                $finalWorkSheet.Cells.Item($row,$Column) = [string]$Publish.workbookName                                        # Add data
                $Column++                                                                                                       # next column
                $finalWorkSheet.Cells.Item($row,$Column) = [string]$Publish.dashboardName                                       # Add data
                $Column++                                                                                                       # next column
                if($Publish.lastUpdated -ne $Null){
                    $finalWorkSheet.Cells.Item($row,$Column) = [string](get-date($Publish.lastUpdated) -format "dd/MM/yyyy")    # Add data
                }
                $Column++                                                                                                       # next column
                $finalWorkSheet.Cells.Item($row,$Column) = [string]$listworkspaces                                              # Add data

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

        #region CONSO_Inventaire _____________________________________________________________________________________________________
            # ADD nouvelle feuille -----------------------------------------------------------------
            Write-Host "--> Start Conso for customer Name [$clientName], Client id [$clientID] (From : $DefaultDebut to $DefaultFin)" -ForegroundColor "Green"
            $NewWorkSheet = $workbook.Worksheets.Add()                                                  # Creation d'un nouvel onglet  
            $finalWorkSheet = $workbook.Worksheets.Item(1)                                              # Selection du nouvel onglet     
            $finalWorkSheet.Name = "Conso"                                                              # Nom de l'onglet

            # Creation du titre (premiere ligne)
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
            $range.BorderAround($xlContinuous,$xlThick,$xlColorIndexBlue) | Out-Null
            $range.Interior.ColorIndex = 14                                                             # Change la couleur des cellule

            # Creation du contenu
            $Column = 2
            $row = 3
            $finalWorkSheet.Cells.Item($row,$Column) = "Prenom" 
            $Column++
            $finalWorkSheet.Cells.Item($row,$Column) = "Nom"
            $Column++
            $finalWorkSheet.Cells.Item($row,$Column) = "Nb_Action"

            # Format first line
            $range = $finalWorkSheet.Range(("B"+$row),("D"+$row))                       # Selectionne les cellules
            $range.Interior.ColorIndex = 14                                             # Change la couleur des cellule
            $range.Font.Name = "Arial"                                                  # Fixe la police
            $range.Font.Size = 18                                                       # Fixe la taille de la police
            $range.Font.Bold = $True                                                    # Met en carracteres gras
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
            $count = 0
            Foreach ($action in $actions)
            {
                $count++
                Write-Host ("`r--> Actions : " + $count + "/" + $actions.count) -NoNewline -ForegroundColor Blue
                
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
            
            # Formated aray with unique user
            $Result2 = @()
            $more_500_action = 0
            $count = 0
            foreach ($user in $UserList){
                $count++
                Write-Host ("`r--> User : " + $count + "/" + $UserList.count) -NoNewline -ForegroundColor Blue
                
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
                
                if($nbAction -gt 500){$more_500_action++}          # Compte le nombre d'action pour l'ANALYZE

                # Mise en page du document
                $range = $finalWorkSheet.Range(("B{0}" -f $row),("D{0}" -f $row))                                       # Selectionne la ligne en cours
                $range.Select() | Out-Null                                                                              # Select range
                $range.Interior.ColorIndex = $color                                                                     # Change la couleur des cellule
                $range.Font.Size = 12                                                                                   # Fixe la taille de la police
                $range.EntireColumn.AutoFit() | Out-Null                                                                # Auto ajustement de la taille de la colonne
                $range.VerticalAlignment = -4160                                                                        # Aligne vertivalement le contenu des cellules
                                                                                            
                if($color -EQ $FirstRowColor){$color = $SecondtRowColor} ELSE {$color = $FirstRowColor}                 # Change color line
             
            }
            
            # Show Formated results
            if($Debug -eq $true){$Result2 | Out-GridView -Title "unique user"}   
            
            # ANALYZE _____________________________________________________________________________________
            if($more_500_action -gt 0){$global:Analyze_conso += New-Object -TypeName psobject -Property @{Comment = [string]"Il y a $more_500_action utilisateur(s) qui ont effectue plus de 500 actions en $global:Action_period jours.";`
                                                                                                    } | Select-Object Comment}
            # _____________________________________________________________________________________________

            $Range = $finalWorkSheet.Range("A1").Select()
            $finalWorkSheet.Application.ActiveWindow.SplitRow = 1                                                       # Select First row for freeze the shutters
            $finalWorkSheet.Application.ActiveWindow.FreezePanes = $true                                                # Freeze the shutters
            Write-Host ""
            Write-Host "--> END conso inventory for Client Name [$clientName], Client id [$clientID]" -ForegroundColor "Green"
            Write-Host "------------------------------------------------------------------------" -ForegroundColor White
        #endregion ___________________________________________________________________________________________________________________

        #region STATS_Inventaire _____________________________________________________________________________________________________
            # ADD nouvelle feuille -----------------------------------------------------------------
            Write-Host "--> Start STATS for customer Name [$clientName], Client id [$clientID]" -ForegroundColor "Green"
            # Creation du contenu
            $NewWorkSheet = $workbook.Worksheets.Add()                                                  # Creation d'un nouvel onglet  
            $finalWorkSheet = $workbook.Worksheets.Item(1)                                              # Selection du nouvel onglet
            $finalWorkSheet.Name = "STATS"                                                              # Nom de l'onglet
            
            # Recuperation des information du compte client
            $client = $clients | Where-Object -FilterScript {$_.id -eq $clientID}  
            
            # Creation du premier titre (premiere ligne)
            $Column = 1
            $row = 1
            
            $global:Analyze_Synthese = @()                                                                                         
            $finalWorkSheet.Cells.Item($row,$Column) = "Statistiques globaux du perimetre"                                                                                        
            $range = $finalWorkSheet.Range("A"+$row,"D"+$row)                                           # Defini la zone
            $range.select() | Out-Null                                                                  # Selectionne la zone
            $range.Interior.ColorIndex = 14                                                             # Change la couleur des cellule
            $range.Font.Name = "Arial"                                                                  # Fixe la police
            $range.Font.Size = 18                                                                       # Fixe la taille de la police
            $range.Font.Bold = $True                                                                    # Met en carracteres gras
            $range.Font.ColorIndex = 2                                                                  # Fixe la couleur de la police
            $range.EntireColumn.AutoFit() | Out-Null                                                    # Auto ajustement de la taille de la colonne
            $range.HorizontalAlignment = $xlCenter                                                      # Aligne horizontale le contenu des cellules
            $range.VerticalAlignment = $xlBottom                                                        # Aligne vertivalement le contenu des cellules
            $range.columnwidth = 30                                                                     # Largeur des colonnes
            $range.MergeCells = $True                                                                   # Fusionne les celulles
            
            # Creation des entetes
            $Column = 2
            $row = 3
            $finalWorkSheet.Cells.Item($row,$Column) = "Type de scénario"
            $Column++
            $finalWorkSheet.Cells.Item($row,$Column) = "NB"

            # Mise en forme des entetes
            $range = $finalWorkSheet.Range("B"+$row,"C"+$row)                                           # Defini la zone
            $range.select() | Out-Null                                                                  # Selectionne la zone
            $range.Interior.ColorIndex = 14                                                             # Change la couleur des cellule
            $range.Font.Name = "Arial"                                                                  # Fixe la police
            $range.Font.Size = 18                                                                       # Fixe la taille de la police
            $range.Font.Bold = $True                                                                    # Met en carracteres gras
            $range.Font.ColorIndex = 2                                                                  # Fixe la couleur de la police
            $range.HorizontalAlignment = $xlCenter                                                      # Aligne horizontale le contenu des cellules
            $range.VerticalAlignment = $xlBottom                                                        # Aligne vertivalement le contenu des cellules
            $range.columnwidth = 50                                                                     # Largeur des colonnes
            
            # Creation des 2 colonnes
            $Column = 2
            $row = 4
            $color = $SecondtRowColor

            # Compte le nombre de scenarios
            #$uri ="$API/adm-api/scenarios?clientId=$clientId"
            #$ListScenarios = Invoke-RestMethod -Uri $uri -Method POST -Headers $headers
            $ActiveScenarios = $global:scenarios | where-object {$_.active -eq 1}
            Write-Host "--> Nb Actif scenarios : " $ActiveScenarios.id.count -ForegroundColor Blue                                                                                        
            
            # Compte le nombre de type de scenario actif
            $nb_HTTPR = 0
            $nb_BPL = 0
            $nb_WEB = 0
            $nb_DESKTOP = 0
            $nb_Android = 0
            $nb_IOS = 0
            $nb_API = 0
            $nb_EXPERT = 0
            $nb_other = 0
            $count = 0

            Foreach ($scenario in $ActiveScenarios)
            {
                $count++
                Write-Host ("`r--> Scenario : " + $count + "/" + $ActiveScenarios.id.count) -NoNewline -ForegroundColor Blue
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

            # Creation du contenu des 2 colonnes                                                                                       
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
                $range = $finalWorkSheet.Range(("B"+$row),("C"+$row))                                                       # Selectionne les cellules
                $range.select() | Out-Null                                                                                  # Selectionne la zone
                $range.Font.ColorIndex = 3
            }

            # Mise en forme des colonnes
            $range = $finalWorkSheet.Range("B4","C"+$row)                                               # Defini la zone
            $range.select() | Out-Null                                                                  # Selectionne la zone
            $range.Font.Name = "Arial"                                                                  # Fixe la police
            $range.Font.Size = 12                                                                       # Fixe la taille de la police
            $range.Font.Bold = $False                                                                   # Met en carracteres gras
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
            $range = $finalWorkSheet.Range("B4","C"+($row+1))                                           # Defini la zone
            $range.select() | Out-Null                                                                  # Selectionne la zone
            $range.Borders.Item($xlInsideHorizontal).Weight = $xlThin
            $range.Borders.Item($xlInsideHorizontal).LineStyle = $xlContinuous
            $range.Borders.Item($xlInsideHorizontal).ColorIndex = 11                                                                                         
            
            Write-Host ""
            Write-Host "--> END STATS inventory for customer Name [$clientName], Client id [$clientID]" -ForegroundColor "Green"
            Write-Host "------------------------------------------------------------------------" -ForegroundColor White
        #endregion ___________________________________________________________________________________________________________________

        #region SYNTHESE _____________________________________________________________________________________________________________
            # ADD nouvelle feuille -----------------------------------------------------------------
            Write-Host "--> Start Synthese for Client Name [$clientName], Client id [$clientID]" -ForegroundColor "Green"
            # Creation du contenu
            $NewWorkSheet = $workbook.Worksheets.Add()                                                  # Creation d'un nouvel onglet  
            $finalWorkSheet = $workbook.Worksheets.Item(1)                                              # Selection du nouvel onglet
            $finalWorkSheet.Name = "Synthese"                                                           # Nom de l'onglet
            
            # Recuperation des information du compte client
            $client = $clients | Where-Object -FilterScript {$_.id -eq $clientID}  
            
            # Creation du premier titre (premiere ligne)
            $Column = 1
            $row = 1
            
            $global:Analyze_Synthese = @()                                                                                         
            $finalWorkSheet.Cells.Item($row,$Column) = "Synthese compte"                                                                                        
            $range = $finalWorkSheet.Range("A"+$row,"D"+$row)                                           # Defini la zone
            $range.select() | Out-Null                                                                  # Selectionne la zone
            $range.Interior.ColorIndex = 14                                                             # Change la couleur des cellule
            $range.Font.Name = "Arial"                                                                  # Fixe la police
            $range.Font.Size = 18                                                                       # Fixe la taille de la police
            $range.Font.Bold = $True                                                                    # Met en carracteres gras
            $range.Font.ColorIndex = 2                                                                  # Fixe la couleur de la police
            $range.EntireColumn.AutoFit() | Out-Null                                                    # Auto ajustement de la taille de la colonne
            $range.HorizontalAlignment = $xlCenter                                                      # Aligne horizontale le contenu des cellules
            $range.VerticalAlignment = $xlBottom                                                        # Aligne vertivalement le contenu des cellules
            $range.columnwidth = 30                                                                     # Largeur des colonnes
            $range.MergeCells = $True                                                                   # Fusionne les celulles
            
            # Creation des entetes
            $Column = 2
            $row = 3
            $finalWorkSheet.Cells.Item($row,$Column) = "Themes"
            $Column++
            $finalWorkSheet.Cells.Item($row,$Column) = "Valeurs"

            # Mise en forme des entêtes
            $range = $finalWorkSheet.Range("B"+$row,"C"+$row)                                           # Defini la zone
            $range.select() | Out-Null                                                                  # Selectionne la zone
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
            $finalWorkSheet.Cells.Item($row,$Column) = "Etat"                                                               # Add data
            if($client.isActif -eq 0){$Status_client="Inactif"}else{$Status_client="Actif"}
            $finalWorkSheet.Cells.Item($row,$Column).Interior.ColorIndex = $color 
            $finalWorkSheet.Cells.Item($row,$Column + 1) = $Status_client                                                   # Add data
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
            
            $features = ""
            foreach ($item in $client.features){
                $features += $item +", `r`n"
            }

            $row++
            if($color -EQ $FirstRowColor){$color = $SecondtRowColor} ELSE {$color = $FirstRowColor}                         # Change color line
            $finalWorkSheet.Cells.Item($row,$Column) = "Option"                                                             # Add data
            $finalWorkSheet.Cells.Item($row,$Column).Interior.ColorIndex = $color 
            $finalWorkSheet.Cells.Item($row,$Column + 1) = $features                                                        # Add data
            $finalWorkSheet.Cells.Item($row,$Column + 1).Interior.ColorIndex = $color                                                             
            

            # Mise en forme des colonnes
            $range = $finalWorkSheet.Range("B4","C"+$row)                                               # Defini la zone
            $range.select() | Out-Null                                                                  # Selectionne la zone
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
            $range = $finalWorkSheet.Range("B4","C"+($row+1))                                           # Defini la zone
            $range.select() | Out-Null                                                                  # Selectionne la zone
            $range.Borders.Item($xlInsideHorizontal).Weight = $xlThin
            $range.Borders.Item($xlInsideHorizontal).LineStyle = $xlContinuous
            $range.Borders.Item($xlInsideHorizontal).ColorIndex = 11                                                                                        


            # Création du second titre (première ligne)
            $Column = 1
            $row = $row + 2
            
            $global:Analyze_Synthese = @() 
            $finalWorkSheet.Cells.Item($row,$Column) = "Synthese perimetre"
            $range = $finalWorkSheet.Range("A"+$row,"D"+$row)                                           # Defini la zone
            $range.select() | Out-Null                                                                  # Selectionne la zone
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
            $range = $finalWorkSheet.Range("B"+$row,"C"+$row)                                           # Defini la zone
            $range.select() | Out-Null                                                                  # Selectionne la zone
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

            $NB_rum = $trackers | select-Object -Property scope | Where-Object {$_.scope -eq "RUM"}
            $NB_rum_browser = $trackers | select-Object -Property scope | Where-Object {$_.scope -eq "Browser Agent"}

            # Création du contenu des 2 colonnes                                                                                       
            if($color -EQ $FirstRowColor){$color = $SecondtRowColor} ELSE {$color = $FirstRowColor}
            $finalWorkSheet.Cells.Item($row,$Column) = "Scenario" 
            $finalWorkSheet.Cells.Item($row,$Column).Interior.ColorIndex = $color                                            # Add data
            $finalWorkSheet.Cells.Item($row,$Column + 1) = [int]$scenario_actif                                              # Add data
            $finalWorkSheet.Cells.Item($row,$Column + 1).Interior.ColorIndex = $color
            [void]$finalWorkSheet.Cells.Item($row,$Column + 1).AddComment("Nombre de scénario ACTIF")                        # Add data into comment
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
            if($color -EQ $FirstRowColor){$color = $SecondtRowColor} ELSE {$color = $FirstRowColor}                          # Change color line
            $finalWorkSheet.Cells.Item($row,$Column) = "Profils utilisateurs"                                                # Add data                           
            $finalWorkSheet.Cells.Item($row,$Column).Interior.ColorIndex = $color 
            $finalWorkSheet.Cells.Item($row,$Column + 1) = [int]$Roles.count                                                 # Add data  
            $finalWorkSheet.Cells.Item($row,$Column + 1).Interior.ColorIndex = $color   
            if([int]$Roles.count -gt 0){
                [void]$finalWorkSheet.Cells.Item($row,$Column + 1).AddComment(""+$List_Roles+"")                             # Add data into comment
                $finalWorkSheet.Cells.Item($row,$Column + 1).Comment.Shape.TextFrame.Characters().Font.Size = 8              # Format comment
                $finalWorkSheet.Cells.Item($row,$Column + 1).Comment.Shape.TextFrame.Characters().Font.Bold = $False         # Format comment
            }

            $row++
            if($color -EQ $FirstRowColor){$color = $SecondtRowColor} ELSE {$color = $FirstRowColor}                          # Change color line
            $finalWorkSheet.Cells.Item($row,$Column) = "WebHooks"                                                            # Add data                           
            $finalWorkSheet.Cells.Item($row,$Column).Interior.ColorIndex = $color 
            $finalWorkSheet.Cells.Item($row,$Column + 1) = [int]$webhooks.name.Count                                         # Add data  
            $finalWorkSheet.Cells.Item($row,$Column + 1).Interior.ColorIndex = $color    

            $row++
            if($color -EQ $FirstRowColor){$color = $SecondtRowColor} ELSE {$color = $FirstRowColor}                          # Change color line
            $finalWorkSheet.Cells.Item($row,$Column) = "SSO"                                                                 # Add data                           
            $finalWorkSheet.Cells.Item($row,$Column).Interior.ColorIndex = $color 
            $finalWorkSheet.Cells.Item($row,$Column + 1) = [int]$NB_SSO_Provider_Name                                        # Add data  
            $finalWorkSheet.Cells.Item($row,$Column + 1).Interior.ColorIndex = $color    

            $row++
            if($color -EQ $FirstRowColor){$color = $SecondtRowColor} ELSE {$color = $FirstRowColor}                          # Change color line
            $finalWorkSheet.Cells.Item($row,$Column) = "API-Key"                                                             # Add data                           
            $finalWorkSheet.Cells.Item($row,$Column).Interior.ColorIndex = $color 
            $finalWorkSheet.Cells.Item($row,$Column + 1) = [int]$apikeys.name.count                                          # Add data  
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
                [void]$finalWorkSheet.Cells.Item($row,$Column + 1).AddComment(""+$unic_privat_site_list+"")                  # Add data into comment
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
                [void]$finalWorkSheet.Cells.Item($row,$Column + 1).AddComment(""+$unic_public_site_list+"")                  # Add data into comment
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
            $finalWorkSheet.Cells.Item($row,$Column + 1) = [int]$NB_scenario_retry                                           # Add data
            $finalWorkSheet.Cells.Item($row,$Column + 1).Interior.ColorIndex = $color

            $row++ 
            if($color -EQ $FirstRowColor){$color = $SecondtRowColor} ELSE {$color = $FirstRowColor}                          # Change color line
            $finalWorkSheet.Cells.Item($row,$Column) = "Rum Tracker"                                                         # Add data
            $finalWorkSheet.Cells.Item($row,$Column).Interior.ColorIndex = $color 
            $finalWorkSheet.Cells.Item($row,$Column + 1) = [int]$NB_rum.count                                                 # Add data
            $finalWorkSheet.Cells.Item($row,$Column + 1).Interior.ColorIndex = $color 

            $row++ 
            if($color -EQ $FirstRowColor){$color = $SecondtRowColor} ELSE {$color = $FirstRowColor}                           # Change color line
            $finalWorkSheet.Cells.Item($row,$Column) = "Rum Browser"                                                          # Add data
            $finalWorkSheet.Cells.Item($row,$Column).Interior.ColorIndex = $color 
            $finalWorkSheet.Cells.Item($row,$Column + 1) = [int]$NB_rum_browser.count                                         # Add data
            $finalWorkSheet.Cells.Item($row,$Column + 1).Interior.ColorIndex = $color 

            $row++ 
            if($color -EQ $FirstRowColor){$color = $SecondtRowColor} ELSE {$color = $FirstRowColor}                           # Change color line
            $finalWorkSheet.Cells.Item($row,$Column) = "Raports publies"                                                      # Add data
            $finalWorkSheet.Cells.Item($row,$Column).Interior.ColorIndex = $color 
            $finalWorkSheet.Cells.Item($row,$Column + 1) = [int]$PublishReports.id.count                                      # Add data
            $finalWorkSheet.Cells.Item($row,$Column + 1).Interior.ColorIndex = $color 
            
            $row++
            if($color -EQ $FirstRowColor){$color = $SecondtRowColor} ELSE {$color = $FirstRowColor}                           # Change color line
            $finalWorkSheet.Cells.Item($row,$Column) = "Reports Partages"                                                     # Add data
            $finalWorkSheet.Cells.Item($row,$Column).Interior.ColorIndex = $color 
            $finalWorkSheet.Cells.Item($row,$Column + 1) = [int]$ShareReports.id.count                                        # Add data 
            $finalWorkSheet.Cells.Item($row,$Column + 1).Interior.ColorIndex = $color                                                                                        
            
            # Mise en forme des colonnes B et C
            $range = $finalWorkSheet.Range("B19","C"+$row)                                              # Defini la zone
            $range.select() | Out-Null                                                                  # Selectionne la zone
            $range.Font.Name = "Arial"                                                                  # Fixe la police
            $range.Font.Size = 12                                                                       # Fixe la taille de la police
            $range.Font.Bold = $False                                                                   # Met en carractères gras
            $range.Font.ColorIndex = 55                                                                 # Fixe la couleur de la police
            $range.HorizontalAlignment = $xlCenter                                                      # Aligne horizontale le contenu des cellules
            $range.VerticalAlignment = $xlBottom                                                        # Aligne vertivalement le contenu des cellules
            #$range.EntireColumn.AutoFit() | Out-Null                                                   # Auto ajustement de la taille de la colonne
            $range.WrapText = $false
            $range.Orientation = 0
            $range.ShrinkToFit = $false
            $range.ReadingOrder = $xlContext
            $range.MergeCells = $false

            # Mise en forme des cadres
            $range = $finalWorkSheet.Range("B19","C"+($row+1))                                          # Defini la zone
            $range.select() | Out-Null                                                                  # Selectionne la zone
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
                    $range = $finalWorkSheet.Range(("B"+$row),("C"+$row))                                       # Selectionne les cellules
                    $range.select() | Out-Null                                                                  # Selectionne la zone
                    $range.Font.ColorIndex = 3
                    $row++
                }
            }
            if(($global:Analyze_Zones.comment).count -gt 0){
                foreach($comment in $global:Analyze_Zones.comment){
                    $finalWorkSheet.Cells.Item($row,$Column) = [string]$comment
                    $range = $finalWorkSheet.Range(("B"+$row),("C"+$row))                                       # Selectionne les cellules
                    $range.select() | Out-Null                                                                  # Selectionne la zone
                    $range.Font.ColorIndex = 3
                    $row++
                }
            }
            if(($global:Analyze_conso.comment).count -gt 0){
                foreach($comment in $global:Analyze_conso.comment){
                    $finalWorkSheet.Cells.Item($row,$Column) = [string]$comment
                    $range = $finalWorkSheet.Range(("B"+$row),("C"+$row))                                       # Selectionne les cellules
                    $range.select() | Out-Null                                                                  # Selectionne la zone
                    $range.Font.ColorIndex = 3
                    $row++
                }
            }

            if(($global:Analyze_users.comment).count -gt 0){
                foreach($comment in $global:Analyze_users.comment){
                    $finalWorkSheet.Cells.Item($row,$Column) = [string]$comment
                    $range = $finalWorkSheet.Range(("B"+$row),("C"+$row))                                       # Selectionne les cellules
                    $range.select() | Out-Null                                                                  # Selectionne la zone
                    $range.Font.ColorIndex = 3
                    $row++
                }
            }

            if(($global:shared_data.data.name).count -eq 0){
                $finalWorkSheet.Cells.Item($row,$Column) = [string]"Le donnees partagees ne sont pas utilisees (" + ($shared_data.data.name).count + ")"
                $range = $finalWorkSheet.Range(("B"+$row),("C"+$row))                                           # Selectionne les cellules
                $range.select() | Out-Null                                                                      # Selectionne la zone
                $range.Font.ColorIndex = 3
                $row++
            }

            if(($global:Analyze_Rum.comment).count -gt 0){
                foreach($comment in $global:Analyze_Rum.comment){
                    $finalWorkSheet.Cells.Item($row,$Column) = [string]$comment
                    $range = $finalWorkSheet.Range(("B"+$row),("C"+$row))                                       # Selectionne les cellules
                    $range.select() | Out-Null                                                                  # Selectionne la zone
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
            
            Write-Host ""
            Write-Host "--> END Synthese for Client Name [$clientName], Client id [$clientID]" -ForegroundColor "Green"
            Write-Host "------------------------------------------------------------------------" -ForegroundColor White
        #endregion ___________________________________________________________________________________________________________________
            
        #region HOME _________________________________________________________________________________________________________________
            # ADD nouvelle feuille -----------------------------------------------------------------
            Write-Host "--> Add Home tab for Client Name [$clientName], Client id [$clientID]" -ForegroundColor "Green"
            #Image au format Base64String.(correspond à une image PNG 96*96)
            $base64PNG = "iVBORw0KGgoAAAANSUhEUgAAAGAAAABgCAMAAADVRocKAAADAFBMVEVdWVcsKSY4NDE7NzXtdAE3MzDs7OwQDwtbWFbe3d3f3t/1mCnyk0gtKifu7u41Mi9DPz3c29s6NjMvLCkxLiswLCfh4OA9Ojc0MS7b2dorKyuysLDNzMzudQMmJCLn5+dWU1EXFRPBv78xLithXFhGQ0BCPT3/14ZZVlNMSUc4NDHswKDzkEC8uro2MzNUUE5PTEqKiIfi39traGasq6ssKSbW1NQzLy0nGg3f3t45NTJIRUMgHRswLSpFOTn8lTcsKSZwbWo8OTXk5OTq6uppY1/f3t7e3NzFw8M3Iw2tbR7h2dTzlD8SEQ8wLSvwjTvvfRbh3Nzg398KCQjY19f0lTs0MC0vLCk4NDFAPDnzkig4NDH0kjtiYF2cmplzcW9ILhEsKSYsKSY1MS6VlJPveAvKyMg4MzDe3d3f3t7e3d3e3d2ioKCYYB4uKiddWVc4MzAsKSZfVU84NDHqlCzyj0DtlT2Rj47e3d3QZgIbGRWGhIPe3d3f3t57e3tXMw3e3d3f3t5dRigsKSVbWFZfXFphXVvufRcyLiyMbUnn3dbwgyQ6NjPzlkrT0dJRTUs5NTLf3t7xiTF4dXM6NTNhWlbrbAA2MS/zk0goJSPzk0Lf3t7zk0jibwLuexGWSwPvkzzf3t6NWhuCgH6fZR2lpKQ1Mi/f3t7xijPf3t51TBmlaB3zlEjg39+sfUcvLChcWFddWVbvfRX0lEHUiDc4NDDueQ/Hxsbe3d25t7drZ2b0lzrc3NzzhyXpcwR9URpuSB74tW1dWVf0mSpdWFcsKCb1mClfW1iWlZT0ljRdWVdaVVSvrazyk0i5WgHtdQSgnp3e3d3f3d1dUkRMSEZbTTu4cCzSfDb5v4r0n1H1nzhZVlOGakspKSmSkI7ueAvZ2NiBQASvraxfPBhBODGTj46tqqp9eXj1pkv1rl28YhHRpXrDw8H5v4j5vYGTj4/0mCk1GgHIxsZpZWMWFhfwhCVKRkORjY1cWFjvgyT/2V4mkgB5AFcAvFBHcExHcEyXR9h+AAABAHRSTlP///////////////7///////////81/////wH//////////yz//wQB///8Ax//DP///wX//9z////1////SQcH9f8V////Igz///8awv8h//8SZP//d4JZ9/9f6gv/////6vxq////3/o3t6f//6j51r3UPv//Lv/C////cvH//9RS/9QlYEC4nv8r/8sU/f6TRf//wMATtPPJVYH4/5n/quv/////UYjS2////8n/cZblfJP/m6linPrkhJE5////8nHoMo5K62LZgqSd5f/51eSv/+L///9MaP/v/f9B1fb/8v9g5M7k/9bnx758xn3v3NWB9PPwlEvJncn/ZAAAgroDtAAACqtJREFUaN7smXdQG1cex5cVK2l3UASSQBISRQIhiS5QEAjRRbewDTLFNsUGU0wJmLGxaTa4MYDpOfeCG+c2TpxzTVzmxpnLXezk4pa5yTjVl5u5uZmb63d/cW9XElqJXQkS5798/wL98f28X3v79i00/xML+hnwM+CnAUR3FpH/DYv2f60A/5Dr7SnkH4pKU6JfIyAsZTS8/BX5l5DdG9rKQqJT/V8LYPXedhiGS/0dAB4e6bs7PrxaloLLl6SQoszosGUBhh99Afz1Q+ScFB32wLVxY/oGXG+StPtwR1tpmW+m/1IBw5s1DACARzPJNW/zcKWN6W8evFrkvyRAzhN5DIID1nWSq16b7uFG6R0pYUsAxB1trCMCgMOvzjsVwZ0Oh7gHKJsuJGRwCQCvjVyE1A/dA9LL3APi5rB6S4ZguIG8oMyhjTS2gQaZicVimboD01+tj3AHKNyORSmsgPJaUobaaGoQKCspKUmsqclnlbASt5x6f+3dA2muAE2fYRJvKwA+uNBHvh006zeUGF7m3y+4dLagdacHN/HBr449XvHekcgIOoByoI8EWOgj3w6a9ctMqurBgtaaxMT8iwUPKvs5NZ9f9vI6tGLNtjSaCIobSQB4ItUyZUMdG6j9Zf1RT8/e7zbA3SzTzheDVVqYIHgdW7EmUukESC0CXawEAHnyAqCdKHN0belhygwZWP0+dRcvJZqmEypBmS+2Zpsluu9W4gSAOJLmAAh7dRB3K07C6oMXAOGlIAR/37I2Sv/AkmAzGt/6gGWqloYCQH7Bc3lQlaHVy6LHaw6QAZ3t4aUghMk+zDYHthDCQq5SV0BmuG1MgFp3sljdwSIASLw0KMnyg/I/n9pKEA69d8AOyByyuDVdQFHNgr8lBF/qIQ4sqczKSgjFAbKMZCvAj6k1nNy0amQlkSYrAQekrANuE9HzhZtRTGuvMtzuC2aYukUNpnPMLHP16fv4kMkAoKbgqRwzipN33gu4d2/VFB7DmjQrIHWCaMvzYJIbMSlkB4Rfjw5po2mhl1Jmlk/LYEE+i5Ds9OmMBNRorM7fcfzkyeObRkCiHh+xAvaMEnYzRfOTs6iPhmdP0rqylCHqEpiSfQBAWnn6LEHovlgwqBVjRmZd4iff7Dx53DNgFcjTikgCEPdXS2uC9eI5ioq1hwCPlnVQD5mJ7ccM8hFLnp8uuFhT893ZS4OaBHGWkalN/GTHjuOeQIBwbE0EDhj+vXX/WVebWpyEmvv1dkB5O80UmyA/ptFPLJaEPmgtKDh7/3lXgtgniAAEBHh6Wgl4CNB877e2Da7hfM52EILaHgKXQ7OLmtQo0whCEEurMlRPszVRZrHYz2hkViXusPoDwtZDfwCA3GaJDQCPdhb3oeIeexUEdLu/TDDGNBqzxGKxT4JUOiYGKMwI1J9vBwSMeJ1aPw8VXiEteWZfM4rlhdqGTUQXgIdBpgUAYxYwtgqs38hEY3ce91zQppW/3QYAu6TshekNP/jRBRQ7ZwuJNgAwaPFBOCHIz4JAs/D1M28bTgbYAQEjh47gAPM0qfVnHvVhZg0DcVkBPEeBchwAEFlAQcSfzKx4UgmIKrwfAQBYS4WdwJv5HkxbhjVDHBchVGIWgl3Mc4Zv7pH8PQOmTq2HVm/G6iH7Dgfrn33rg8lDiQwJ+PQhdJs0QY4EplxY4xCAp+fIigNQ3Bzq08WwA2BkvMWM3QbbNiIQuQB4mAxdWSQCkylX5L+YWrXJIUe/jITmB2ax+lAumaDoGsMkwYhIh7gCBJpk1VLga7X3a2Hkv7h82WslmbEJBxRuR9EW0h4KCMJqOSpRiUQuARxwphBq6v2Yb7zxRpBZC/3zb3/+19TUyq1bt06N2BAEIKI3CUs4Q9riAEGvOmeOqtTTzgFHB9AckT5Q9lLV31XXEy+UsQ773l177cSJG/uBbpzY/ylRZRwwj+8P9fGwgxBFf17eNIMGIEpW6zx03jGQKpkPnvjg6CULFBBHNf+0X+BanxZ57SucQACUA5+h1r4hEXiqKnkGNQAeZ3tz4Fg2BMSOSRYiIoQhrPj3fx3fY+7eAEH87wD+RMs9moSiUU4EGKmIhygAfA53HBoX8GIgm9iqmOCeFu3tm0rHN7HIE5967icA84W3GsEuGgojjgSqGgu4fAXEriD5E+oS+2FzzmfTyBsBJ9ZbzkU5zYAgj2c4ERYDBHqRSA3FMMYd/dnxURh6K875XLrtq2sR1pMdQZCeUTgQEIFDisDGwdfzRDw1pE6GnFXlgz7JdQakXVtLnCpSwbtA4Qd9KGZuUelJCERHDoEj4nNgBg9hjLPVKmd/9pkEbPPqRWf3bZEEoHOiM3U+txds1GjemWRkAQFmmQTQwXwdg8HjVQSzgxcFwM6QUgHmlQTgVXlDqW90RFMzHoSkWrGAQEiTxgcl0TOAFIuWTxRBTgmwHLxqy+HwhuvnMwuLryRhWIJkWq3nIs6PNBFDpMP9GWqICpDtBoAfKmZKU77sxRHmKE2ot54LIPYUgfRzYdxfGEMFgACgOZcOcL7c+jhbNzPxj0d//xrUQirRZMdW6O0AAYOB8IgMsekiWNymNoBvg/2RWd7w7E5PlTYqr16urToTbEkRB4wvg0cEwIilDADUAD1K+xKYedCh+7mwsEIxzq7MzsiItQAEOo7FncHwVtEA8pKKlXQA/9py530IQbi4dBxiEkQiPs/iL6QOAACkfZP078lffoHAlBLx8T7iwFyBFaBQ0QCmx3bl0AMWjqfO4vK5IAQBD7F0KKOCuoUAQGOmbiILYHWzRkhJQASIwAMUGOZaAGo2DQCqauxV0gOUxWM9FZQEEVcAArB1EF2CIHalZLbJ1V1FzvaELtImRK62jo+bE5uEdzDd+sFWRD3HNoByYFaszRZyKRAiEcPWobT+oARf02TI9qaf29uHSeugCmRRGAjX5h9D6w+FancNu7nxyj06i/lIW6pjK2AuYhesBw8Ad+uH2NX1cxHurtRyi3ehYCvNO9eTAakV3rgUsWwFKDAxYMku/CGobleT+0vBiKZb4PiCYeCVJS9Kgksu1ygsA+Y9znZhD57IR+OWcuuYOzl3pa8Rh1hU36Ow2KtdLh+cKbYPL/FaMy5nYO7JlQuzfUlJSWOgrfD1V7izZ8ffLlYu/eY3IjenaXJg4OZ/7iTzeAwhWD3btT8UWvfB6mXevqcWnR9qCAebtjpGBbkTe/ovw8u4Wk7NLPJ9NTHq7Z0cG+N27ZZn2feTy7m7Lvp4y7Px8RgVeynmRIL+eDNuOYDUve9Cy9KfXPsvrkHYsgihdz6KW+4HitS9W5YOcOtP1UVh+95+a2n27Lf3RfyQj0T+e975zVL83/r4ofuvLpRzoIzeu4Qgtvw684d/R/Pf87strhHvvvMw7Ed9qAt7+P/htoJD5fiGNjuKZwLVWg5N3YhhBzAHqtxZ97qMOlONTm2HThcdzwebCwIMKiquneu2lNw6+FabKhYAI0Ovpa311Iapd4Cgs3Pdui1bSriB4MU7b2pOlqqVLSgBAW44uHXwBjUtKPPhRgdUtUBnwYeTPptDo0AgFGrXyy7qWaB+bPessl57aK147DDI/MM3jahoQQzynEzMbqANJ6PsaTcjHnMsNHQv0eaPLhoYtWAoWAAAdrOQ1ZQbSUMAAAAASUVORK5CYII="
            $PNG1 = [System.Convert]::FromBase64String($base64PNG)                                      # Convert base64 image to image
            $tempImagePath = [System.IO.Path]::Combine($env:TEMP, "temp_image.png")                     # Create temporary file
            [System.IO.File]::WriteAllBytes($tempImagePath, $PNG1)                                      # Add data to image
            
            # Creation du contenu
            $NewWorkSheet = $workbook.Worksheets.Add()                                                  # Creation d'un nouvel onglet  
            $workbook.WorkSheets.item(1).activate | Out-Null                                            # Sélectionne la feuille
            $finalWorkSheet = $workbook.Worksheets.Item(1)                                              # Selection du nouvel onglet
            $finalWorkSheet.Name = "HOME"                                                               # Nom de l'onglet                                                    
            $finalWorkSheet.Tab.Color = 16711680                                                        # Couleur de l'onglet
            $finalWorkSheet.Shapes.AddPicture($tempImagePath,1,1,20,60,96,96) | Out-Null                # Ajoute l'image dans la feuille

            # Passe en fond blanc
            $range = $finalWorkSheet.Range("A1:Z1").EntireColumn                                        # Définie une zone avec tout le contenu des colonnes
            $range.select() | Out-Null                                                                  # Sélectionne
            $range.Interior.ColorIndex = 2                                                              # Change le couleur de la zone

            # Creation du premier titre (premiere ligne)
            $Column = 1
            $row = 1

            $finalWorkSheet.Cells.Item($row,$Column) = "INVENTAIRE EKARA"                               # Ajoute du texte sur la zone définie                                                                                        
            $range = $finalWorkSheet.Range("A"+$row,"F"+$row)                                           # Defini la zone
            $range.select() | Out-Null                                                                  # Selectionne la zone
            $range.MergeCells = $True                                                                   # Fusionne les celulles
            $range.columnwidth = 30                                                                     # Largeur des colonnes
            $range.EntireColumn.AutoFit() | Out-Null                                                    # Auto ajustement de la taille de la colonne
            $range.HorizontalAlignment = $xlCenter                                                      # Aligne horizontale le contenu des cellules
            $range.VerticalAlignment = $xlBottom                                                        # Aligne vertivalement le contenu des cellules
            $range.Interior.ColorIndex = 14                                                             # Change la couleur des cellule
            $range.Font.Name = "Arial"                                                                  # Fixe la police
            $range.Font.Size = 18                                                                       # Fixe la taille de la police
            $range.Font.Bold = $True                                                                    # Met en carracteres gras
            $range.Font.ColorIndex = 2                                                                  # Fixe la couleur de la police

            $Column = 2
            $row = 5

            $texte = "Inventaire de l'environnement EKARA (V $global:Version)`n`n  Client [$clientName]`n  Date [$date]`n"  # Défini un texte

            $finalWorkSheet.Cells.Item($row, $Column) = $texte                                          # Ajoute du texte sur la zone définie
            $finalWorkSheet.Cells.Item($row, $Column).WrapText = $true
            $range = $finalWorkSheet.Range("B5:E10")                                                    # Defini la zone
            $range.select() | Out-Null  
            $range.MergeCells = $True                                                                   # Fusionne les celulles
            $range.columnwidth = 30                                                                     # Largeur des colonnes
            $range.EntireColumn.AutoFit() | Out-Null                                                    # Auto ajustement de la taille de la colonne
            $range.HorizontalAlignment = $HAlignLeft                                                    # Aligne horizontale le contenu des cellules
            $range.VerticalAlignment = $VAlignTop                                                       # Aligne vertivalement le contenu des cellules
            $range.Interior.ColorIndex = 14                                                             # Change la couleur des cellule
            $range.Font.Name = "Arial"                                                                  # Fixe la police
            $range.Font.Size = 12                                                                       # Fixe la taille de la police
            $range.Font.Bold = $False                                                                   # Met en carracteres gras
            $range.Font.ColorIndex = 2                                                                  # Fixe la couleur de la police (2 = blanc)                                                    

            $Column = 2
            $row = 13

            $cell = $finalWorkSheet.Cells.Item($row, $Column)                                           # Sélectionne un célulle
            # Ajout du texte avec un lien hypertexte
            $finalWorkSheet.Hyperlinks.Add(
                $cell,                                      # La cellule cible
                "https://ekara.ip-label.net",               # L'URL du lien hypertexte
                "",                                         # Pas de SubAddress (pas de lien interne)
                "",                                         # Pas d'info-bulle (ScreenTip)
                "Portail EKARA"                             # Le texte à afficher dans la cellule
            ) | Out-Null

            $Column = 4
            $row = 13

            $cell = $finalWorkSheet.Cells.Item($row, $Column)                                           # Sélectionne un célulle
            # Ajout du texte avec un lien hypertexte
            $finalWorkSheet.Hyperlinks.Add(
                $cell,                                      # La cellule cible
                "https://ip-label.com/fr/",                 # L'URL du lien hypertexte
                "",                                         # Pas de SubAddress (pas de lien interne)
                "",                                         # Pas d'info-bulle (ScreenTip)
                "Portail ip-label"                          # Le texte à afficher dans la cellule
            ) | Out-Null                                                    

            Write-Host ""
            Write-Host "--> END Home tab for Client Name [$clientName], Client id [$clientID]" -ForegroundColor "Green"
            Write-Host "------------------------------------------------------------------------" -ForegroundColor White                                                    
        #endregion ___________________________________________________________________________________________________________________                                                    

        #region RECORD XLS File ______________________________________________________________________________________________________
            # Enregistrement du fichier ------------------------------------------------------------------------
            Write-Host ("--> Saving data Excel file. [" + $Path + "\" + $NewFile +  "]") -ForegroundColor Blue
            $workbook.SaveAs($ExcelPath)                                                                    # Save file
            
            # Supprimer le fichier image temporaire, après insertion dans le fichier XLS
            Remove-Item -Path $tempImagePath                                                                # Supprime le fichier image temporaire             

            # Fermeture du fichier 
            Write-Host ("--> Closing Excel file. [" + $Path + "\" + $NewFile +  "]") -ForegroundColor Blue
            $Range = $workbook.Close()                                                                      # close workbook
            $objExcel.Quit()                                                                                # Quit Excel
            Remove-Variable -Name objExcel                                                                  # Supprime la variable
            [gc]::collect()
            [gc]::WaitForPendingFinalizers()                                                                # Attend que tous les processus soient terminés

            # Ouvre le fichier XLS
            Write-Host "--> Open Excel file." -ForegroundColor Blue
            Invoke-Item -Path "$Path\$NewFile" 
            return $Result_OK
        #endregion ___________________________________________________________________________________________________________________                                                        

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
function get_scenarios(){
    #List all scenarios for customer
    if($Debug -eq $true){Write-Host ("Search Scenarios for client : "+ $clientID) -ForegroundColor Yellow}
    $uri ="$API/adm-api/scenarios?clientId=$clientId"                                                                                   # Construct the URI
    $global:scenarios = Invoke-RestMethod -Uri $uri -Method POST -Headers $headers                                                      # Invoke the REST method to get plannings
    if($Debug -eq $true){Write-Host ("Nb Scenarios funded :" + $scenarios.count) -ForegroundColor Yellow}
    return $scenarios
}

function get_plannings(){
    #List all Plannings for customer
    if($Debug -eq $true){Write-Host ("Search Planning for client : "+ $clientID) -ForegroundColor Yellow}
    $uri1 ="$API/adm-api/plannings?clientId=$clientId"                                                                                  # Construct the URI
    $global:Plannings = Invoke-RestMethod -Uri $uri1 -Method GET -Headers $headers                                                      # Invoke the REST method to get plannings
    if($Debug -eq $true){Write-Host ("Nb Plannings finded :" + $Plannings.count) -ForegroundColor Yellow}
    return $Plannings
}

function get_alerts(){
    #List all Alerts for customer
    if($Debug -eq $true){Write-Host ("Search Alerts for client : "+ $clientID) -ForegroundColor Yellow}
    $uri2 ="$API/adm-api/alerts?clientId=$clientId"                                                                                     # Construct the URI
    $global:Alerts = Invoke-RestMethod -Uri $uri2 -Method GET -Headers $headers                                                         # Invoke the REST method to get plannings
    if($Debug -eq $true){Write-Host ("Nb alerts finded :" + $Alerts.count) -ForegroundColor Yellow}
    return $Alerts
}

function get_users(){
    #List all Users for customer
    if($Debug -eq $true){Write-Host ("Search Users for client : "+ $clientID) -ForegroundColor Yellow}
    $uri3 ="$API/adm-api/client/users?clientId=$clientId"                                                                               # Construct the URI
    $global:Users = Invoke-RestMethod -Uri $uri3 -Method GET -Headers $headers                                                          # Invoke the REST method to get plannings
    if($Debug -eq $true){Write-Host ("Nb users finded :" + $Users.count) -ForegroundColor Yellow}
    return $Users
}

function get_script($id_script){
    #List script for ID for customer
    if($Debug -eq $true){Write-Host ("Search parcours for client : "+ $clientID) -ForegroundColor Yellow}
    $uri4 =($API+"/script-api/script/"+$id_script+"?clientId="+$clientId)                                                               # Construct the URI
    $global:Script = Invoke-RestMethod -Uri $uri4 -Method GET -Headers $headers                                                         # Invoke the REST method to get plannings
    if($Debug -eq $true){Write-Host ("Nb Scripts finded :" + $Script.message.name.count) -ForegroundColor Yellow}
    Return $Script
}

function get_scripts{
    #List all scripts for customer
    if($Debug -eq $true){Write-Host ("Search all parcours for client : "+ $clientID) -ForegroundColor Yellow}
    $uri4 =($API+"/script-api/scripts?clientId="+$clientId)                                                                             # Construct the URI
    $global:Scripts = Invoke-RestMethod -Uri $uri4 -Method POST -Headers $headers                                                       # Invoke the REST method to get plannings
    if($Debug -eq $true){Write-Host ("Nb Scripts finded :" + $Scripts.message.name.Count) -ForegroundColor Yellow}
    Return $Scripts
}

function get_applications(){
    #List all applications for customer
    if($Debug -eq $true){Write-Host ("Search applications for client : "+ $clientID) -ForegroundColor Yellow}
    $uri5 ="$API/adm-api/applications?clientId=$clientId"                                                                               # Construct the URI
    $global:applications = Invoke-RestMethod -Uri $uri5 -Method GET -Headers $headers                                                   # Invoke the REST method to get plannings
    if($Debug -eq $true){Write-Host ("Nb applications finded :" + $applications.count) -ForegroundColor Yellow}
    return $applications
}

function get_zones(){
    #List all zones for customer
    if($Debug -eq $true){Write-Host ("Search zones for client : "+ $clientID) -ForegroundColor Yellow}
    $uri5 ="$API/adm-api/zones?clientId=$clientId"                                                                                      # Construct the URI
    $global:zones = Invoke-RestMethod -Uri $uri5 -Method POST -Headers $headers                                                         # Invoke the REST method to get plannings
    if($Debug -eq $true){Write-Host ("Nb zones finded :" + $zones.count) -ForegroundColor Yellow}
    return $zones
}

function get_sites(){
    #List all sites for customer
    if($Debug -eq $true){Write-Host ("Search zones for client : "+ $clientID) -ForegroundColor Yellow}
    $uri5 ="$API/adm-api/sites?clientId=$clientId"                                                                                      # Construct the URI
    $global:sites = Invoke-RestMethod -Uri $uri5 -Method GET -Headers $headers                                                          # Invoke the REST method to get plannings
    if($Debug -eq $true){Write-Host ("Nb sites finded :" + $sites.count) -ForegroundColor Yellow}
    return $sites
}

function get_workspaces(){
    #List all workspaces for customer
    if($Debug -eq $true){Write-Host ("Search workspaces for client : "+ $clientID) -ForegroundColor Yellow}
    $uri5 ="$API/adm-api/workspaces?clientId=$clientId"                                                                                 # Construct the URI
    $global:workspaces = Invoke-RestMethod -Uri $uri5 -Method GET -Headers $headers                                                     # Invoke the REST method to get plannings
    if($Debug -eq $true){Write-Host ("Nb workspaces finded :" + $workspaces.count) -ForegroundColor Yellow}
    return $workspaces
}

function get_actions(){
    #List all actions for customer
    if($Debug -eq $true){Write-Host ("Search all actions for client : "+ $clientID) -ForegroundColor Yellow}
    $global:DefaultDebut = ((Get-Date).adddays(-$global:Action_period)).tostring("yyyy-MM-dd") + " 00:00:00"                            # Set the start date
    $global:DefaultFin = (Get-Date).tostring("yyyy-MM-dd") + " 00:00:00"                                                                # Set the End Date
    if($Debug -eq $true){Write-Host "Periode de $global:Action_period jours : du $DefaultDebut au $DefaultFin" -ForegroundColor Yellow}
    $uri ="$API/adm-api/actions?startTime=$DefaultDebut&endTime=$DefaultFin&clientId=$clientId"                                         # Construct the URI
    $global:actions = Invoke-RestMethod -Uri $uri -Method POST -Headers $headers                                                        # Invoke the REST method to get plannings
    if($Debug -eq $true){Write-Host ("--> Total Nb actions : " + $actions.count) -ForegroundColor Yellow}
    return $actions
}

function get_Webhook(){
    # List Webhook
    if($Debug -eq $true){Write-Host ("Search all Webhooks for client : "+ $clientID) -ForegroundColor Yellow}
    $uri ="$API/adm-api/webhooks?clientId=$clientId"                                                                                    # Construct the URI
    $global:webhooks = Invoke-RestMethod -Uri $uri -Method GET -Headers $headers                                                        # Invoke the REST method to get plannings
    if($Debug -eq $true){Write-Host ("Nb webhooks finded :" + $webhooks.count) -ForegroundColor Yellow}
    Return $webhooks
}

function get_API_KEY(){
    # List API_KEY
    if($Debug -eq $true){Write-Host ("Search all API_KEY for client : "+ $clientID) -ForegroundColor Yellow}
    $uri ="$API/adm-api/apikeys?clientId=$clientId"                                                                                     # Construct the URI
    $global:apikeys = Invoke-RestMethod -Uri $uri -Method GET -Headers $headers                                                         # Invoke the REST method to get plannings
    if($Debug -eq $true){Write-Host ("Nb API_KEY finded :" + $webhooks.count) -ForegroundColor Yellow}
    Return $apikeys
}


function get_SSO_provider(){
    # List provider
    if($Debug -eq $true){Write-Host ("Search all SSO_provider for client : "+ $clientID) -ForegroundColor Yellow}
    $uri ="$API/adm-api/identityproviders?linked_client=$clientId"                                                                      # Construct the URI
    $global:SSO_Providers = Invoke-RestMethod -Uri $uri -Method GET -Headers $headers                                                   # Invoke the REST method to get plannings
    if($Debug -eq $true){Write-Host ("Nb SSO_Providers finded :" + $SSO_Providers.count) -ForegroundColor Yellow}
    Return $SSO_Providers
}

function get_robots(){
	#List all robots for customer
	if($Debug -eq $true){Write-Host ("Search Robots for client : "+ $clientID) -ForegroundColor Yellow}
	$uri5 ="$API/infra-api/inventories?clientId=$clientId"                                                                              # Construct the URI
	$global:robots = Invoke-RestMethod -Uri $uri5 -Method POST -Headers $headers                                                        # Invoke the REST method to get plannings
	if($Debug -eq $true){Write-Host ("Nb Robots finded : " + $robots.count) -ForegroundColor Yellow}
	return $robots
}

function get_shared_data(){
    #List all shared data for customer
    if($Debug -eq $true){Write-Host ("Search all shared data for client : "+ $clientID) -ForegroundColor Yellow}
    $uri ="$API/adm-api/shareddata?clientId=$clientId"                                                                                  # Construct the URI
    $global:shared_data = Invoke-RestMethod -Uri $uri -Method GET -Headers $headers                                                     # Invoke the REST method to get plannings
    if($Debug -eq $true){Write-Host ("Nb shared_data finded :" + $shared_data.count) -ForegroundColor Yellow}
    return $shared_data
}

function get_rum_tracker(){
    #List all tracker for customer
    if($Debug -eq $true){Write-Host ("Search Trackers for client : "+ $clientID) -ForegroundColor Yellow}
    $uri5 ="$API/rum-restit/trk?clientId=$clientId"                                                                                     # Construct the URI
    $global:trackers = Invoke-RestMethod -Uri $uri5 -Method GET -Headers $headers                                                       # Invoke the REST method to get plannings
    if($Debug -eq $true){Write-Host ("Nb trackers finded : " + $trackers.count) -ForegroundColor Yellow}
    return $trackers
}

function get_RUM_metrics(){
    #List all metrics
    if($Debug -eq $true){Write-Host ("Search All RUM metrics") -ForegroundColor Yellow}
    $uri5 ="$API/rum-restit/metrics?clientId=$clientId"                                                                                 # Construct the URI
    $global:RUMmetrics = Invoke-RestMethod -Uri $uri5 -Method GET -Headers $headers                                                     # Invoke the REST method to get plannings
    if($Debug -eq $true){Write-Host ("Nb RUM Metrics finded :" + $RUMmetrics.count) -ForegroundColor Yellow}
    return $RUMmetrics
}

function get_RUM_overview([int]$trackerID, [int]$metricID){
    #List RUM overview for tracker ID
    if($Debug -eq $true){Write-Host ("List RUM overview for tracker ID : "+ $trackerID) -ForegroundColor Yellow}
    $global:Rum_period = 30
    $DefaultDebut = ((Get-Date).adddays(-$global:Rum_period)).tostring("yyyy-MM-dd") + " 00:00:00"                                      # Set the start date
    $DefaultFin = (Get-Date).tostring("yyyy-MM-dd") + " 00:00:00"                                                                       # Set the End Date
    if($Debug -eq $true){Write-Host "Periode de $global:Rum_period jours : du $DefaultDebut au $DefaultFin" -ForegroundColor Yellow}    
    $uri5 ="$API/rum-restit/trk/$trackerID/results/$metricID/overview?clientId=$clientId&from=$DefaultDebut&to=$DefaultFin"             # Construct the URI
    $global:RUMoverview = Invoke-RestMethod -Uri $uri5 -Method POST -Headers $headers                                                   # Invoke the REST method to get plannings
    if($Debug -eq $true){Write-Host ("Nb RUM overview finded :" + $RUMoverview.count) -ForegroundColor Yellow}
    return $RUMoverview
}

function get_RUM_pageGroup($trackerID){
    #GET /rum-restit/trk/{trackerId}/urlgroups
    if($Debug -eq $true){Write-Host ("List RUM Parge Group for tracker ID : "+ $trackerID) -ForegroundColor Yellow}
    $uri5 ="$API/rum-restit/trk/$trackerID/urlgroups?clientId=$clientId"                                                                # Construct the URI
    $global:RUMpageGroup = Invoke-RestMethod -Uri $uri5 -Method GET -Headers $headers                                                   # Invoke the REST method to get plannings
    if($Debug -eq $true){Write-Host ("Nb RUM overview finded :" + $RUMpageGroup.count) -ForegroundColor Yellow}
    return $RUMpageGroup
}

function get_RUM_business_Dimension($trackerID){
    #POST /rum-restit/customdim/business
    if($Debug -eq $true){Write-Host ("List RUM business Dimension for tracker ID : "+ $trackerID) -ForegroundColor Yellow}
    $uri5 ="$API/rum-restit/customdim/business?clientId=$clientId"                                                                      # Construct the URI
    
    $body = @{
        'trackers' = @($trackerID)
    }
     
    $json = $Body | ConvertTo-Json -Depth 5
    $global:RUMbusinessDimension = Invoke-RestMethod -Uri $uri5 -Method POST -Headers $headers -Body $json -ContentType 'application/json'  # Invoke the REST method to get plannings
    if($Debug -eq $true){Write-Host ("Nb RUM business Dimension finded : " + $RUMbusinessDimension.count) -ForegroundColor Yellow}
    return $RUMbusinessDimension
}

function get_RUM_custom_Dimension($trackerID){
    #POST /rum-restit/customdim/custom
    if($Debug -eq $true){Write-Host ("List RUM custom Dimension for tracker ID : "+ $trackerID) -ForegroundColor Yellow}
    $uri5 ="$API/rum-restit/customdim/custom?clientId=$clientId"                                                                        # Construct the URI
    $body = @{
        'trackers' = @($trackerID)
    }
     
    $json = $Body | ConvertTo-Json -Depth 5
    $global:RUMCustomDimension = Invoke-RestMethod -Uri $uri5 -Method POST -Headers $headers -Body $json -ContentType 'application/json'    # Invoke the REST method to get plannings
    if($Debug -eq $true){Write-Host ("Nb RUM custom Dimension finded : " + $RUMCustomDimension.count) -ForegroundColor Yellow}
    return $RUMCustomDimension
}

function get_RUM_infra_Dimension($trackerID){
    #POST /rum-restit/customdim/infra
    if($Debug -eq $true){Write-Host ("List RUM infra Dimension for tracker ID : "+ $trackerID) -ForegroundColor Yellow}
    $uri5 ="$API/rum-restit/customdim/infra?clientId=$clientId"                                                                         # Construct the URI
    $body = @{
        'trackers' = @($trackerID)
    }
     
    $json = $Body | ConvertTo-Json -Depth 5
    $global:RUMinfraDimension = Invoke-RestMethod -Uri $uri5 -Method POST -Headers $headers -Body $json -ContentType 'application/json' # Invoke the REST method to get plannings
    if($Debug -eq $true){Write-Host ("Nb RUM infra Dimension finded : " + $RUMinfraDimension.count) -ForegroundColor Yellow}
    return $RUMinfraDimension
}

function get_RUM_version_Dimension($trackerID){
    #POST /rum-restit/customdim/version
    if($Debug -eq $true){Write-Host ("List RUM version Dimension for tracker ID : "+ $trackerID) -ForegroundColor Yellow}
    $uri5 ="$API/rum-restit/customdim/version?clientId=$clientId"                                                                       # Construct the URI
    $body = @{
        'trackers' = @($trackerID)
    }
     
    $json = $Body | ConvertTo-Json -Depth 5
    $global:RUMversionDimension = Invoke-RestMethod -Uri $uri5 -Method POST -Headers $headers -Body $json -ContentType 'application/json'   # Invoke the REST method to get plannings
    if($Debug -eq $true){Write-Host ("Nb RUM version Dimension finded : " + $RUMversionDimension.count) -ForegroundColor Yellow}
    return $RUMversionDimension
}

function get_Publish_reports(){
    #GET /adm-api/reports/views
    if($Debug -eq $true){Write-Host ("List All published reports for tracker ID : "+ $trackerID) -ForegroundColor Yellow}
    $uri5 ="$API/adm-api/reports/views?clientId=$clientId"                                                                              # Construct the URI
    $Publish_reports = Invoke-RestMethod -Uri $uri5 -Method GET -Headers $headers                                                       # Invoke the REST method to get plannings
    if($Debug -eq $true){Write-Host ("Nb published reports finded : " + $Publish_reports.count) -ForegroundColor Yellow}
    return $Publish_reports
}

function get_Share_reports(){
    #GET /adm-api/reports/schedules
    if($Debug -eq $true){Write-Host ("List All shared reports for tracker ID : "+ $trackerID) -ForegroundColor Yellow}
    $uri5 ="$API/adm-api/reports/schedules?clientId=$clientId"                                                                          # Construct the URI
    $Share_reports = Invoke-RestMethod -Uri $uri5 -Method GET -Headers $headers                                                         # Invoke the REST method to get plannings
    if($Debug -eq $true){Write-Host ("Nb shared reports finded : " + [int]$Share_reports.count) -ForegroundColor Yellow}
    return $Share_reports
}

function get_Profils(){
    #GET /adm-api/roles
    if($Debug -eq $true){Write-Host ("List All users profils") -ForegroundColor Yellow}
    $uri5 ="$API/adm-api/roles?clientId=$clientId"                                                                                      # Construct the URI
    $Roles = Invoke-RestMethod -Uri $uri5 -Method GET -Headers $headers                                                                 # Invoke the REST method to get plannings
    if($Debug -eq $true){Write-Host ("Nb users profils : " + $Roles.count) -ForegroundColor Yellow}
    return $Roles
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