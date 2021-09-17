#-------------------------------------------------------
# POSITIVO INFORMÁTICA S.A.
# ENGENHARIA INDUSTRIAL - TESTES
#
# Script that integrates the Ragentek RTS program with
# the shopfloor (MII).
#
# Author: Leandro Gustavo Biss Becker
#-------------------------------------------------------
param([parameter(Mandatory=$true)][ValidateSet("Label", "DataCheck", "Antenna", "ESCR_NS", "PSN", "Current", IgnoreCase=$true)]$Global:OperationPhase)

$Global:currentExecutingPath = $MyInvocation.MyCommand.Path | Split-Path -Parent
$currentExecutingScript = $MyInvocation.MyCommand.Path

# -------------------------------------------
# CONFIGURATION
#
$MsgBoxTitle = "RTS - Celulares"
$Global:LogsFolder = "C:\_Celular_Tablets\Logs"
# ------------------------------------------


# Include files
.(Join-Path $currentExecutingPath "ScriptLibrary\HelperFunctions.ps1")
.(Join-Path $currentExecutingPath "ScriptLibrary\Forms.ps1")
.(Join-Path $currentExecutingPath "ScriptLibrary\MII.ps1")

Helpers-CheckPowershellVersion

switch(AskText -msg "Localidade" -ComboBox -InitialText "Curitiba", "Manaus")
{
    0 { $Global:Locality = "CWB"; break }
    1 { $Global:Locality = "MAO"; break }
    default { exit }
}

# Load config file
.(Join-Path $currentExecutingPath "Config.ps1")

# kill any "old" running powershell with window title like ours
Helpers-KillPowerShellByWindowTitle "Positivo Informática S.A."

if ($Host.Name -notlike "*PowerShell ISE*") {
    $host.ui.RawUI.WindowTitle = "$MsgBoxTitle - $Global:OperationPhase"
}

$userConfiguration = Helpers-GetUserConfiguration -configFilePath (Join-Path $currentExecutingPath "userconfig.txt") -doNotAskForDevsNum

try
{
    # Launch the background script that will handle MII activity finalization and log sending.
    Start-Process -FilePath "powershell.exe" -ArgumentList @("-f", "$currentExecutingPath\ScriptLibrary\MIIActivityWorker.ps1", "-Locality", $Global:Locality) -WindowStyle Minimized

    # creates the pipe that will be used to send the SNs to the script that finalizes MII activities
    $pipe = New-Object System.IO.Pipes.NamedPipeClientStream("\\.\pipe\MIIFinalizationNamedPipe")
    $pipe.Connect()
    $finalizationScriptStreamPipe = New-Object System.IO.StreamWriter($pipe)
}
catch
{
    Write-Host "Erro disparando ScriptLibrary\MIIActivityWorker.ps1 ou criando/abrindo pipe`n\\.\pipe\MIIFinalizationNamedPipe para comunicação com o script de fechamento de atividades no MII.`n$($_.ToString())" -ForegroundColor Red
    Start-Sleep -Seconds 10
    exit
}


# Function to be called to generate the response file to RTS
$Global:ResponseGenerationFunction = "GenericWriteResponse"
# Setup the MII suite names and load extra scripts
switch ($Global:OperationPhase)
{
    "Label" {
        .(Join-Path $currentExecutingPath "Label.ps1")
        $Global:MiiSuite = "GRAVA_IMEI"
        $Global:ResponseGenerationFunction = "LabelWriteResponse"
    }
    "DataCheck" {
        .(Join-Path $currentExecutingPath "Label.ps1")
        $Global:MiiSuite = "CHECK_IMEI"
        $Global:ResponseGenerationFunction = "LabelWriteResponse"
    }
    "Antenna" {
        $Global:MiiSuite = "ANTENA"
    }
    "PSN" {
        $Global:MiiSuite = "PSN"
        $Global:ResponseGenerationFunction = "PSNWriteResponse"
    }
    "ESCR_NS" {
        $Global:MiiSuite = "ESCR_NS"
        
    }

    "Current" {
        $Global:MiiSuite = "CORRENTE"
    }
}

# defines jiga's name based on production stage
$Jigas = @("$($Global:MiiSuite)_1", "$($Global:MiiSuite)_2", "$($Global:MiiSuite)_3", "$($Global:MiiSuite)_4", "$($Global:MiiSuite)_5", "$($Global:MiiSuite)_6")

# if not running as admin, relaunch asking for privilege elevation
#Helpers-RelaunchElevatedIfNeeded @("-f", $MyInvocation.MyCommand.Path, "-TestMode", $TestMode)

#
# Gets information about the production order
#
function GetOpInfo($SN)
{
    $res = MII-ConsultaOP -WebServiceURL "$MIIServer/XMII/SOAPRunner/MES/ProduzirHardware/Transaction/WS_CITS_ConsultaOP" -SerialNumbers $SN
    
    if ($res -eq $null) {
        MessageBox -Text "Ordem de produção não encontrada para NS $SN" -Title $MsgBoxTitle -ErrorHappened
        return $null
    }

    Helpers-TimedTextOut "-----------------------------------------------------------------"
    Helpers-TimedTextOut "--               Informações sobre o material                  --"
    Helpers-TimedTextOut "-----------------------------------------------------------------"
    $res | Get-Member -MemberType Property |? { $_.Name -notlike "*Specified*" } |% {
        $prop = $_.Name + (New-Object String -ArgumentList (' ', (19 - $_.Name.Length)))
        Helpers-TimedTextOut "$($prop): $($res.($_.Name))"
    }
    Helpers-TimedTextOut "-----------------------------------------------------------------"

    return $res
}

#
# Writes the general response file with result of the request
#
function GenericWriteResponse($request, $success) 
{
    if ($success) { $response = "PASS" }
    else { $response = "FAIL" }

    Helpers-TimedTextOut "Resposta tipo $response enviada ao RTS."
    
    $ResponseXmlFileFolder = (Join-Path $Global:RTSFolders[$Global:OperationPhase]["RESPONSE"] "$($request.result.serialNumber).xml") 
    # creates as a temp file to rename latter thus avoiding that the RTS try to read an unfinished file
    $response | Out-File ($ResponseXmlFileFolder+".temp") -Encoding ascii
    Rename-Item -Path ($ResponseXmlFileFolder+".temp") -NewName $ResponseXmlFileFolder

    return $success
}

#
# Writes the response for PSN
#
function PSNWriteResponse($request, $success) 
{
    if ($success) { 
        $response = "<result>`r`n" +
            "    <SWVER>$($Global:ConfigurationCsv.SoftwareVersion)</SWVER>`r`n" +
            "</result>"

        Helpers-TimedTextOut "Resposta tipo PASS com versão de software`r`n$($Global:ConfigurationCsv.SoftwareVersion) enviada ao RTS."
    }
    else { 
        Helpers-TimedTextOut "Resposta tipo FAIL enviada ao RTS."
        $response = "FAIL"
    }

    $ResponseXmlFileFolder = (Join-Path $Global:RTSFolders[$Global:OperationPhase]["RESPONSE"] "$($request.result.serialNumber).xml") 
    $response | Out-File $ResponseXmlFileFolder -Encoding ascii

    return $success
}

#
# Function that reads the request generated by RTS program
#
function GetRequestOperation($requestFilePath)
{
    Helpers-TimedTextOut "Lendo arquivo de requisição do RTS." -ForegroundColor Yellow

    #$requestFilePath = Join-Path $Global:RTSFolders[$Global:OperationPhase]["REQUEST"] "SN.xml"
    # read input request XML
    try
    {
        for ($i = 0; $i -lt 5; $i++)
        {
            $request = $null
            $request = [xml] (Get-Content -Path $requestFilePath)
            if ($request) { break }
            Start-Sleep -Milliseconds (500 + $i * 100)
        }
    }
    catch
    {
        Helpers-TimedTextOut "Erro ao abrir arquivo de requisição $requestFilePath para inicio do processo de Label do RTS." -ForegroundColor Red
        return $null
    }

    Helpers-TimedTextOut "Dados da requisição:`r`n`tLinha: $($request.result.line)`r`n`tEstação: $($request.result.station)`r`n`tNúmero de Série: $($request.result.serialNumber)`r`n`tColaborador: $($request.result.Employee)" -ForegroundColor Yellow

    return $request
}

#
# Function that reads the result of the operation done by the RTS program
#
function GetOperationResult($responseFilePath)
{
    Helpers-TimedTextOut "Lendo arquivo de resposta de operação do RTS." -ForegroundColor Yellow

    # read input request XML
    try
    {
        for ($i = 0; $i -lt 100; $i++) {
            if (Test-Path $responseFilePath) {
                break
            }
            Start-Sleep -Seconds 1
        }

        if (-not (Test-Path $responseFilePath)) {
            Helpers-TimedTextOut "Arquivo $responseFilePath não foi criado em 1':40''." -ForegroundColor Red
            return $null
        }

        for ($i = 0; $i -lt 10; $i++)
        {
            $response = $null
            $response = [xml] (Get-Content -Path $responseFilePath -ErrorAction SilentlyContinue)
            if ($response -and $response.result.status ) { break }
            Start-Sleep -Milliseconds (500 + $i * 200)
         }
       
	    if ($response) {
		    Helpers-TimedTextOut "Dados da reposta:`r`n`tLinha: $($response.result.line)`r`n`tEstação: $($response.result.station)`r`n`tNúmero de Série: $($response.result.serialNumber)`r`n`tColaborador: $($response.result.Employee)`r`n`tStatus: $($response.result.status)`r`n`tError code: $($response.result.errorCode)" -ForegroundColor Yellow
	    } else {
            Helpers-TimedTextOut "Não foi possível abrir o arquivo $responseFilePath." -ForegroundColor Red
        }
    }
    catch
    {
        Helpers-TimedTextOut "Erro ao abrir arquivo de resposta $responseFilePath para inicio do processo de Label do RTS." -ForegroundColor Red
        return $null
    }

    return $response
}

#
#
# FILE SYSTEM WATCHER
# Return an array, the first object is the powershell job associated with the event and
# the second object is the IO.FileSystemWatcher .NET object and the third is a System.Threading.EventWaitHandle
# kernel event that can be used to wait when events happens
#
function StartMonitoringFolder($folder, $mutexEventName)
{
    #By BigTeddy 05 September 2011 
 
    #This script uses the .NET FileSystemWatcher class to monitor file events in folder(s). 
    #The advantage of this method over using WMI eventing is that this can monitor sub-folders. 
    #The -Action parameter can contain any valid Powershell commands.  I have just included two for example. 
    #The script can be set to a wildcard filter, and IncludeSubdirectories can be changed to $true. 
    #You need not subscribe to all three types of event.  All three are shown for example. 
    # Version 1.1 
 
    $filter = '*.xml'  # You can enter a wildcard filter here. 
 
    # In the following line, you can change 'IncludeSubdirectories to $true if required.                           
    $fsw = New-Object IO.FileSystemWatcher $folder, $filter -Property @{IncludeSubdirectories = $false; NotifyFilter = [IO.NotifyFilters]'LastWrite'} 

    # Creates the kernel event that will be signaled by the file system watcher event to wake us
    $kernelEvent = New-Object System.Threading.EventWaitHandle($false, [System.Threading.EventResetMode]::ManualReset, $mutexEventName)

    # Here, all three events are registerd.  You need only subscribe to events that you need: 
    $job  = Register-ObjectEvent $fsw Changed -SourceIdentifier $mutexEventName -MessageData $mutexEventName -Action {
        $kernelEvent = $null
        try
        {
            # Open the kernel event and signal it to the main script
            if ([System.Threading.EventWaitHandle]::TryOpenExisting($Event.MessageData, [ref]$kernelEvent)) {
                $kernelEvent.Set() | Out-Null
            }
        }
        finally
        {   
            if ($kernelEvent) {
                $kernelEvent.Dispose()
            }
        }

        $Event.SourceEventArgs.Name # outputs the name of changed file as job result
    }

    return @($job, $fsw, $kernelEvent)
}

function MonitoringFolderTEE($folder, $mutexEventName)
{
    #By BigTeddy 05 September 2011 
 
    #This script uses the .NET FileSystemWatcher class to monitor file events in folder(s). 
    #The advantage of this method over using WMI eventing is that this can monitor sub-folders. 
    #The -Action parameter can contain any valid Powershell commands.  I have just included two for example. 
    #The script can be set to a wildcard filter, and IncludeSubdirectories can be changed to $true. 
    #You need not subscribe to all three types of event.  All three are shown for example. 
    # Version 1.1 
 
    $filter = '*.*'  # You can enter a wildcard filter here. 
 
    # In the following line, you can change 'IncludeSubdirectories to $true if required.                           
    $fsw = New-Object IO.FileSystemWatcher $folder, $filter -Property @{IncludeSubdirectories = $true; NotifyFilter = [IO.NotifyFilters]'LastWrite'} 

    # Creates the kernel event that will be signaled by the file system watcher event to wake us
    $kernelEvent = New-Object System.Threading.EventWaitHandle($false, [System.Threading.EventResetMode]::ManualReset, $mutexEventName)

    # Here, all three events are registerd.  You need only subscribe to events that you need: 
    $job  = Register-ObjectEvent $fsw Changed -SourceIdentifier $mutexEventName -MessageData $mutexEventName -Action {
        $kernelEvent = $null
        try
        {
            # Open the kernel event and signal it to the main script
            if ([System.Threading.EventWaitHandle]::TryOpenExisting($Event.MessageData, [ref]$kernelEvent)) {
                $kernelEvent.Set() | Out-Null
            }
        }
        finally
        {   
            if ($kernelEvent) {
                $kernelEvent.Dispose()
            }
        }

        $Event.SourceEventArgs.Name # outputs the name of changed file as job result
    }

    return @($job, $fsw, $kernelEvent)
}

function RenameFileForced($Path, $NewName)
{
    if (Test-Path $newName) {
        Remove-Item $newName -Force
    }

    Rename-Item -Path $Path -NewName $NewName -Force -ErrorAction SilentlyContinue
}

Write-Host "-----------------------------------------------------------------"
Write-Host "Positivo Informática S.A.`r`nEngenharia Industrial - Testes.`r`nSistema Integração RTS - MII para produção de celulares.`r`nPressione Ctrl+C para encerrar."
Write-Host "-----------------------------------------------------------------"

# Start monitoring the RTS request and operation result (acknowledge) folder
$requestFolderMonitoring = StartMonitoringFolder $Global:RTSFolders[$Global:OperationPhase]["REQUEST"] "RequestFileCreated"
$responseFolderMonitoring = StartMonitoringFolder $Global:RTSFolders[$Global:OperationPhase]["ACK"] "AckFileCreated"
$TeeFolderMonitoring = MonitoringFolderTEE $Global:RTSFolders[$Global:OperationPhase]["TEE"] "MiiFileCreated"


$jiga = AskText -msg "Selecione a jiga" -Color Yellow -ComboBox -InitialText $Jigas
$jiga = $UserConfiguration.StationCode + "_" + $Jigas[$jiga]

# Creates the post-it with the configured model
$model = ""
$model = Get-Content (Join-Path $Global:currentExecutingPath "..\..\StationConfig.txt") -ErrorAction SilentlyContinue -ReadCount 1
if (-not $model) {
    $model = Get-Content (Join-Path $Global:currentExecutingPath "..\StationConfig.txt") -ErrorAction SilentlyContinue -ReadCount 1
}
if (-not $model) {
    $model = Get-Content (Join-Path $Global:currentExecutingPath "..\..\..\StationConfig.txt") -ErrorAction SilentlyContinue -ReadCount 1
}

$postit = $null
if ($model) {
    # show post-it
    $postit = PostIt -text $model -status ($MsgBoxTitle + " - " + $jiga) -Top
    
    # adjust log folder
    $Global:LogsFolder = Join-Path $Global:LogsFolder $model
    
    # get RTS XML config file for the model
    $XMLcfgFileName = Join-Path $((Get-Item $Global:currentExecutingPath).Parent.Parent.FullName) "$("Arquivos_Produtos\$model\" + $model + "_" + $Global:OperationPhase + ".xml")"
    if (-not (Test-Path $XMLcfgFileName)) {
        Write-Host "Não foi possível encontrar arquivo $XMLcfgFileName." -ForegroundColor Red
        exit
    }
    
    # loads XML config file and change the value of Station
    # replace XML's version to 1.0 otherwise powershell can't load version 1.1
    [xml] $XML = Get-Content -Path $XMLcfgFileName |% {$_.Replace('?xml version="1.1"','?xml version="1.0"') }
    $node = $XML.Config.TabCtrl.Items.Item(0).Edit | where {$_.Name -eq "_Station"}
    if ($node -eq $null){
        Write-Host "Não foi possível obter informação da estação do arquivo XML $XMLcfgFileName." -ForegroundColor Red
    } else {
        # updates value of Station
        $node.Value = $jiga
    }
    
    # output the content of the changed XML file without losing indention format
    $StringWriter = New-Object System.IO.StringWriter
    $XmlWriter = New-Object System.XMl.XmlTextWriter $StringWriter
    $xmlWriter.Formatting = "indented"
    $xmlWriter.Indentation = 4
    $XML.WriteContentTo($XmlWriter)
    $XmlWriter.Flush()
    $StringWriter.Flush()
    $outputfile = $StringWriter.ToString()
		
    # change back to XML version 1.1, save XML file as a temporary file and then replace the original
    $outputfile = $outputfile.Replace('?xml version="1.0"','?xml version="1.1"')
    $outputfile | Out-File -FilePath $($XMLcfgFileName + ".tmp") -Encoding ascii
    RenameFileForced -Path $($XMLcfgFileName + ".tmp") -NewName $XMLcfgFileName
}



# Create folder where the CSVs with processed data from logs will be kept.
mkdir (Join-Path $Global:LogsFolder "Processados") -ErrorAction SilentlyContinue | out-null
mkdir (Join-Path $Global:LogsFolder "ProcessadosBkp") -ErrorAction SilentlyContinue | out-null

#@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@#
#@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@#
#--------------- Function added by Gblima for move all logs from SIMBA folder. Used when Statemachine is equal to WaitingRequest -----------#
function moveLog{
    $dateLog = Get-Date -UFormat "%Y_%m_%d"
	if (-not (Test-Path "C:\_Celular_Tablets\App_ODM\SPRD\LogFiles")) { New-Item -Path "C:\_Celular_Tablets\App_ODM\SPRD" -Name "LogFiles" -ItemType "directory" }
	if (-not (Test-Path "C:\_Celular_Tablets\App_ODM\SPRD\Log\$($dateLog)")) { New-Item -Path "C:\_Celular_Tablets\App_ODM\SPRD\Log" -Name "$dateLog" -ItemType "directory" }
    Move-Item -Path "C:\_Celular_Tablets\App_ODM\SPRD\Log\$($dateLog)\DUT1\*.zip" -Destination "C:\_Celular_Tablets\App_ODM\SPRD\LogFiles"
    #Move-Item -Path "C:\_Celular_Tablets\App_ODM\SPRD\Log\$($dateLog)\DUT1\*.txt" -Destination "C:\_Celular_Tablets\App_ODM\SPRD\LogFiles"
}
#@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@#
#@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@#
# keep looking for changes at the RTS folders
$stateMachine = "WaitingRequest"
$stateTEE= "WaitingRequest"
while (-not (Helpers-CheckCtrlC))
{
	# Pump windows forms events
    [System.Windows.Forms.Application]::DoEvents()
	
    if ($stateMachine -eq "WaitingRequest" -and $requestFolderMonitoring[2].WaitOne(1000))
    {
        Helpers-TimedTextOutReset
        moveLog
        $requestFolderMonitoring[2].Reset() | out-null

        # new file detected, process it
        $requestFileName = @(Receive-Job $requestFolderMonitoring[0])[0]
        $requestFileName = Join-Path $Global:RTSFolders[$Global:OperationPhase]["REQUEST"] $requestFileName

        $startDate = Get-Date
        Helpers-TimedTextOut "Processando requisição RTS <=> MII" -ForegroundColor Yellow
        $request = GetRequestOperation $requestFileName
        if (-not $request) {
            continue;
        }

	    # deletes an eventual old result/response files
        #Remove-Item (Join-Path $Global:RTSFolders[$Global:OperationPhase]["ACK"] "$($request.result.serialNumber).xml") -ErrorAction SilentlyContinue
        #Remove-Item (Join-Path $Global:RTSFolders[$Global:OperationPhase]["RESPONSE"] "$($request.result.serialNumber).xml") -ErrorAction SilentlyContinue


        $success = $true
          # Only label and data check operations needs info about PO
        if ("ESCR_NS" -contains $Global:OperationPhase) {
            Start-Sleep -Milliseconds 500
            # Post message to the script that handles MII in background
            $finalizationScriptStreamPipe.WriteLine("$($request.result.serialNumber)#$startDate#$Global:MiiSuite#OPEN")
            $finalizationScriptStreamPipe.Flush()
            $stateMachine = "WaitingResult"
            Helpers-TimedTextOut "Aguardando resposta do RTS operação $Global:OperationPhase..." -ForegroundColor Yellow  
        }
            
        # to start automation tool
        out-file   -filepath C:\_Celular_Tablets\SFC\scripts\start\start.txt -Force
        
    }

    Start-Sleep -Milliseconds 500


    if ($stateMachine -eq "WaitingResult" -and $responseFolderMonitoring[2].WaitOne(1000))
    {
        $responseFolderMonitoring[2].Reset() | Out-Null

        Receive-Job $responseFolderMonitoring[0] | Out-Null

        $resultFileName = Join-Path $Global:RTSFolders[$Global:OperationPhase]["ACK"] "$($request.result.serialNumber).xml"
        $responseFileName = Join-Path $Global:RTSFolders[$Global:OperationPhase]["RESPONSE"] "$($request.result.serialNumber).xml"
        $requestFileName = Join-Path $Global:RTSFolders[$Global:OperationPhase]["REQUEST"] "$($request.result.serialNumber).xml"

        $stateMachine = "WaitingRequest"
        $stateTEE= "WaitingResult"
        $response = GetOperationResult $resultFileName
        if (-not $response -or $response.result.status -ne "Passed") {
            Helpers-TimedTextOut "Operação $Global:OperationPhase falhou" -ForegroundColor Red

            # renames the files adding to then the .fail extension to allow us analyze the problem
            RenameFileForced -Path $resultFileName -NewName ($resultFileName +".fail") 
            RenameFileForced -Path $responseFileName -NewName ($responseFileName +".fail") 
            RenameFileForced -Path $requestFileName -NewName ($requestFileName +".fail")
            # remove failed files from MII folders ->> GBLIMA
            Remove-Item "C:\_Celular_Tablets\SFC\Request\*.*" | Where-Object { ! $_.PSIsContainer }
            Remove-Item "C:\_Celular_Tablets\SFC\Response\*.*" | Where-Object { ! $_.PSIsContainer }
            Remove-Item "C:\_Celular_Tablets\SFC\Result\*.*" | Where-Object { ! $_.PSIsContainer }
            #Remove-Item "C:\_Celular_Tablets\SFC\Temp3\*.*" | Where-Object { ! $_.PSIsContainer }

            if (-not $response) {
                MessageBox -Text "FALHA RTS" -ErrorHappened -Title $MsgBoxTitle -StatusText "Tecle ENTER ou clique OK."
            } else {
                MessageBox -Text "NS $($response.result.serialNumber)`r`nFALHA" -ErrorHappened -Title $MsgBoxTitle -StatusText "Tecle ENTER ou clique OK."
            }
        }
        else
        {
            $finishDate = Get-Date

            # Post message to the script that handles MII in background
            Helpers-TimedTextOut "Operação $Global:OperationPhase finalizada.`nEnviando NSs para finalização das atividades do MII em segundo plano." -ForegroundColor Yellow
            $finalizationScriptStreamPipe.WriteLine("$($request.result.serialNumber)#$finishDate#$Global:MiiSuite#CLOSE")
            $finalizationScriptStreamPipe.Flush()

            $form = MessageBox -Text "NS $($response.result.serialNumber)`r`nSUCESSO" -Title $MsgBoxTitle -ForegroundColor Green -Modeless
            Start-Sleep -Seconds 2
            $form.Dispose()

            # deletes the files to avoid reusing then here in the end of process
            Remove-Item $resultFileName -ErrorAction SilentlyContinue -Force
            Remove-Item $responseFileName -ErrorAction SilentlyContinue -Force
            Remove-Item $requestFileName -ErrorAction SilentlyContinue -Force
        }
    }

    if ($enableTEEKey -eq $true -and $stateTEE -eq "WaitingResult" -and $TeeFolderMonitoring[2].WaitOne(1000))
    {
        Helpers-TimedTextOutReset

        $TeeFolderMonitoring[2].Reset() | Out-Null

        Receive-Job $TeeFolderMonitoring[0] | Out-Null

        $stateTEE= "WaitingRequest"

  if ("Label", "DataCheck" -contains $Global:OperationPhase) # Only label and data check operations needs info about PO
        {
     
   Helpers-TimedTextOut "Efetuando upload das chaves TEE para o servidor FTP..." -ForegroundColor Cyan

    #we specify the directory where all files that we want to upload  
     $Dir="C:/_Celular_Tablets/TEE/$model/"
     
    # $file_count = [System.IO.Directory]::GetFiles("$Dir").Count


     #if ($file_count -eq 4){    
     

      try{  
 
      #ftp server 
      $ftp = "ftp://10.70.120.206/TEE_Key/$model/" 
      $user = "userftpd" 
      $pass = "A524INgMXV"  
 
      $webclient = New-Object System.Net.WebClient 
 
      $webclient.Credentials = New-Object System.Net.NetworkCredential($user,$pass)  
 
      #list every sql server trace file 
      foreach($item in (dir $Dir "*.*")){ 
      "Uploading TEE File $item..." 
      $uri = New-Object System.Uri($ftp+$item.Name) 
      $webclient.UploadFile($uri, $item.FullName)
            } 

         }
           catch
         {

         Helpers-TimedTextOut "Falha durante o upload." -ForegroundColor Red
         New-Item -Path (Join-Path $currentExecutingPath "..\..\SFC\Result") -Name "$($response.result.serialNumber).xml" -ItemType file -ErrorAction SilentlyContinue
         Set-Content -Path (Join-Path $currentExecutingPath "..\..\SFC\Result\$($response.result.serialNumber).xml") -Value "<Result>`r`n`t<Line>L1</Line>`r`n`t<Station>Station1</Station>`r`n`t<serialNumber>$($response.result.serialNumber)</serialNumber>`r`n`t<Employee>F001</Employee>`r`n`t<status>Failed</status>`r`n`t<description>$description</description>`r`n`t<errorCode>0</errorCode>`r`n</Result>"

         }


       #  }

   # else { 
        # Helpers-TimedTextOut "Falha em upload. Número de arquivos incompleto." -ForegroundColor Red
        # New-Item -Path (Join-Path $currentExecutingPath "..\..\SFC\Result") -Name "$($response.result.serialNumber).xml" -ItemType file -ErrorAction SilentlyContinue
        # Set-Content -Path (Join-Path $currentExecutingPath "..\..\SFC\Result\$($response.result.serialNumber).xml") -Value "<Result>`r`n`t<Line>L1</Line>`r`n`t<Station>Station1</Station>`r`n`t<serialNumber>$($response.result.serialNumber)</serialNumber>`r`n`t<Employee>F001</Employee>`r`n`t<status>Failed</status>`r`n`t<description>$description</description>`r`n`t<errorCode>0</errorCode>`r`n</Result>"
 
         
           # }
 
         # Backup of files to network (MAOWVENG02)
     try{
                $BCK = "\\matrizmao\Android_TEE\$model"
                $net = new-object -ComObject WScript.Network
                $net.MapNetworkDrive("r:", "$BCK", $false, "MAOWVENG02\LogFalhasLinha", "LogFalhasLinha123**")
                Start-Sleep -Seconds 1
                Copy-Item -Path $Dir\*.* -Destination  $BCK\
         }

      catch{}


     Start-Sleep -Seconds 2
     
     $net.RemoveNetworkDrive('r:', $true, $true)
     Remove-Item $Dir\* -ErrorAction SilentlyContinue -Force
     Helpers-TimedTextOut "Efetuado Backup do TEE para MAOWVENG02\LogFalhasLinha\$model." -ForegroundColor Green
   

 }
     
    }

    
    # check if exists CSV files and try to send to CEPP server
    if ($enableCEPPLog -and (Test-Path (Join-Path $Global:LogsFolder "Processados\*.csv")))
    {
        #Sending Processed files to Cepp server and copying them from "Processados" folder to "Processados_Bkp" folder
        Write-Host "Verificando comunicação com servidor CEPP..." -ForegroundColor Yellow
        if (-not (Test-Path $CeppServer -pathType container)) {
            Write-Host "Não foi possível enviar os logs para a pasta $CeppServer." -ForegroundColor Red
        } else {
            foreach ($file in (Get-ChildItem -Path (Join-Path $Global:LogsFolder "Processados\*.csv") )){
                Write-Host "Enviando para servidor CEPP o arquivo $file" -ForegroundColor Yellow
                Copy-Item $file $CeppServer -ErrorAction SilentlyContinue -ErrorVariable Erros
                if (-not $Erros) {
                    Copy-Item $file (Join-Path $Global:LogsFolder "ProcessadosBkp")
                    Remove-item $file
                } else {
                    Write-Host "Falha ao copiar o arquivo: $Erros" -ForegroundColor Red
                }
            }
        }
    }
    
}

if ($postit) {
    $postit.Dispose()
}

Unregister-Event RequestFileCreated
Remove-Job $requestFolderMonitoring[0]
$requestFolderMonitoring[1].Dispose()
$requestFolderMonitoring[2].Dispose()

Unregister-Event MiiFileCreated
Remove-Job $TeeFolderMonitoring[0]
$TeeFolderMonitoring[1].Dispose()
$TeeFolderMonitoring[2].Dispose()


Unregister-Event AckFileCreated
Remove-Job $responseFolderMonitoring[0]
$responseFolderMonitoring[1].Dispose()
$responseFolderMonitoring[2].Dispose()

$finalizationScriptStreamPipe.WriteLine("exit")
$finalizationScriptStreamPipe.Flush()

$finalizationScriptStreamPipe.Dispose()
$pipe.Dispose()
