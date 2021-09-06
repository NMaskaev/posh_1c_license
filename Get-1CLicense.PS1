# Скрипт получения перечня выданных лицензий на кластером 1С
$ServerName           = "<Server1C>"
$ComConnectorName     = "V83.COMConnector"

$ClusterAdmin         = "<Cluster admin login>"
$ClusterAdminPassword = "<Cluster admin passw>"

# Подключаемся к агенту сервера
$V83Connector    = New-Object -ComObject $ComConnectorName
$AgentConnector  = $V83Connector.ConnectAgent($ServerName)

# Получаем список кластеров
$Clusters    = $AgentConnector.GetClusters()
$MainCluster = $Clusters[0]

# Авторизуемся в нашем кластере
$AgentConnector.Authenticate($MainCluster,$ClusterAdmin,$ClusterAdminPassword)

# Получим список сеансов
$ClusterSessions = $AgentConnector.GetSessions($MainCluster)

# Выбираем сеансы которым назначены лицензии
$SessionsWithLicense = $ClusterSessions | Where-Object {$_.License -ne $null}

$LicenseTable = $SessionsWithLicense | Select-Object Host,AppId,userName, @{Label="BaseName"; E={$_.InfoBase.Name}},@{L="License";E={$_.License.ShortPresentation}},`
 @{Label = "Спящая"; E= {$_.Hibernate}} | Sort userName

$dubledSession = 0
$curr          = "_"
$base          = "_"

foreach($licItem in $LicenseTable){
    
    if( $licItem.userName -eq $curr`
        -and`
        $licItem.BaseName -eq $base )
    {
        $dubledSession +=1

    } else {
        
        $curr = $licItem.userName
        $base = $licItem.BaseName
    }
}

$Lan50      = @()
$Lan100     = @()
$Prog100    = @()
$Hibernated = @()
$WebClients = @()

$Lan50      += $LicenseTable | Where-Object {$_.License -match "50$"}
$Lan100     += $LicenseTable | Where-Object {$_.License -match "Сет 100$"}
$Prog100    += $LicenseTable | Where-Object {$_.License -match "100 100$"}
$Hibernated += $LicenseTable | Where-Object {$_.Спящая  -match $true}
$WebClients += $LicenseTable | Where-object {$_.AppID   -match "WebClient"}

Write-Host "Занято сетевых    50  :" ($Lan50.Count)
Write-Host "Занято сетевых    100 :" ($Lan100.Count)
Write-Host "Занято програмных 100 :" ($Prog100.Count)
Write-Host "Задвоенных сессий     :" ($dubledSession)
Write-Host "Спящих сессий         :" ($Hibernated.Count)
Write-host "Веб сессий            :" ($WebClients.Count)

Write-Host ""
Write-Host "Спящие сессии"
$Hibernated | sort License | ft


Write-Host "Все сессии"
$LicenseTable | sort License | ft



