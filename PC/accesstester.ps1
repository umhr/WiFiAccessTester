<#
# PowerShellでこのシステムではスクリプトの実行が無効になっているため、ファイル hoge.ps1 を読み込むことができません。となったときの対応方法
# PowerShellを管理者権限で
# Set-ExecutionPolicy -ExecutionPolicy Bypass -Scope LocalMachine
# を実行する。

powershellから
netsh wlan show interface
が使えるか確認する。

ネットワーク シェル コマンドは、WLAN 情報にアクセスするには位置情報のアクセス許可が必要です。[プライバシーとセキュリティ] 設定の [位置情報] ページで位置情報サービスを有効にします。

設定アプリの [位置情報] ページの URI は次のとおりです:
ms-settings:privacy-location
設定アプリで [位置情報] ページを開くには、Ctrl キーを押しながらリンクを選択するか、次のコマンドを実行します。
start ms-settings:privacy-location
#>

Param(
    [String]$target = "./targetlist.csv",
    [boolean]$init = $true
)

$isParam = $true
if($target -eq ""){
    $target = "./targetlist.csv"
    $isParam = $false
}

$IncludeFile = Join-Path -Path $PSScriptRoot -ChildPath ".\WiFi-Class.ps1"
if( -not (Test-Path $IncludeFile) ){
    echo "[FAIL] $IncludeFile not found !"
    exit
}
. $IncludeFile
function  Set-Log {
    param (
        [string]$message
    )
    Write-Host $message
    $logFile = Join-Path -Path $PSScriptRoot -ChildPath "log.txt"
    try{
        Add-Content -Path $logFile -Value "$(Get-Date -Format 'yyyy/MM/dd HH:mm:ss') - $message"
    }catch{
        Write-Host "Save Error"
    }
}


$wifiClass = New-Object WiFiClass

# 現在OSで電波を認識しているWi-Fiのリスト。電波の強い順
$currentNetworks = $wifiClass.GetCurrentNetworks()

# 現在OSに登録済みのWi-Fiのリストを取得して、優先度を10に設定する
$wifiClass.GetProfiles()

# ターゲットとなるWi-FiのリストをCSVファイルから取得
# data.csvが存在しない場合は、./otaru_targetlist.csvを使用
$targetCSV = $wifiClass.SetTargetCSV($target, $isParam)

# targetCSV上の情報をPCに反映して、優先度を1-4に設定する
if($init){
    $wifiClass.SaveListFile($targetCSV)
}

# PCの名前を取得
$hostname = [System.Net.Dns]::GetHostName()
$Progress = $($wifiClass.GetLowPriorites($targetCSV, 5).Length.ToString() + " / " + $wifiClass.GetLowPriorites($targetCSV, 10).Length.ToString() + " / " + $targetCSV.Length.ToString())
Set-Log $("Hostname: $hostname, Progress: " + $Progress)

$logCount = 0

while($true){
    # 現在のWi-Fiの強さを取得
    $currentNetworks = $wifiClass.GetCurrentNetworks()
    # $targetCSVに存在して、Wi-Fi電波が強いものは優先度を高くし、優先度10未満のリストを返す
    $Priorites = $wifiClass.SetPriorities($currentNetworks, $targetCSV)
    if($Priorites.Length -eq 0){
        # 優先度10未満のWi-Fiが無い場合は、ループを抜ける
        Set-Log "No low priority Wi-Fi found."
        #continue
        break
    }
    if([int16]$Priorites[0].priority -ge 5){
        # 優先度が5以上のWi-Fiしか無い場合は、5秒待機後ループを抜ける
        if($logCount -eq 0){
            Set-Log "No low priority Wi-Fi found with priority less than 5."
        }elseif($logCount%100 -eq 0){
            write-host "."
        }else{
            write-host -NoNewline "."
        }
        $logCount ++
        Start-Sleep -Seconds 5
        continue
    }
    if($logCount -gt 0){
        write-host ""
    }
    $logCount = 0

    # 現在接続中のSSID
    $accesed = $wifiClass.GetInterfaces()
    if($accesed.Length -eq 0){
        # 接続が無い場合は優先度の高いWi-Fiに接続へ
        Set-Log "Connecting to Wi-Fi with SSID: $($Priorites[0].ssid) with priority: $($Priorites[0].priority)"
        netsh wlan connect name=$($Priorites[0].ssid) | Out-Null # 標準出力に出さない
        Start-Sleep -Seconds 10
        continue
    }
    # 今繋がっているWiFiのSSIDがtargetCSVにあるか確認
    $item = $wifiClass.GetTarget($accesed[0].ssid, $targetCSV)
    if($null -eq $item -or $([int16]$item.priority -ge 10)){
        Set-Log "Current SSID: $($accesed[0].ssid) is not in target CSV or has low priority."
        # 現在接続中のSSIDがtargetCSVに無い場合 or 優先度が10以上の場合は、接続解除へ
        netsh wlan disconnect | Out-Null # 標準出力に出さない
        Start-Sleep -Seconds 10
        continue
    }

    $AccessCount = 1
    $AccessCount += $wifiClass.SendPing("74.125.26.147")*10 # GoogleのIPアドレス
    if($AccessCount -eq 11){
        $AccessCount += $wifiClass.SendPing("192.168.10.5")*100 # ローカルエリア内のIPアドレス
    }
    if($AccessCount -eq 111){
        $url = "https://www.yahoo.co.jp/"
        $AccessCount += $wifiClass.SendHttp($url)*1000
    }
    if($AccessCount -eq 1111){
        $url = "http://192.168.10.5:8080/access/set?name=" + $($accesed[0].ssid + "&") + "language=" + $hostname
        $AccessCount += $wifiClass.SendHttp($url)*10000
    }

    if($AccessCount -eq 11111){
        #write-host "All pings and HTTP requests succeeded."
        $IPv4 = $wifiClass.GetIPAddress()
        foreach($item in $targetCSV){
            if($item.ssid -eq $accesed[0].ssid){
                $item.priority = 10
                $item.datetime = $(Get-Date -Format "yyyy/MM/dd HH:mm:ss")
                $item.count = $AccessCount
                #if($null -ne $item.ipv4){
                    $item.ipv4 = $IPv4
                #}
                break
            }
        }
        $wifiClass.SetPriority($accesed[0].ssid, 10)
    }else{
        #write-host "Some pings or HTTP requests failed."
        foreach($item in $targetCSV){
            if($item.ssid -eq $accesed[0].ssid){
                $item.datetime = $(Get-Date -Format "yyyy/MM/dd HH:mm:ss")
                $item.count = $AccessCount
                break
            }
        }
    }
    netsh wlan disconnect | Out-Null # 標準出力に出さない
    $wifiClass.SaveTargetCSV($targetCSV)
    $Progress = $($wifiClass.GetLowPriorites($targetCSV, 5).Length.ToString() + " / " + $wifiClass.GetLowPriorites($targetCSV, 10).Length.ToString() + " / " + $targetCSV.Length.ToString())
    Set-Log "Processed SSID: $($accesed[0].ssid), AccessCount: $AccessCount, Progress: $Progress"

    Start-Sleep -Seconds 10
}

Set-Log "All low priority Wi-Fi have been processed."

