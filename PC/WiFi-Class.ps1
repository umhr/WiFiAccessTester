class WiFiClass{
    # 保存済みCSVファイル
    [string] $dataCSVFile = "./data.csv"
    WiFiClass(){
        chcp 65001 # utf-8
        # chcp 932 # s-jis
    }

    [System.Object] SetTargetCSV($targetCSVFile, $isParam){
        # ターゲットのCSVファイルを読み込む
        if($isParam -eq $false){
            if(Test-Path $this.dataCSVFile){
                $targetCSVFile = $this.dataCSVFile
            }
        }
        return Import-Csv -Path $targetCSVFile -Encoding UTF8
    }

    [Void] SaveTargetCSV($targetCSV){
        # ターゲットのCSVファイルを保存する
        $targetCSV | Export-Csv -Path $this.dataCSVFile -NoTypeInformation -Encoding UTF8
    }

    # 現在OSで電波を認識しているWi-Fiのリスト。電波の強い順
    [PSCustomObject] GetCurrentNetworks(){
        $result = $this.WlanShowNetworks()
        # Wi-Fiのリストが2つ未満の場合は、設定画面を開いて電波を取得しなおす
        if($result.Length -lt 2){
            $this.OpenSystemSettings()
            $result = $this.WlanShowNetworks()
        }
        # 電波の強い順に並べる
        return $result | Sort-Object -Property signal -Descending
    }

    [PSCustomObject] WlanShowNetworks(){
        # 配列
        $result = @()
        $networks = netsh wlan show networks mode=bssid
        $array = $networks.Split("`r`n")
        $str = ""
        foreach($item in $array){
            if($item.IndexOf("SSID") -eq 0){
                if($str.Length -gt 0){
                    $str += "======"
                }
                $n = [int] $item.Substring($item.IndexOf("SSID") + 5, 2)
                $str += ($n-1)
                $str += ","
                $str += $item.Substring($item.IndexOf(":") + 2)
            }
            if($item.IndexOf("Signal") -gt -1){
                $str += ","
                $str += $item.Substring($item.IndexOf(":") + 2).Trim()
            }
        }

        $List = $str.Split("======")
        foreach($ite in $List){
            $lis = $ite.Split(",")
            if($lis[1].Length -gt 0){
                $signal = 0
                if(-not ([string]::IsNullOrEmpty($lis[2]))){
                    $signal = [int]$lis[2].Substring(0, $lis[2].IndexOf("%"))
                }
                $result += New-Object PSObject -Property @{ssid = $lis[1]; signal = $signal}
            }
        }
        return $result
    }
    
    [Void] OpenSystemSettings(){
        $processName = "SystemSettings"
        # 設定画面プロセスが開いているかの確認
        $ssResult = Get-Process -Name $processName -EA 0
        if ($ssResult.Length -eq 1) {
            # SystemSettingsは存在するので閉じる
            Stop-Process -Name $processName
        }
        # 設定画面を開く
        start ms-settings:wifi-provisioning
        # 5秒待機 待たないと十分に電波を取得できないことがあるっぽい
        Start-Sleep -s 5
        Stop-Process -Name $processName
    }

    [Void] SaveListFile($List){
        $DirectoryName = (Get-Item .\).FullName
        $filename = (Get-Item Wi-Fi.xml).FullName
        $xmlFile = [xml] (Get-Content $filename)
        New-Item ".\__tempdata" -ItemType Directory -ErrorAction SilentlyContinue

        foreach($item in $List){
            $hex = $item.ssid | Format-Hex -Encoding UTF8
            $hexString = ([System.BitConverter]::ToString($hex.Bytes) -replace '-', '')
            
            # XMLを書き換えて保存する
            $xmlFile.WLANProfile.SSIDConfig.SSID.hex = $hexString
            $xmlFile.WLANProfile.name = $item.ssid
            $xmlFile.WLANProfile.SSIDConfig.SSID.name = $item.ssid
            $xmlFile.WLANProfile.MSM.security.sharedKey.keyMaterial = $item.pw
            $xmlFile.WLANProfile.MSM.security.authEncryption.authentication = "WPA2PSK"
            $xmlFile.WLANProfile.MSM.security.authEncryption.encryption = "AES"
            $file = $DirectoryName + "\__tempdata\Wi-Fi-" + $item.ssid + ".xml"
            if(Test-Path $file){
                Remove-Item $file -Force
            }
            $xmlFile.Save($file)
            # 優先度を1に設定
            $this.SetPriority($item.ssid, 1)
        }
        
        # 生成したxmlをPCに読み込ませる
        Get-ChildItem -Recurse ./__tempdata\*.xml | ForEach-Object {
            $file = $_.FullName
            netsh wlan add profile filename="$file"
        }
        # 自動接続を有効にする
        foreach($item in $List){
            netsh wlan set profile name=$($item.ssid) autoconnect=enabled user=true
        }
        # 一時生成したフォルダを削除
        # errorAction SilentlyContinueをつけないと、削除できない場合にエラーが出る
        Remove-Item -Path ".\__tempdata" -Recurse -Force -ErrorAction SilentlyContinue

    }

    [string] SetPriority($ssid, $priority){
        # 優先度を設定する
        return (netsh wlan set profileorder name="$ssid" interface=Wi-Fi priority=$priority | Out-Null)
    }

    [PSCustomObject] SetPriorities($currentNetworks, $targetCSV){
        # 優先度を設定する
        foreach($item in $targetCSV){
            # 現在のOSで認識されているSSIDの優先度を設定
            $wifi = $this.GetTarget($item.ssid, $currentNetworks)
            if([int]$item.priority -ge 10){
                # 優先度が10以上の場合はスキップ
                continue
            }
            if($null -eq $wifi){
                $item.priority = Get-Random -Minimum 5 -Maximum 8
                continue
            }
            if([int]$wifi.signal -gt 80){
                # 電波が80%以上の場合は優先度を1-2に設定
                $priority = Get-Random -Minimum 1 -Maximum 2
            }elseif([int]$wifi.signal -gt 50) {
                # 電波が50%以上の場合は優先度を2-3に設定
                $priority = Get-Random -Minimum 2 -Maximum 3
            }else{
                # 電波が50%未満の場合は優先度を3-4に設定
                $priority = Get-Random -Minimum 3 -Maximum 4
            }

            # 優先度を設定する
            $this.SetPriority($item.ssid, $priority)
            $item.priority = $priority
        }
        return $this.GetLowPriorites($targetCSV, 10)
    }
    [PSCustomObject] GetLowPriorites($targetCSV, $priority){
        $list = $targetCSV | Sort-Object -Property priority
        $result = @()
        foreach($item in $list){
            if($([int16]$item.priority -lt $priority)){
                $result += $item
            }
        }
        return $result
    }
    
    [boolean] Has5LowPriority($targetCSV){
        # 優先度5未満が設定されているか確認
        foreach($item in $targetCSV){
            if($([int16]$item.priority -lt 5)){
                return $true
            }
        }
        return $false
    }

    [PSCustomObject] GetTarget($ssid, $targetCSV){
        foreach($item in $targetCSV){
            if($item.ssid -eq $ssid){
                return $item
            }
        }
        return $null
    }

    [PSCustomObject] GetProfiles(){
        # 現在OSに登録済みWiFiプロファイルを取得して、優先度を設定する
        $result = @()
        $str = netsh wlan show profiles
        $array = $str.Split("`r`n")
        foreach($line in $array){
            if($line.IndexOf("All User Profile     : ") -gt -1){
                $ssid = $line.Substring($line.IndexOf("All User Profile     : ") + "All User Profile     : ".Length)
                $result += New-Object PSObject -Property @{ssid = $ssid}
                $this.SetPriority($ssid, 10)
            }
        }
        return $result
    }

    [PSCustomObject] GetInterfaces(){
        # 現在接続中のSSID
        $result = @()
        $str = netsh wlan show interfaces
        $array = $str.Split("`r`n")
        foreach($line in $array){
            if($line.IndexOf(" SSID") -gt 0){
                $ssid = $line.Substring($line.IndexOf(":") + 2)
                $result += New-Object PSObject -Property @{ssid = $ssid}
            }
        }
        return $result
    }

    [int16] SendPing($IPAddress){
        # ターゲットにpingを打つ
        $result = Test-Connection -ComputerName $IPAddress -Count 1 -ErrorAction SilentlyContinue
        if($result){
            #write-host "Ping to $IPAddress successful."
            return 1
        } else {
            #write-host "Ping to $IPAddress failed."
            return 0
        }
    }

    [int16] SendHttp($url){
        #write-host "Sending HTTP request to $url"
        try{
            #[Microsoft.PowerShell.Commands.HtmlWebResponseObject]
            $resp = Invoke-WebRequest ($url)
            $statusCode = $resp.StatusCode
        }catch{
            $statusCode = $_.Exception.Response
        }
        if($statusCode -eq 200){
            #write-host "HTTP request to $url successful."
            return 1
        } else {
            #write-host "HTTP request to $url failed with status code: $statusCode"
            return 0
        }
    }

    [string] GetIPAddress(){
        $IPv4 = "0.0.0.0"
        $str = ipconfig
        $array = $str.Split("`r`n")
        foreach($line in $array){
            if($line.IndexOf("IPv4 Address") -gt -1){
                $IPv4 = $line.Substring($line.IndexOf(" : ") + 3)
            }
        }
        return $IPv4
    }

    [Void] DeleteProfiles($targetCSV){
        # targetCSVに記載されているSSIDのWi-Fiプロファイルを削除する
        $str = netsh wlan show profiles
        $array = $str.Split("`r`n")
        [array]::Reverse( $array )
        foreach($obj in $targetCSV){
            Write-Host $($obj.ssid + ":" + $targetCSV.Length + " / " + $array.Length)
            foreach($line in $array){
                if($line.IndexOf("All User Profile     : ") -gt -1){
                    $ssid = $line.Substring($line.IndexOf("All User Profile     : ") + "All User Profile     : ".Length)
                    Write-Host $($obj.ssid + " - " + $ssid)
                    if($obj.ssid -eq $ssid){
                        Write-Host "Deleting profile: $ssid with priority $($obj.priority)"
                        netsh wlan delete profile name="$ssid" | Out-Null
                        $index = [array]::IndexOf($array, $line)
                        $array[$index] = ""
                        break
                    }
                }
            }
        }
    }

}

<#
    start ms-settings:network-wifi
    Windows 10 バージョン1703で見つかった「ms-settings:URIスキーム」の全リスト
    https://atmarkit.itmedia.co.jp/ait/articles/1707/11/news009_2.html
#>