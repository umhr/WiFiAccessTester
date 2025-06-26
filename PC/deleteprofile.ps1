# ./deleteprofile.ps1 -target ./sp_targetlist.csv
# とすると、明示的にリストの指定ができる

Param(
    [String] $target = ""
)

$isParam = $true
if($target -eq ""){
    $target = "./targetlist.csv"
    $isParam = $false
}

. ".\WiFi-Class.ps1"

$wifiClass = New-Object WiFiClass

# ターゲットとなるWi-FiのリストをCSVファイルから取得
# data.csvが存在しない場合は、./targetlist.csvを使用
$targetCSV = $wifiClass.SetTargetCSV($target, $isParam)

# targetCSVに記載されているSSIDのWi-Fiプロファイルを削除する
$wifiClass.DeleteProfiles($targetCSV)
