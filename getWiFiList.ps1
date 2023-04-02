chcp 65001

# 登録済みのWi-Fiリストを取得
$str = netsh wlan show profile | findstr "All User Profile     :"
$array = $str.Split("`r`n")
$array = $array | Sort-Object

# リストからSSIDを取り出し、構造体を作りwifiListに入れる
$wifiList = @()
foreach ($item in $array)
{
    if($item.IndexOf(":") + 2 -lt $item.Length){
        $obj = [PSCustomObject]@{
            ssid = $item.Substring($item.IndexOf(":") + 2)
            pw = ''
        }
        $wifiList += $obj
    }
}
$wifiList.Length

# pw取得して構造体に書き込む
foreach ($item in $wifiList)
{
    # 文字列をコマンドレットとして実行
    $str = Invoke-Expression ('netsh wlan show profiles name="' + $item.ssid + '" key=clear')
    $str = $str | findstr "Key Content"
    if($str.Length -gt 3){
        if($str.IndexOf(":") + 2 -lt $str.Length){
            $item.pw = $str.Substring($str.IndexOf(":") + 2)
        }
    }
}

# ファイル出力
$wifiList | Export-Csv "wifilist.csv" -Encoding default -NoTypeInformation
