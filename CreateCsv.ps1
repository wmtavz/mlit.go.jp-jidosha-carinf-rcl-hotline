Add-Type -AssemblyName System.Net.Http
# https://htmlagilitypack.codeplex.com/ からDLL入手
Add-Type -Path "C:\Apps\HtmlAgilityPack.1.4.6\Net40\HtmlAgilityPack.dll"
$xpathList=(
  # 1:番号
    "/td[1]",

  # 2:受付日
    "td[2]/div[1]",
  # 3:性別
    "td[2]/div[2]",
  # 4:住所
    "td[2]/div[3]",
  # 5:申告方法
    "td[2]/div[4]",

  # 6:車名
    "td[3]/div[1]",
  # 7:通称名
    "td[3]/div[2]",
  # 8:初度登録年月
    "td[3]/div[3]",
  # 9:総走行距離
    "td[3]/div[4]",
 # 10:型式
    "td[3]/div[5]",
 # 11:原動機型式
    "td[3]/div[6]",

 # 12:不具合装置
    "td[4]/div[1]",
 # 13:発生時期
    "td[4]/div[2]",
 # 14:申告内容の要約
    "td[4]/div[3]"
)


# $h="D:\Watson\CC2_Data\table_html_9.html"
# $c="${h}.csv"

function html2csv($h,$c){
    $doc = New-Object HtmlAgilityPack.HtmlDocument
    $doc.Load($h)
    
    # 位置ファイルにMax　10行がある
    1..10|%{
        $xpathTR="/table[1]/tbody[1]/tr[$_]"
        if($doc.DocumentNode.SelectNodes($xpathTR) -eq $null){break}    
        $line='';
        foreach($xpath in $xpathList){
            if($line.Length -ne 0){$line+=","}        
            $myXpath="$xpathTR/$xpath"
            # $myXpath
            $cell=$doc.DocumentNode.SelectNodes($myXpath).InnerText
            if($cell.IndexOf('"') -ne -1){$cell=$cell.Replace('"','""');}        
            if($cell.IndexOf(',') -ne -1){$cell="`"$cell`"";}        
            $line+="$cell"
        }
        "$line"|Out-File -Append $c
    }
}

# function checkConvert($htmlList){
    # $list2=@()
    # $htmlList|%{
        # $h=$_
        # $c="${h}.csv"
        # if(test-path $c){}else{
            # # echo $c;
            # $list2+=@($h);
        # }
    # }
    # return $list2;
# }

$htmlList=1..5000|%{"D:\Watson\CC2_Data\html\table_html_$_.html"}|?{ test-path $_}
function doConvert($htmlList){
    $csvFolder='';
    foreach($h in $htmlList){
        # $h=$_
        $h
        if($csvFolder.Length -eq 0){
            $csvFolder="{0}_{1:yyyyMMdd-HHmmss}\" -f (Split-Path $h -Parent),(Get-Date)
        }
        if(test-path $csvFolder){}else{mkdir $csvFolder}
        $c="${csvFolder}\{0}.csv" -f (Split-Path $h -Leaf)
        if(test-path $c){continue;}
        echo "converting... $h to $c"
        html2csv  $h $c
    }
}
doConvert  $htmlList

# checkConvert  $htmlList
# doConvert (checkConvert  $htmlList)