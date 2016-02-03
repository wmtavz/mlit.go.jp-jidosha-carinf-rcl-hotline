
$urlPath='http://carinf.mlit.go.jp/jidosha/carinf/opn/'
function createUrl([int] $page){return "{0}{1}{2}" -f $urlPath,'search.html?selCarTp=1&lstCarNo=000&txtFrDat=1000/01/01&txtToDat=9999/12/31&txtNamNm=&txtMdlNm=&txtEgmNm=&chkDevCd=&page=',$page}

1..5000|%{
    echo "Page:$_"
    $date1=(get-date)
    
    $web=Invoke-WebRequest -Uri (createUrl $_)
    $table=$web.ParsedHtml.getElementById("r1")[0]
    $table.outerHTML|out-file "C:\Users\esadmin\Desktop\Work\CC2_Data\table_html_${_}.html" 
    
    $nextLink=$null;
    $web.ParsedHtml.links|?{($_.href -Like "*search.html?selCarTp*") -and ($_.innerHTML -like "*次のページ*")}|%{$nextLink=$_}
    if($nextLink -eq $null){break}
    
    $date2=(get-date)    
    "`t`t{0}" -f ($date2-$date1).TotalSeconds
}
exit


