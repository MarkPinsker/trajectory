write-host "open excel";
$excel = New-Object -ComObject excel.application ;
$excel.visible = $True;
$workbook = $excel.Workbooks.Add();
write-host "add workbook";


$macros = Get-ChildItem -Path $psscriptroot\Macrochartv1.bas -File;
foreach ($macro in $macros) {
    $workbook.VBProject.VBComponents.Import($macro.FullName)|out-null;
    $macroname= $macro.FullName;
    write-host "add macro $macroname";
}
sleep 10
