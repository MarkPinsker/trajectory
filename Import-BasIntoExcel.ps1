Add-Type -AssemblyName PresentationFramework

write-host "open excel";
$excel = New-Object -ComObject excel.application ;
$excel.visible = $True;
$workbook = $excel.Workbooks.Add();
write-host "add workbook";

try {
	$macros = Get-ChildItem -Path $psscriptroot\Macrochartv1.bas -File;
}
catch {
	[system.windows.messagebox]::show("Unable to list macros. 
		Error message was $_");
	$workbook.close();
	$excel.quit();
	exit 1
}

try {
	foreach ($macro in $macros) {
		$workbook.VBProject.VBComponents.Import($macro.FullName)|out-null;
		$macroname= $macro.FullName;
		write-host "add macro $macroname";
	}
}
catch {
	[system.windows.messagebox]::show("unable to import macros. 
	Try this in Excel:-  File->Options->trust centre->Trust centre settings  
    Macro settings  
	Enable VBA macros  
	Trust access to VBA project object model 
	Error message was 
	$_");
	$workbook.close();
	$excel.quit();
	exit 2
}

write-host "Run macro mcrCalculate";
$excel.run('mcrCalculate');
sleep 10

