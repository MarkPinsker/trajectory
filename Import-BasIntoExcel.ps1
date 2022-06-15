Add-Type -AssemblyName PresentationFramework

try {
	write-host "Open excel";
	$excel = New-Object -ComObject excel.application -ErrorAction Stop;
	$excel.visible = $True;
	$workbook = $excel.Workbooks.Add();
	write-host "add workbook";
}
catch {
	[system.windows.messagebox]::show("Unable to open Excel. 
		Error message was $_","Trajectory");
	exit 1;
}

try {
	$macros = Get-ChildItem -Path $psscriptroot\Macrochartv1.bas -File -ErrorAction Stop;
}
catch {
	[system.windows.messagebox]::show("Unable to find macro Macrochartv1.bas in $psscriptroot. 
		Error message was:- 
		$_","Trajectory");
	$workbook.close();
	$excel.quit();
	exit 2;
}

try {
	foreach ($macro in $macros) {
		$workbook.VBProject.VBComponents.Import($macro.FullName)|out-null;
		$macroname= $macro.FullName;
		write-host "add macro $macroname";
	}
}
catch {
	[system.windows.messagebox]::show("Unable to import macro $macroname into VBA. 
	Try this in Excel:-  File->Options->trust centre->Trust centre settings  
    Macro settings  
	Enable VBA macros  
	Trust access to VBA project object model 
	Error message was:- 
	$_","Trajectory");
	$workbook.close();
	$excel.quit();
	exit 3;
}


try {
	$VBAfuncname = 'mcrCalculate';	
	write-host "Run macro $VBAfuncname";
	$excel.run($VBAfuncname);
}
catch {
	[system.windows.messagebox]::show("Unable to run macro $VBAfuncname. 
	Error message was 
	$_","Trajectory");
	exit 4;
}


