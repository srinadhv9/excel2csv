$file = $args[0]
#$file = 'C:\Users\502362723\Desktop\10-Mar-2017-CPS-Report-FW-10_LV'   # xls, xlsx, xlsb or whatever file would be possible here.
Set-ExecutionPolicy RemoteSigned
$xlApp = New-Object -Com Excel.Application
$xlApp.Visible = $false
$wb = $xlApp.Workbooks.Open($file + '.xlsb')
$tabName = "CPS"
foreach ($ws in $wb.Worksheets)
{
  if($ws.Name -eq $tabName){
  $ws.SaveAs($file + '_' + $ws.Name, [Microsoft.Office.Interop.Excel.xlFileFormat]::xlCSVWindows)
}
}
$wb.Close(0)
$xlApp.Quit()

#powershell -ExecutionPolicy ByPass -File C:\\Users\\502362723\\Desktop\\xlsbtoxlsx.ps1 'C:\Users\502362723\Desktop\10-Mar-2017-CPS-Report-FW-10_LV' Set-ExecutionPolicy RemoteSigned