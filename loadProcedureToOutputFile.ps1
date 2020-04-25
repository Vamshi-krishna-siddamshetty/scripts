Param
(

[Parameter(Mandatory=$True)]
[string]$procedureName= "`procedurename`()",
[Parameter(Mandatory=$True)]
[string]$filename= "Output"
)

function Connect-MySQL([string]$user, [string]$pass, [string]$MySQLHost, [string]$database, [string]$filepath)
{

[void][system.reflection.Assembly]::LoadWithPartialName("MySql.Data")
[void][system.reflection.Assembly]::LoadWithPartialName("MySql.DataAdapter")
$connStr = "server=" + $MySQLHost + ";port=3306;uid=" + $user + ";pwd=" + $pass + ";database=" + $database + ";Allow User Variables=True;Pooling=FALSE"
$conn = New-Object MySql.Data.MySqlClient.MySqlConnection($connStr)
$conn.Open()

$excel = New-Object -ComObject Excel.Application
$workbook = $excel.Workbooks.add()
$worksheetA = $workbook.Worksheets.Add()
$sheet1 = $workbook.worksheets.Item(1)

$query = "CALL $procedureName"
$command = New-Object MySql.Data.MySqlClient.MySqlCommand($query, $connStr)
$command.CommandTimeout = 0;
$dataAdapter = New-Object MySql.Data.MySqlClient.MySqlDataAdapter($command)
$dataSet = new-object "System.Data.DataSet" "Sheet1"
Write-Output $dataAdapter
$dataAdapter.Fill($dataSet) | Out-Null

$sheet1.name = "Sheet1 $(get-date -f yyyy-MM-ddHHmmss)"

$conn.Close()

$dataTable = new-object "System.Data.DataTable" "SampleOutput"
$dataTable = $dataSet.Tables[0]

#assign column names

$sheet1.cells.item(1, 1) = "COLUMN_NAME"

$d = $sheet1.UsedRange
$d.Interior.ColorIndex = 19
$d.Font.ColorIndex = 11
$d.Font.Bold = $True

$x=2

$dataTable | FOREACH-OBJECT{

$sheet1.cells.item($x, 1) = $_.COLUMN_NAME

$x++

}

$range1 = $sheet1.UsedRange
$range1.EntireColumn.AutoFit()

$excel.ActiveWorkbook.SaveAs("$filepath $filename.xlsx ")
$excel.quit()

}

$result = Connect-MySQL  $user $password $mysqlhost $database $Outputfilepath
