function Connect-MySQL([string]$user, [string]$pass, [string]$MySQLHost, [string]$database, [string]$filepath, [string]$ArchivePath)
{
    # Load MySQL .NET Connector Objects
    [void][system.reflection.Assembly]::LoadWithPartialName("MySql.Data")

    $connStr = "server=" + $MySQLHost + ";port=1234;uid=" + $user + ";pwd=" + $pass + ";database=" + $database + ";Pooling=FALSE"
    $conn = New-Object MySql.Data.MySqlClient.MySqlConnection($connStr)
    $conn.Open()

#create object to open Excel workbook

$Excel = New-Object -ComObject Excel.Application

$Workbook = $Excel.Workbooks.Open($filepath)

$Worksheet = $Workbook.Worksheets.Item(1)

$startRow = 2
$sourcefolder = "C:"

$FileName = Get-ChildItem "FileSystem::$sourceFolder\*" -Include @("*.xlsx")

$batchcounter=0
$batchsize=1000
$MysqlValues = New-Object Collections.ArrayList

Do {

$ColValues1 = $Worksheet.Cells.Item($startRow, 1).Value()

$ColValues2 = $Worksheet.Cells.Item($startRow, 2).Value()

$MysqlValues.Add("(
    '$ColValues1',
    '$ColValues2'))")

$startRow++
$batchcounter++

if ($batchcounter % $batchsize -eq 0) {
    $Mysql = "INSERT INTO `table_name(Column_name1,
Column_name2
)` values {0}" -f ($MysqlValues.ToArray() -join "`r`n,")

    $command = New-Object MySql.Data.MySqlClient.MySqlCommand
    $command.Connection = $connStr
    $command = $conn.CreateCommand()
    $command.CommandText = $Mysql
    $result = $command.ExecuteNonQuery()
    $MysqlValues.Clear()
}

}

While ($Worksheet.Cells.Item($startRow,1).Value() -ne $null)

if ($batchcounter -gt 0) {
        $Mysql = "INSERT INTO `table_name(Column_name1,
Column_name2)` values {0}" -f ($MysqlValues.ToArray() -join "`r`n,")

    $command = New-Object MySql.Data.MySqlClient.MySqlCommand
    $command.Connection = $connStr
    $command = $conn.CreateCommand()
    $command.CommandText = $Mysql
    $result = $command.ExecuteNonQuery()
    $MysqlValues.Clear()
}

$Excel.Quit()

   $conn.Close()

}

$result = Connect-MySQL  $user $password $mysqlhost $database $filepath $ArchivePath
