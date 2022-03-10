Clear-Host

Write-Output "--== S T A R T ==--"

$destination_path_prefix = "$PSScriptRoot\out\"
Remove-Item $destination_path_prefix -Recurse -Force
New-Item  -ItemType "directory" $destination_path_prefix


function UpdateString($datafile, $first_name, $last_name, $title, $email, $phone, $mobile) {

  $datafile = $datafile -replace "%%name%%", ($first_name + " " + $last_name)
  $datafile = $datafile -replace "%%title%%", $title
  $datafile = $datafile -replace "%%email%%", $email

  if ( $phone ) {
    $datafile = $datafile -replace "%%phone%%", ("phone: " + $phone)
  }
  else {
    $datafile = $datafile -replace "%%phone%%", ("")
  }


 if ( $mobile ) {
    $datafile = $datafile -replace "%%mobile%%", ("mobile: " + $mobile)
  }
  else {
    $datafile = $datafile -replace "%%mobile%%", ("")
  }

  return $datafile

}


$Db = ""$PSScriptRoot\employees.mdb"
$cursor = 3
$lock = 3

$conn = New-Object -ComObject ADODB.Connection
$recordset = New-Object -ComObject ADODB.Recordset
$conn.Open("Provider=Microsoft.Ace.OLEDB.12.0;Data Source=$Db")

$query = "Select * from  employees"
$recordset.open($query,$conn,$cursor,$lock)

while (!$recordset.eof){

    $first_name = $recordset.Fields("first_name").value
    $last_name = $recordset.Fields("last_name").value
    $title = $recordset.Fields("title").value
    $email = $recordset.Fields("email").value
    $phone = $recordset.Fields("phone").value
    $mobile = $recordset.Fields("mobile").value


  $Destination_path = $destination_path_prefix + $employee.first_name + "_" + $employee.last_name

  New-Item -ItemType "directory" -Path $Destination_path

  $file_txt = Get-Content "data\$template.txt" -Encoding Default -Raw
  $file_rtf = Get-Content "data\$template.rtf" -Encoding Default -Raw
  $file_htm = Get-Content "data\$template.htm" -Encoding Default -Raw

  $file_txt = UpdateString $file_txt $employee.first_name $employee.last_name $employee.title $employee.email $employee.phone $employee.mobile
  $file_htm = UpdateString $file_htm $employee.first_name $employee.last_name $employee.title $employee.email $employee.phone $employee.mobile
  $file_rtf = UpdateString $file_rtf $employee.first_name $employee.last_name $employee.title $employee.email $employee.phone $employee.mobile

  Out-File -FilePath ("$Destination_path\$template.txt") -InputObject $file_txt -Encoding Default
  Out-File -FilePath ("$Destination_path\$template.rtf") -InputObject $file_rtf -Encoding Default
  Out-File -FilePath ("$Destination_path\$template.htm") -InputObject $file_htm -Encoding Default

  Copy-Item -Path ("data\" + $template + "_files") -Destination $Destination_path -Recurse
  
  $recordset.MoveNext()

}

Write-Output "--== E N D ==--"

   
