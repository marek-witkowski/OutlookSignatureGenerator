  Clear-Host

  Write-Output "--== S T A R T ==--"

  $destination_path_prefix = "$PSScriptRoot\out\"
  Remove-Item $destination_path_prefix -Recurse -Force
  New-Item  -ItemType "directory" $destination_path_prefix


  function UpdateString($datafile, $first_name, $last_name, $title, $email, $phone, $mobile)
      {

        $datafile = $datafile -replace "%%name%%", ($first_name + " " + $last_name)
        $datafile = $datafile -replace "%%title%%", $title
        $datafile = $datafile -replace "%%email%%", $email
        

        if ( $email )
        {
          $datafile = $datafile -replace "%%phone%%",  ("e-mail: " + $email)
        }
      else
        {
          $datafile = $datafile -replace "%%phone%%",  ("")
        }


        if ( $phone )
          {
            $datafile = $datafile -replace "%%phone%%",  ("phone: " + $phone)
          }
        else
          {
            $datafile = $datafile -replace "%%phone%%",  ("")
          }



        if ( $mobile )
          {
            $datafile = $datafile -replace "%%mobile%%", ("mobile: " + $mobile)
          }
        else
          {
            $datafile = $datafile -replace "%%mobile%%", ("")
          }

          return $datafile

        }

  $CSV = Import-Csv -Path "$PSScriptRoot\employees.csv" -Encoding Default -Delimiter ";"
  foreach ($employee in $CSV){

      
      $Destination_path = $destination_path_prefix + $employee.first_name + "_" + $employee.last_name

      New-Item -ItemType "directory" -Path $Destination_path


      $file_txt = Get-Content "data\$template.txt" -Encoding Default -Raw
      $file_rtf = Get-Content "data\$template.rtf" -Encoding Default -Raw
      $file_htm = Get-Content "data\$template.htm" -Encoding Default -Raw

      $file_txt = UpdateString $file_txt $employee.name $employee.last_name $employee.title $employee.email $employee.phone $employee.mobile
      $file_htm = UpdateString $file_htm $employee.name $employee.last_name $employee.title $employee.email $employee.phone $employee.mobile
      $file_rtf = UpdateString $file_rtf $employee.name $employee.last_name $employee.title $employee.email $employee.phone $employee.mobile

      Out-File -FilePath ("$Destination_path\$template.txt") -InputObject $file_txt -Encoding Default
      Out-File -FilePath ("$Destination_path\$template.rtf") -InputObject $file_rtf -Encoding Default
      Out-File -FilePath ("$Destination_path\$template.htm") -InputObject $file_htm -Encoding Default

      Copy-Item -Path ("data\" + $template + "_files") -Destination $Destination_path -Recurse

  }

  Write-Output "--== E N D ==--"
