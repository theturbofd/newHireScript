#+-------------------------------------------------------------------+            
#      Author:  Giovanni Ramos                                       |
#       Email:  Gramos147@gmail.com                                  |
# Description:  Script to pull new employee info from db and         |
#               autmotate the account creation process               |
#+-------------------------------------------------------------------+ 


#---------Import the needed module-----------------
Import-Module ActiveDirectory
 
#---------Function to check if username is in use--------
function preCheckUser($userName){
		$userCheck = Get-ADUser $userName -Server domain.com:3268
		$newName = $userName
		if ($userCheck -eq $null){
			Write-Host "User name $userName not in use!"
			return $newName
		}
		else{
			Write-Host "User name in use, switching naming convention!"
			$FirstU = $FirstName.Substring(0,2)
			$LastU = $LastName.Substring(0,3)
			$newName = "$FirstU$LastU"
            $newName = $newName.ToLower()
			Write-Host " User name will now be $newName"
			return $newName
		}
}

#----------Set up the connection to the database------------
$Query = 'SELECT empname,as4id,effdate FROM giotest.empmaintlog WHERE ad_addtimestamp IS NULL AND `action`="Add" ORDER BY effdate DESC'
$SQLUserName = READ-HOST "Username?"
$SQLPassword = READ-HOST "Password?" -asSecureString
$SQLPassword = [Runtime.InteropServices.Marshal]::PtrToStringAuto([Runtime.InteropServices.Marshal]::SecureStringToBSTR($SQLPassword))
$SQLDatabase = 'giotest'
$SQLHost = '*server Name'
$ConnectionString = "server=" + $SQLHost + ";port=3306;uid=" + $SQLUserName + ";pwd=" + $SQLPassword + ";database="+$SQLDatabase
$redo = 1


Try {
  [void][System.Reflection.Assembly]::LoadFrom("C:\Program Files (x86)\MySQL\MySQL Connector Net 6.9.8\Assemblies\v4.5\MySQL.Data.dll")
  $Connection = New-Object MySql.Data.MySqlClient.MySqlConnection
  $Connection.ConnectionString = $ConnectionString
  $Connection.Open()

  $Command = New-Object MySql.Data.MySqlClient.MySqlCommand($Query, $Connection)
  $DataAdapter = New-Object MySql.Data.MySqlClient.MySqlDataAdapter($Command)
  $DataSet = New-Object System.Data.DataSet
  $RecordCount = $dataAdapter.Fill($dataSet, "data")
}

Catch {
  Write-Host "ERROR : Unable to run query : $query `n$Error[0]"
}

Finally{
    $Connection.Close()
}

While ($redo -eq 1){
    #--------This stores the table data into the data variable------
    $data = $DataSet.Tables[0]
    $count = 0  
    
    #--------This creates the numbers next to the employee's --------
    foreach($dataName in $data.Rows){
	    $empName = $dataName[0]
        $as4id = $dataName[1]
        $effdate = $dataname[2]
        $out = new-object psobject
        $out | add-member noteproperty Choice $count
        $out | add-member noteproperty Employee_Name $empName
        $out | add-member noteproperty As400_Id $as4id
        $out | add-member noteproperty Effective_Date $effdate
        $out | Format-Table -autosize
	    $count = $count + 1
    }

	#--------Prompt user for selection based on the table they're shown---
    $User_Selection = READ-HOST "Make a selection"
    $Selection = $data.Rows[$User_Selection].empName
    #--------------------Splits the name to first,last and grabs username-----
    $LastName, $FirstName = $Selection.split(",")
    $FirstName = $FirstName.Substring(1)
    $userName = $data.Rows[$User_Selection].as4id
    #--------------------grabs supervisor,site,dept info----------------------
    $Command.CommandText = "SELECT location,supervisor,dept FROM *table* WHERE as4id= `"$username`""
    $DataAdapter.fill($DataSet, "info")
    $UserInfo = $DataSet.Tables[1]
    #--------------------Creates a table to display info to operator-----------
    foreach($info in $UserInfo.Rows){
        $location = $Info[0]
        $supervisor = $Info[1]
        $department = $Info[2]
        $out2 = new-object psobject
        $out2 | add-member noteproperty Name $selection
        $out2 | add-member noteproperty Username $username
        $out2 | add-member noteproperty Location $location
        $out2 | add-member noteproperty Supervisor $supervisor
        $out2 | add-member noteproperty Department $department
    }
    $out2 | Format-Table -autosize
    $Choice = READ-HOST "You want to create an account for this person ? (y/n)"

	
	#--------If the user chooses yes it creates user---------
    if($Choice -eq "y" -or $Choice -eq "yes"){

        #--------------------Checks the username and see if it's available---------
	    $checked_username = preCheckUser($userName)
	    $checked_username = $checked_username.ToLower()

        
        if($location -eq "Arc" -or $location -eq "Balt"){

    	    #---------This creates the user----------------------- 
    	    New-ADUser `
    	    -Name "$FirstName $LastName" `
    	    -GivenName "$FirstName" `
    	    -Surname "$LastName" `
    	    -Path "*insert path*" `
            -Server "domain.com" `
    	    -UserPrincipalName "$checked_username@domain.com" `
    	    -SamAccountName  "$checked_username" `
    	    -DisplayName "$FirstName $LastName" `
    	    -AccountPassword (ConvertTo-SecureString "InsertTempPassword" -AsPlainText -Force) `
    	    -ChangePasswordAtLogon $true `
            -PasswordNeverExpires $False `
    	    -Enabled $true

    	    #----------This adds the necessary groups--------------
    	    Start-Sleep -s 12
    	    
            Add-ADGroupMember -identity *insert adgroup* -members $checked_username -
    	    Add-ADGroupMember -identity *insert adgroup* -members $checked_username
    	    Add-ADGroupMember -identity *insert adgroup* -members $checked_username
            $Command.CommandText = "UPDATE *insert table* SET ad_addtimestamp = now() WHERE as4id = `"$username`""
            $RecordCount = $DataAdapter.Fill($dataset, "data")
        }
        elseif($location -eq "Clev"){

            if($department -eq "col"){
                #---------This creates the user----------------------- 
                New-ADUser `
                -Name "$FirstName $LastName" `
                -GivenName "$FirstName" `
                -Surname "$LastName" `
                -Path "*insert path*" `
                -Server "*insert servername*" `
                -UserPrincipalName "$checked_username@domain.com" `
                -SamAccountName  "$checked_username" `
                -DisplayName "$FirstName $LastName" `
                -AccountPassword (ConvertTo-SecureString "InsertTempPassword" -AsPlainText -Force) `
                -PasswordNeverExpires $False `
                -ChangePasswordAtLogon $true `
                -Enabled $true

                Start-Sleep -s 12
            }
            elseif($department -eq "elg"){
                #---------This creates the user----------------------- 
                New-ADUser `
                -Name "$FirstName $LastName" `
                -GivenName "$FirstName" `
                -Surname "$LastName" `
                -Path "*insert path*" `
                -Server "*insert servername*" `
                -UserPrincipalName "$checked_username@domain.com" `
                -SamAccountName  "$checked_username" `
                -DisplayName "$FirstName $LastName" `
                -AccountPassword (ConvertTo-SecureString "InsertTempPassword" -AsPlainText -Force) `
                -PasswordNeverExpires $False `
                -ChangePasswordAtLogon $true `
                -Enabled $true

                Start-Sleep -s 12
            }
            elseif($department -eq "fin"){

                #---------This creates the user----------------------- 
                New-ADUser `
                -Name "$FirstName $LastName" `
                -GivenName "$FirstName" `
                -Surname "$LastName" `
                -Path "*insert path*" `
                -Server "*insert servername*" `
                -UserPrincipalName "$checked_username@domain.com" `
                -SamAccountName  "$checked_username" `
                -DisplayName "$FirstName $LastName" `
                -AccountPassword (ConvertTo-SecureString "InsertTempPassword" -AsPlainText -Force) `
                -PasswordNeverExpires $False `
                -ChangePasswordAtLogon $true `
                -Enabled $true

                Start-Sleep -s 12
            }
            elseif($department -eq "ins"){

                #---------This creates the user----------------------- 
                New-ADUser `
                -Name "$FirstName $LastName" `
                -GivenName "$FirstName" `
                -Surname "$LastName" `
                -Path "*insert path*" `
                -Server "*insert servername*" `
                -UserPrincipalName "$checked_username@domain.com" `
                -SamAccountName  "$checked_username" `
                -DisplayName "$FirstName $LastName" `
                -AccountPassword (ConvertTo-SecureString "InsertTempPassword" -AsPlainText -Force) `
                -PasswordNeverExpires $False `
                -ChangePasswordAtLogon $true `
                -Enabled $true

                Start-Sleep -s 12
            }
            elseif($department -eq "sal"){

                #---------This creates the user----------------------- 
                New-ADUser `
                -Name "$FirstName $LastName" `
                -GivenName "$FirstName" `
                -Surname "$LastName" `
                -Path "*insert path*" `
                -Server "*insert servername*" `
                -UserPrincipalName "$checked_username@domain.com" `
                -SamAccountName  "$checked_username" `
                -DisplayName "$FirstName $LastName" `
                -AccountPassword (ConvertTo-SecureString "InsertTempPassword" -AsPlainText -Force) `
                -PasswordNeverExpires $False `
                -ChangePasswordAtLogon $true `
                -Enabled $true

                Start-Sleep -s 12
            }
            else{
                exit
            }
            Add-ADGroupMember -identity *insert adgroup* -members $checked_username -Server "*insert servername*"
            Add-ADGroupMember -identity *insert adgroup* -members $checked_username -Server "*insert servername*"
            Add-ADGroupMember -identity *insert adgroup* -members $checked_username -Server "*insert servername*"
            $Command.CommandText = "UPDATE *insert table* SET ad_addtimestamp = now() WHERE as4id = `"$username`""
            $RecordCount = $DataAdapter.Fill($dataset, "data")
        }
        elseif($location -eq "Bost"){

            if($department -eq "col"){
                #---------This creates the user----------------------- 
                New-ADUser `
                -Name "$FirstName $LastName" `
                -GivenName "$FirstName" `
                -Surname "$LastName" `
                -Path "*insert path*" `
                -Server "*insert servername*" `
                -UserPrincipalName "$checked_username@domain.com" `
                -SamAccountName  "$checked_username" `
                -DisplayName "$FirstName $LastName" `
                -AccountPassword (ConvertTo-SecureString "InsertTempPassword" -AsPlainText -Force) `
                -PasswordNeverExpires $False `
                -ChangePasswordAtLogon $true `
                -Enabled $true

                Start-Sleep -s 12
            }
            elseif($department -eq "elg"){
                #---------This creates the user----------------------- 
                New-ADUser `
                -Name "$FirstName $LastName" `
                -GivenName "$FirstName" `
                -Surname "$LastName" `
                -Path "*insert path*" `
                -Server "*insert servername*" `
                -UserPrincipalName "$checked_username@domain.com" `
                -SamAccountName  "$checked_username" `
                -DisplayName "$FirstName $LastName" `
                -AccountPassword (ConvertTo-SecureString "InsertTempPassword" -AsPlainText -Force) `
                -PasswordNeverExpires $False `
                -ChangePasswordAtLogon $true `
                -Enabled $true

                Start-Sleep -s 12
            }
            elseif($department -eq "fin"){
                #---------This creates the user----------------------- 
                New-ADUser `
                -Name "$FirstName $LastName" `
                -GivenName "$FirstName" `
                -Surname "$LastName" `
                -Path "*insert path*" `
                -Server "*insert servername*" `
                -UserPrincipalName "$checked_username@domain.com" `
                -SamAccountName  "$checked_username" `
                -DisplayName "$FirstName $LastName" `
                -AccountPassword (ConvertTo-SecureString "InsertTempPassword" -AsPlainText -Force) `
                -PasswordNeverExpires $False `
                -ChangePasswordAtLogon $true `
                -Enabled $true

                Start-Sleep -s 12
            }
            elseif($department -eq "ins"){
                #---------This creates the user----------------------- 
                New-ADUser `
                -Name "$FirstName $LastName" `
                -GivenName "$FirstName" `
                -Surname "$LastName" `
                -Path "*insert path*" `
                -Server "*insert servername*" `
                -UserPrincipalName "$checked_username@domain.com" `
                -SamAccountName  "$checked_username" `
                -DisplayName "$FirstName $LastName" `
                -AccountPassword (ConvertTo-SecureString "InsertTempPassword" -AsPlainText -Force) `
                -PasswordNeverExpires $False `
                -ChangePasswordAtLogon $true `
                -Enabled $true

                Start-Sleep -s 12
            }
            elseif($department -eq "sal"){
                #---------This creates the user----------------------- 
                New-ADUser `
                -Name "$FirstName $LastName" `
                -GivenName "$FirstName" `
                -Surname "$LastName" `
                -Path "*insert path*" `
                -Server "*insert servername*" `
                -UserPrincipalName "$checked_username@domain.com" `
                -SamAccountName  "$checked_username" `
                -DisplayName "$FirstName $LastName" `
                -AccountPassword (ConvertTo-SecureString "InsertTempPassword" -AsPlainText -Force) `
                -PasswordNeverExpires $False `
                -ChangePasswordAtLogon $true `
                -Enabled $true

                Start-Sleep -s 12
            }
            else{
                exit
            }

            Add-ADGroupMember -identity bostan -members $checked_username -Server "*insert servername*"
            Add-ADGroupMember -identity unsurance -members $checked_username -Server "*insert servername*"
            Add-ADGroupMember -identity ROINetbostan -members $checked_username -Server "*insert servername*"
            $Command.CommandText = "UPDATE *insert table* SET ad_addtimestamp = now() WHERE as4id = `"$username`""
            $RecordCount = $DataAdapter.Fill($dataset, "data")
        }
        else {
            break
        }     
    }   
    else{
        $redo = 1
    }
    
    $newchoice = Read-Host "Do you want to select another account (y/n)?"

    if($newchoice -eq "y" -or $newchoice -eq "yes"){
        $redo = 1
    }
    else{
        $redo = 0
    }
}
$connection.close()

