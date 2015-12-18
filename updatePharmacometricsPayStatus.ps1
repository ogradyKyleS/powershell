<#
    ***PLEASE NOTE***
        -This script relies on sensitive information which has been redacted for the purpose of distributing the code for review. 
         All fields relating to credentials or resource locations have been replaced with text stating "REDACTED".
    SCRIPT AUTHOR: KYLE O'GRADY
    SCRIPT DATE: 10/1/2015

    NOTICE: This script was created for use by the Unversity at Buffalo School of Pharmacy and Pharmaceutical Sciences Only
    PURPOSE: The purpose of this script is to update the payment status and amount charged for pharmacometrics registrants
    DETAILS: 
        -Requirements: 
            -This script requires Powershell Version 3 or higher.
            -This script requires MySql.Data and System.Data assemblies.
            -This script requires credentials encrypted with Windows Data Protection API
        - This script carries out the following operations in the following order:
            1) obtain the pharmacometrics form and all of its fields.
            2) obtain the number of pages of statusUpdates.
            3) obtain all of the submissions on the current page where the current status is attending.
            4) select only those submissions that are not for guest type registrants.
            5) foreach submission check all of the event fields to check if registered and all of the corresponding payment status fields to check if not paid, grab all registered not paid sessions.
            6) obtain UBF.Gateway records that correspond with the id's of the submissions obtained in step 5 above.
            7) foreach record returned from UBF lookup the corresponding submission and compare the app_id contents (holds the sessions paid for) with the pending payment sessions obtained above.
            8) foreach not paid session with a corresponding Gateway payment record build the update url and update the submission at formstack

        -UPDATES:
            UPDATE 10/13/2015
                -added ability to update the amount charged by submission id
            UPDATE 10/29/2015
                -added ability to use encrypted credentials from an external file
            UPDATE 12/18/2015
                -simplified process of creating form field dictionaries by combining all types into 1 loop instead of 1 loop per type
                -added error logging
                -added exit on fatal errors
                -altered return type of buildMySqlConnection function to return null on fail instead of false

        -FIXED:
            -fixed issue that was allowing a blank 'in clause' to be executed, resulting in error
            - 11/4/15
                -fixed issue that was grabbing incorrect value for payment fields for students and gov type due to an improper reference that searched for student or gov in the individual field dictionary.
#>

#function to load assemblies: takes 1 string arg, returns false on fail, true on success
function loadAssembly([string]$assemblyName){
    $result =[System.Reflection.Assembly]::LoadWithPartialName($assemblyName)
    $result -ne $null
}

#function to return a connection object for a mysql db using mysql.data.mysqlclient: has 4 string args, returns a mysql.data.mysqlclient.connection object on success, returns null on fail
function buildMySqlConnection([string]$server, [string]$user, [string]$pass, [string]$database){
    if($server.Length -ne 0 -and $user.Length -ne 0 -and $pass.Length -ne 0 -and $database.Length -ne 0){
        $connObj = New-Object MySql.Data.MySqlClient.MySqlConnection
        if($connObj -ne $null -and $connObj.GetType().Name -eq "MySqlConnection"){
            $connObj.connectionString = "server=$server;uid=$user;pwd=$pass;database=$database;"
        }
    }
    
    $connObj
}

#function to return a hash table containing key username => value decrypted password, reads a clixml file returning a hash table on success or null on fail
function Get-CredentialDictionary([string]$credentialPath){
    if((Test-Path -Path $credentialPath) -or (Test-Path -LiteralPath $credentialPath)){
        $cred = Import-Clixml $credentialPath
        $credHash = @{}
        $credHash.Add('UserName',$cred.UserName)
        $credHash.Add('Password',$cred.GetNetworkCredential().Password)
        $credHash
    }
    else{
        $null
    }
}

#function to create error logs
function out_ErrorLog([string]$exception, [string]$outPutErrorPath){
    #set time error occured
    $errTime = (Get-Date -Format "MM-dd-yy-HH:mm")
    #append time to error message and add newline incase of multiple writes
    $errLogEntry = $exception + " occured at: " + $errTime + "`r`n"
    #check if file exists
    if((Test-Path -Path $outPutErrorPath)){
        #file exists, append
        $errLogEntry | Out-File -FilePath $outPutErrorPath -Append
    }
    else{
        #file does not exist, create
        $errLogEntry | Out-File -FilePath $outPutErrorPath
    }
}

#set default error log path
$errorLogPath = "$PWD\errorLog.txt"

#load assemblies
if ((loadAssembly "MySql.Data") -eq $false){
    out_ErrorLog "MySql Assembly Failed To Load." $errorLogPath
    #fatal error
    exit
}
if ((loadAssembly "System.Data") -eq $false){
    out_ErrorLog "System.Data Assembly Failed To Load." $errorLogPath
    #fatal error
    exit
}

#auth info
$token = (Get-CredentialDictionary 'CREDENTIAL PATH REDACTED').Password
#if the credential is not loaded then exit the script
if($token -eq $null){
    out_ErrorLog "The Credential Was Null." $errorLogPath
    #fatal error
    exit
}

$baseUrl = "DATA URL REDACTED"

#build the http request header as a hash table
$header = @{}
$header.Add('Host','www.formstack.com')
$header.Add('Authorization', 'Bearer ' + $token)
$header.Add('Accept','application/json')
$header.Add('Content-Type','application/json')

#get all forms
$uri=$baseUrl + "/form.json"
$restGetAllForms = Invoke-RestMethod -Uri $uri -Headers $header -Method Get

#get the pharmacometrics form
$pharmacometricsForm = $restGetAllForms.forms | ? {$_.name -eq "Buffalo Pharmacometrics Workshops"}

#get the fields of the pharmacometrics form
$uri = $baseUrl + "/form/" + $pharmacometricsForm.id
$restPharmacoFormFields = Invoke-RestMethod -Uri $uri -Headers $header -Method Get

#get the registrant type field
$fieldRegistrationType = $restPharmacoFormFields.fields | ? {$_.name -eq "registrant_type"}

#get the current status field
$fieldCurrentStatus = $restPharmacoFormFields.fields | ? {$_.name -eq "current_status"}

#get the amount charged field
$fieldAmountCharged = $restPharmacoFormFields.fields | ? {$_.name -eq "amount_charged"}

#get the payment fields
$fieldsPaymentStatus = $restPharmacoFormFields.fields | ? {$_.name -like "payment_status*"}

#make a hash of session number => payment field
$sessionNumberPaymentFieldDictionary = @{}
$sessionStatusField = $fieldsPaymentStatus | ?{$_.name -eq 'payment_status_1'}
$sessionNumberPaymentFieldDictionary.Add('1',$sessionStatusField.id)
$sessionStatusField = $fieldsPaymentStatus | ?{$_.name -eq 'payment_status_2'}
$sessionNumberPaymentFieldDictionary.Add('2',$sessionStatusField.id)
$sessionStatusField = $fieldsPaymentStatus | ?{$_.name -eq 'payment_status_3'}
$sessionNumberPaymentFieldDictionary.Add('3',$sessionStatusField.id)
$sessionStatusField = $fieldsPaymentStatus | ?{$_.name -eq 'payment_status_4'}
$sessionNumberPaymentFieldDictionary.Add('4',$sessionStatusField.id)
$sessionStatusField = $fieldsPaymentStatus | ?{$_.name -eq 'payment_status_5'}
$sessionNumberPaymentFieldDictionary.Add('5',$sessionStatusField.id)

#get all individual session fields
$fieldsIndividualSessions = $restPharmacoFormFields.fields | ? {$_.name -like "session_*_individual"}

#make Individual reg type hash
$regTypeIndividualDictionary = @{}

#populate individual hash
for($fieldNumber=1; $fieldNumber -le $fieldsIndividualSessions.Count; $fieldNumber++){
    
    $sessionField = $fieldsIndividualSessions.Where({$_.name -eq -join('session_',$fieldNumber,'_individual')})
    $paymentField = $fieldsPaymentStatus | ?{$_.name -eq "payment_status_$fieldNumber"}
    $regTypeIndividualDictionary.Add($sessionField.id,$paymentField.id)
}


#get all gov session fields
$fieldsGovernmentSessions = $restPharmacoFormFields.fields | ? {$_.name -like "session_*_government"}

#make gov reg type hash
$regTypeGovDictionary = @{}

#populate gov hash
for($fieldNumber=1; $fieldNumber -le $fieldsGovernmentSessions.Count; $fieldNumber++){
    
    $sessionField = $fieldsGovernmentSessions.Where({$_.name -eq -join('session_',$fieldNumber,'_government')})
    $paymentField = $fieldsPaymentStatus | ?{$_.name -eq "payment_status_$fieldNumber"}
    $regTypeGovDictionary.Add($sessionField.id,$paymentField.id)
}


#get all student session fields
$fieldsStudentSessions = $restPharmacoFormFields.fields | ? {$_.name -like "session_*_student"}

#make student reg type hash
$regTypeStudentDictionary = @{}

#populate student hash
for($fieldNumber=1; $fieldNumber -le $fieldsStudentSessions.Count; $fieldNumber++){
    
    $sessionField = $fieldsStudentSessions.Where({$_.name -eq -join('session_',$fieldNumber,'_student')})
    $paymentField = $fieldsPaymentStatus | ?{$_.name -eq "payment_status_$fieldNumber"}
    $regTypeStudentDictionary.Add($sessionField.id,$paymentField.id)
}

<#
    Obtain the total number of pages of statusUpdates.
    restrict to attending only.
#>
$uri =$baseUrl + "/form/" + $pharmacometricsForm.id + "/submission.json"
$body = @{}
$body.Add('search_field_0',$currentStatusField.id)
$body.Add('search_value_0','Attending')
$body.Add('per_page','25')
$body.Add('data','false')
$restGetNumberOfPages = Invoke-RestMethod -Uri $uri -Headers $header -Body $body -Method Get

#remove data = false from request body and replace with data = true
$body.Remove('data')
$body.Add('data','true')
<#
    foreach page send a request to obtain all statusUpdates on that page number.
    foreach result grab the submissions that are not from guests
#>
$nonGuestSubmissions = @()
for($currentPage = 1; $currentPage -le $restGetNumberOfPages.pages; $currentPage++){
    $body.Add('page',$currentPage)
    $restGetAllSubmissionsOnPage = Invoke-RestMethod -Uri $uri -Headers $header -Body $body -Method Get
    #get only those submissions that are not guests
    $nonGuestSubmissions += $restGetAllSubmissionsOnPage.submissions | ?{$_.data.($fieldRegistrationType.id).value -ne 'Guest'}
    $body.Remove('page')
}

#foreach nonGuestSubmission find each session that is selected but listed as Not Paid and add to dictionary
$submissionsRequiringUpdatesDictionary = @{}

if($nonGuestSubmissions.Count -gt 0){
    $regTypes = @('Individual','Government','Registered MS or PhD')
    $regTypes.ForEach({
        $currentType = $_
        $submissionsOfCurrentType = $nonGuestSubmissions.Where({$_.data.($fieldRegistrationType.id).value -eq $currentType})
        if($submissionsOfCurrentType.count -gt 0){
            $submissionsOfCurrentType.ForEach({
                $currentSubmissionRecord = $_
                $submissionFieldsToUpdate = @()
                if($currentSubmissionRecord.data.($fieldRegistrationType.id).value -eq 'Individual'){
                    $regTypeIndividualDictionary.GetEnumerator().ForEach({
                        $sessionValue = $currentSubmissionRecord.data.($_.name).value | ConvertFrom-StringData
                        if($sessionValue.quantity -eq '1' -and $currentSubmissionRecord.data.($_.value).value -eq 'Not Paid'){
                            #field is reg and not paid, need to add to update list
                           $submissionFieldsToUpdate += $_.value
                        }
                    })
                    $paymentFieldsString = [System.String]::Join(',',$submissionFieldsToUpdate)
                    if($submissionFieldsToUpdate.Count -gt 0){
                        $submissionsRequiringUpdatesDictionary.Add($_.id,$paymentFieldsString)
                    }
                }
                elseif($currentSubmissionRecord.data.($fieldRegistrationType.id).value -eq 'Government'){
                    $regTypeGovDictionary.GetEnumerator().ForEach({
                        $sessionValue = $currentSubmissionRecord.data.($_.name).value | ConvertFrom-StringData
                        if($sessionValue.quantity -eq '1' -and $currentSubmissionRecord.data.($_.value).value -eq 'Not Paid'){
                            #field is reg and not paid, need to add to update list
                           $submissionFieldsToUpdate += $_.value
                        }
                    })
                    $paymentFieldsString = [System.String]::Join(',',$submissionFieldsToUpdate)
                    if($submissionFieldsToUpdate.Count -gt 0){
                        $submissionsRequiringUpdatesDictionary.Add($_.id,$paymentFieldsString)
                    }
                }
                elseif($currentSubmissionRecord.data.($fieldRegistrationType.id).value -eq 'Registered MS or PhD'){
                    $regTypeStudentDictionary.GetEnumerator().ForEach({
                        $sessionValue = $currentSubmissionRecord.data.($_.name).value | ConvertFrom-StringData
                        if($sessionValue.quantity -eq '1' -and $currentSubmissionRecord.data.($_.value).value -eq 'Not Paid'){
                            #field is reg and not paid, need to add to update list
                           $submissionFieldsToUpdate += $_.value
                        }
                    })
                    $paymentFieldsString = [System.String]::Join(',',$submissionFieldsToUpdate)
                    if($submissionFieldsToUpdate.Count -gt 0){
                        $submissionsRequiringUpdatesDictionary.Add($_.id,$paymentFieldsString)
                    }
                }
            })
        }
    })
}

#get mysql related info
$credentialDictionary = Get-CredentialDictionary 'CREDENTIAL PATH REDACTED'
if($credentialDictionary -eq $null){
    out_ErrorLog "The Credential Is Null." $errorLogPath
    #fatal error
    exit
}
$mySqlConnection = buildMySqlConnection "DATABASE SERVER REDACTED" $credentialDictionary.UserName $credentialDictionary.Password "DATABASE NAME REDACTED"
if($mySqlConnection -eq $null){
    out_ErrorLog "The Connection Object Is Null." $errorLogPath
    #fatal error
    exit
}

$inClause = ''
#if any fields were added to the dictionary then query ubf table to determine if any of those fields have been paid.
if($submissionsRequiringUpdatesDictionary.Count -gt 0){
    #build in clause
    $idArray = @()
    $submissionsRequiringUpdatesDictionary.GetEnumerator().ForEach({
        $idArray += $_.name
    })

    $inClause = [System.String]::Join(',',$idArray)

    #query UBF.Gateway to obtain payment info for the id's collected above
    $sqlString = "SELECT COLUMN NAMES REDACTED FROM TABLE NAME REDACTED WHERE COLUMN NAME REDACTED IN ($inClause) AND COLUMN NAME REDACTED = UNIQUE ID REDACTED AND COLUMN NAME REDACTED = 1 AND COLUMN NAME REDACTED NOT LIKE '%TESTMODE%'"
    $mySqlCommandObject = New-Object MySql.Data.MySqlClient.MySqlCommand
    $sqlAdapter = New-Object MySql.Data.MySqlClient.MySqlDataAdapter
    $dataSet = New-Object System.Data.DataSet
    $dataSet.Tables.Add('results')
    $mySqlCommandObject.Connection = $mySqlConnection
    $mySqlCOmmandObject.CommandText = $sqlString
    $sqlAdapter.SelectCommand = $mySqlCommandObject
    $mySqlConnection.Open()
    $sqlAdapter.Fill($dataSet.Tables['results'])
    $mySqlConnection.Close()

    $dataSet.Tables['results'].Rows.ForEach({
        $currentRow = $_
        #get app_id and break apart on comma into array
        $rawFieldValue = $_.app_id
        $sessionPaymentDictionaryKeys = $rawFieldValue.Split(',')
        $verifiedUpdateArray = @()
        $sessionPaymentDictionaryKeys.foreach({
            #look up the corresponding payment status field and get the id
            $correspondingPaymentStatusField = $sessionNumberPaymentFieldDictionary.$_
            #use the app_session (submission id) of the current row to look up the corresponding submissions to update hash row
            $correspondingSubmissionDictionaryRow = $submissionsRequiringUpdatesDictionary.Item($currentRow.app_session)
            #check if the submission hash row contains the field id looked up above
            if($correspondingSubmissionDictionaryRow.Contains($correspondingPaymentStatusField)){
                #add field id to verified update array
                $verifiedUpdateArray += $correspondingPaymentStatusField
            }
        })
        #create an update uri using the field id's in the verified array
        if($verifiedUpdateArray.Count -gt 0){
            $body = @{}
            $verifiedUpdateArray.ForEach({
                $body.Add("field_$_",'Paid')
            })
            #do update
            $uri = $baseUrl + "/submission/" + $currentRow.app_session + ".json"
            $updateAttempt = Invoke-RestMethod -Uri $uri -Headers $header -Body $body -Method Put
        }
    })
}

#query UBF.Gateway to obtain total amount for each submission id, only if inClause is NOT NULL
if($inClause -ne $null -and $inClause.Length -ne 0){
    $sqlString = "SELECT COLUMN NAMES REDACTED FROM TABLE NAME REDACTED WHERE COLUMN NAME REDACTED IN ($inClause) AND COLUMN NAME REDACTED = UNIQUE ID REDACTED AND COLUMN NAME REDACTED = 1 AND COLUMN NAME REDACTED NOT LIKE '%TESTMODE%' GROUP BY COLUMN NAME REDACTED"
    $mySqlCommandObject = New-Object MySql.Data.MySqlClient.MySqlCommand
    $sqlAdapter = New-Object MySql.Data.MySqlClient.MySqlDataAdapter
    $dataSet = New-Object System.Data.DataSet
    $dataSet.Tables.Add('results')
    $mySqlCommandObject.Connection = $mySqlConnection
    $mySqlCOmmandObject.CommandText = $sqlString
    $sqlAdapter.SelectCommand = $mySqlCommandObject
    $mySqlConnection.Open()
    $sqlAdapter.Fill($dataSet.Tables['results'])
    $mySqlConnection.Close()

    #update the amount paid fields
    $dataSet.Tables['results'].Rows.ForEach({
        $body = @{}
        $body.Add("field_" + $amountChargedField.id,$_.total_amount)
        $uri = $baseUrl + "/submission/" + $_.app_session + ".json"
        $updateAttempt = Invoke-RestMethod -Uri $uri -Headers $header -Body $body -Method Put
    })
}

if($dataSet -ne $null){
    $dataSet.Dispose()
}
if($mySqlConnection -ne $null){
    $mySqlConnection.Dispose()
}
