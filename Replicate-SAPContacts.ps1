<#
.SYNOPSIS
  Replicates SAP Business Partners to Exchange Online Contacts
.DESCRIPTION
  Replicates SAP Business Partners to Exchange Online Contacts
.INPUTS
  None
.OUTPUTS
  None
.NOTES
  Version:        1.2
  Author:         Tobias Meier
  Creation Date:  05.02.2022
  Purpose/Change: e.164 conversion fixed
  
.EXAMPLE
  Replicate-SAPContacts.ps1
#>

#---------------------------------------------------------[Initialisations]--------------------------------------------------------

#Set Error Action to silently continue
$ErrorActionPreference = "SilentlyContinue"

#Set this to enable if you want to delete all created contacts
$deleteAllContacts=$false

#SQL server variables
$SQLServerName=""
$SQLDataBase=""

<#
#SQL query
get all Business partners(table OCPR) and the company(table OCRD)
exclude contacts without phone number and email address

SAP Business One tables https://blogs.sap.com/2017/04/27/list-of-object-types/
may you want to replace SELECT with SELECT TOP 5 for testing / first run
#>
$SQLQuery='
SELECT dbo.OCPR.CntctCode as "id",dbo.OCPR.Name as "name", dbo.OCPR.FirstName as firstname, dbo.OCPR.LastName as lastname, dbo.OCPR.E_MailL as mail, dbo.OCRD.CardName as company, dbo.OCPR.Tel1 as primaryphone, dbo.OCPR.Tel2 as secondaryphone, dbo.OCPR.Cellolar as mobile
FROM dbo.OCPR
INNER JOIN dbo.OCRD ON dbo.OCPR.CardCode=dbo.OCRD.CardCode
WHERE (dbo.OCPR.E_MailL IS NOT NULL AND DATALENGTH(dbo.OCPR.E_MailL) > 0)
AND 
(dbo.OCPR.Tel1 IS NOT NULL OR dbo.OCPR.Tel2 IS NOT NULL OR dbo.OCPR.Cellolar IS NOT NULL);
'

#-----------------------------------------------------------[Functions]------------------------------------------------------------

function Convert-PhoneNumber {
    [CmdletBinding()]
    param(
        [Parameter()]
        [string]$Number
    )
    

    #Case 079 or 044: Convert to +41 79
    if ( $Number -match "^0[1-9][0-9].*" ) {
        $NewNumber=$Number -replace '^0','+41 '
    }
    #Case 0044 or 0041: Convert to +44
    elseif ( $Number -match "^00[1-9][1-9].*" ) {
        $NewNumber=$Number -replace '^00','+'
    }

    else {
        $NewNumber = $Number
    }

    return $NewNumber
}

function Delete-Contacts { 
    foreach ( $Contact in $SQLContacts ) {
        $ContactToRemove = Get-Contact | Where {$_.Notes -eq $Contact.ID -or $_.DisplayName -eq $Contact.DisplayName}
        Remove-MailContact -Identity "$ContactToRemove" -Confirm:$false -ErrorAction SilentlyContinue
        Write-Log -Severity Information -Message "Deleting Contact $ContactToRemove"
    }
}

function Write-Log {
    [CmdletBinding()]
    param(
        [Parameter()]
        [ValidateNotNullOrEmpty()]
        [string]$Message,
 
        [Parameter()]
        [ValidateNotNullOrEmpty()]
        [ValidateSet('Information','Warning','Error')]
        [string]$Severity = 'Information'
    )
 
    [pscustomobject]@{
        Time = (Get-Date -f g)
        Message = $Message
        Severity = $Severity
    } | Export-Csv -Path "Replicate-SAPContacts-Log.csv" -Append -NoTypeInformation
 }

#----------------------------------------------------------[Declarations]----------------------------------------------------------

<#
##### data structure considerations #####

Unique attribute (primary key) on SQL is dbo.OCPR.CntctCode
On EXO email and Identity must be unique
This script does not use a proper synchronization anchor there is a possibility that entries lose the synchronization anchor manual clean-up is needed then
(To do it properly you have to use a separate data store (for example .csv file) and write unique IDs in there...)

Synchronization anchor design
SQL primary key is written in the notes property on EXO. Only object with a text in notes will be recognized. If you use notes for original purpose script will not work...
#>

$SQLContacts = @()
$EXOContacts = @()

#Contact data structure
class Contact
{
    [string]$ID
    [string]$FirstName
    [string]$LastName
    [string]$DisplayName
    [string]$Company
    [string]$EMail
    [string]$Phone
    [string]$Mobile
    [string]$SecondaryPhone

    Contact($ID, $FirstName, $LastName, $Company, $EMail, $Phone, $Mobile, $SecondaryPhone) {
        $this.ID=$ID
        $this.FirstName=$FirstName
        $this.LastName=$LastName
        $this.Company=$Company
        $this.EMail=$EMail
        $this.Phone=$Phone
        $this.Mobile=$Mobile
        $this.SecondaryPhone=$SecondaryPhone

        $this.DisplayName = "[$($this.Company)] - $($this.FirstName) $($this.LastName)"

        #Normalize data

        #Replace "00xx" with "+xx"
        $this.Phone=$this.Phone -replace '^00','+'
        $this.Mobile=$this.Mobile -replace '^00','+'
        $this.SecondaryPhone=$this.SecondaryPhone -replace '^00','+'

        #Normalize phone number to E.164
        $this.Phone=$(Convert-PhoneNumber -Number $this.Phone)
        $this.Mobile=$(Convert-PhoneNumber -Number $this.Mobile)
        $this.SecondaryPhone=$(Convert-PhoneNumber -Number $this.SecondaryPhone)

        #Remove excessive text after mail address
        $this.EMail=$this.EMail -replace '^[a-zA-Z0-9_.+-]+@[a-zA-Z0-9-]+\.[a-zA-Z0-9-.]+$','$0'

        #Set Displayname max size to 64 and trim the string (limitations of EXO)
        $this.DisplayName = $($this.DisplayName[0..(64-1)] -join "")
        $this.DisplayName = $this.DisplayName.Trim()

    }
}

 #-----------------------------------------------------------[Execution]------------------------------------------------------------

Write-Log -Severity Information -Message "Script started"

try {
    Connect-ExchangeOnline -CertificateThumbPrint "1e3e53810ed235e26396cbef64a443145f06428d" -AppID "7d11008b-b6db-4920-b59b-d13349eace8a" -Organization "mtfdata.ch"
    Write-Log -Severity Information -Message "Connected to Exchange Online"
}
catch {
    Write-Log -Severity Error -Message "Could not connect to EXO"
    Write-Log -Severity Error -Message $_.Exception
    return
}
$getsessions = Get-PSSession | Select-Object -Property State, Name
$isconnected = (@($getsessions) -like '@{State=Opened; Name=ExchangeOnlineInternalSession*').Count -gt 0
if ($isconnected -ne "True") {
    Write-Log -Severity Error -Message "Could not connect to EXO"
    return
}

#Run SQL query get SQL Contacts
try {
    $SQLContactsRaw = Invoke-Sqlcmd -ServerInstance $SQLServerName -Database $SQLDataBase -Query $SQLQuery -OutputSqlErrors $true -ErrorVariable invokeSqlError
    Write-Log -Severity Information -Message "Loaded SQL data"
}
catch {
    Write-Log -Severity Error -Message "Could not connect to SQL Server"
    Write-Log -Severity Error -Message $_.Exception
    Write-Log -Severity Error -Message $invokeSqlError
    return
}

#Type conversion Contacts class 
$SQLContactsRaw | ForEach-Object { 
    #Since e-mail address must be unique create the object only if the e-mail adress is unique...
    $mail = $_.mail
    if ( $SQLContacts.EMail -notcontains $mail ) {
        $SQLContacts+=$(New-Object -TypeName Contact -ArgumentList $_.id,$_.firstname,$_.lastname,$_.company,$_.mail,$_.primaryphone,$_.mobile,$_.secondaryphone)
    }
}

#Get Mailcontacts
$EXOContactsRaw = Get-Contact -ResultSize unlimited | select * | Where { $_.Notes -ne "" }
Write-Log -Severity Information -Message "Loaded Exchange Online data"
#Type conversion Contacts class 
$EXOContactsRaw | ForEach-Object { $EXOContacts+=$(New-Object -TypeName Contact -ArgumentList $_.Notes,$_.FirstName,$_.LastName,$_.Company,$_.WindowsEmailAddress,$_.Phone,$_.MobilePhone,$_.HomePhone) }

#Compare contacts get contacts to recreate
$ContactsToCreateOrDelete = Compare-Object -ReferenceObject $SQLContacts -DifferenceObject $EXOContacts -Property ID
#Compare contacts get contacts to update
$ContactsToUpdate = Compare-Object -ReferenceObject $SQLContacts -DifferenceObject $EXOContacts -Property DisplayName,FirstName,LastName,Company,EMail,Phone,Mobile,SecondaryPhone

#delete all contacts and stop
if ( $deleteAllContacts ) {
    Delete-Contacts
    return
}

#Create contacts
foreach ($Contact in $($ContactsToCreateOrDelete | Where {$_.SideIndicator -eq "<="})) {

    #Since contact object consists only of ID overwrite it with all properties
    $Contact=$($SQLContacts | Where {$_.ID -eq $Contact.ID})

    #New-MailContact -Name $Contact.DisplayName -DisplayName $Contact.DisplayName -ExternalEmailAddress $Contact.EMail -FirstName $Contact.FirstName -LastName $Contact.LastName
    Set-Contact -Identity $Contact.DisplayName -Notes $Contact.ID
    
    Write-Log -Severity Information -Message "Creating Contact $($Contact.DisplayName)"
}

#Delete contacts
foreach ($Contact in $($ContactsToCreateOrDelete | Where {$_.SideIndicator -eq "=>"})) {

    #Since contact object consists only of ID overwrite it with all properties
    $Contact=$($EXOContacts | Where {$_.ID -eq $Contact.ID})

    Remove-MailContact -Identity $Contact.DisplayName -Confirm:$false -ErrorAction SilentlyContinue
    
    Write-Log -Severity Information -Message "Deleting Contact $($Contact.DisplayName)"
}

#Update contacts
foreach ($Contact in $($ContactsToUpdate | Where {$_.SideIndicator -eq "<="})) {


    #Technical we do not recompare the ID with the only object, this could lead to problem
    Set-Contact -Identity $Contact.DisplayName -Company $Contact.Company -WindowsEmailAddress $Contact.EMail -Phone $Contact.Phone -MobilePhone $Contact.Mobile -HomePhone $Contact.SecondaryPhone
    Write-Log -Severity Information -Message "Updating Contact $($Contact.DisplayName)"
}

Write-Log -Severity Information -Message "Script ended"