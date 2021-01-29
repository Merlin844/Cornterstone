. ./dslSql.ps1

function AldridgeQualificationPrograms() {
  Get-ChildItem -Path "AldridgeQualificationPrograms" | % {
    $file = $_
    $nameParts = $file.Name -split "_"
    $dateParts = $nameParts[-5..-1]
    $date = $dateParts[0]
    $runDate = Get-Date ([datetime]::parseexact($date, "yyyyMMdd", [System.Globalization.CultureInfo]::InvariantCulture)).Date -f "yyyy-MM-dd"
    $records = Import-Csv -Path $("AldridgeQualificationPrograms/" + $file.Name)

    dslSql -ConnString "Data Source=sql.aldridge-electric.com;Integrated Security=SSPI;Initial Catalog=Aspn" {
      CmdConfig -Type "sp"  -Sql "cornerstone.sp_AldridgeQualificationPrograms"
  
      ClearParams
      ParamConfig -Param "RunDate" -SqlType ([System.Data.SqlDbType]::Date) -Value $runDate
      ParamConfig -Param "UserID" -SqlType ([System.Data.SqlDbType]::VarChar) -Value ""
      ParamConfig -Param "UserFullName" -SqlType ([System.Data.SqlDbType]::VarChar) -Value ""
      ParamConfig -Param "Division" -SqlType ([System.Data.SqlDbType]::VarChar) -Value ""
      ParamConfig -Param "UserEmail" -SqlType ([System.Data.SqlDbType]::VarChar) -Value ""
      ParamConfig -Param "UserMobileNumber" -SqlType ([System.Data.SqlDbType]::VarChar) -Value ""
      ParamConfig -Param "AldridgeQualificationProgramsQualification" -SqlType ([System.Data.SqlDbType]::VarChar) -Value ""
      ParamConfig -Param "CLEARDATE" -SqlType ([System.Data.SqlDbType]::Bit) -Value 1
      $a = ExecuteNonQuery
  
      $records | % {
        $record = $_
        $qualifications = $record."Aldridge Qualification Programs - Qualification" -split "," 
        $qualifications | % {
          $qualification = $_
          ClearParams
          ParamConfig -Param "RunDate" -SqlType ([System.Data.SqlDbType]::Date) -Value $runDate
          ParamConfig -Param "UserID" -SqlType ([System.Data.SqlDbType]::VarChar) -Value $record."User ID"
          ParamConfig -Param "UserFullName" -SqlType ([System.Data.SqlDbType]::VarChar) -Value $record."User Full Name"
          ParamConfig -Param "Division" -SqlType ([System.Data.SqlDbType]::VarChar) -Value $record."Division"
          ParamConfig -Param "UserEmail" -SqlType ([System.Data.SqlDbType]::VarChar) -Value $record."User Email"
          ParamConfig -Param "UserMobileNumber" -SqlType ([System.Data.SqlDbType]::VarChar) -Value $record."User Mobile Number"
          ParamConfig -Param "AldridgeQualificationProgramsQualification" -SqlType ([System.Data.SqlDbType]::VarChar) -Value $qualification
          ParamConfig -Param "CLEARDATE" -SqlType ([System.Data.SqlDbType]::Bit) -Value 0
          $a = ExecuteNonQuery
        }
      }
    }

    write-host $file
    $file | Move-Item -Destination "\\001sql02\e$\integrations\cornerstone\Archive\AldridgeQualificationPrograms" -Force
  }
}

function AldridgeSubjectMatterExperts() {
  Get-ChildItem -Path "AldridgeSubjectMatterExperts" | % {
    $file = $_
    $nameParts = $file.Name -split "_"
    $dateParts = $nameParts[-5..-1]
    $date = $dateParts[0]
    $runDate = Get-Date ([datetime]::parseexact($date, "yyyyMMdd", [System.Globalization.CultureInfo]::InvariantCulture)).Date -f "yyyy-MM-dd"
    $records = Import-Csv -Path $("AldridgeSubjectMatterExperts/" + $file.Name)
    dslSql -ConnString "Data Source=sql.aldridge-electric.com;Integrated Security=SSPI;Initial Catalog=Aspn" {
      CmdConfig -Type "sp"  -Sql "cornerstone.sp_AldridgeSubjectMatterExperts"
  
      ClearParams
      ParamConfig -Param "RunDate" -SqlType ([System.Data.SqlDbType]::Date) -Value $runDate
      ParamConfig -Param "UserID" -SqlType ([System.Data.SqlDbType]::VarChar) -Value ""
      ParamConfig -Param "UserFullName" -SqlType ([System.Data.SqlDbType]::VarChar) -Value ""
      ParamConfig -Param "Division" -SqlType ([System.Data.SqlDbType]::VarChar) -Value ""
      ParamConfig -Param "UserEmail" -SqlType ([System.Data.SqlDbType]::VarChar) -Value ""
      ParamConfig -Param "UserMobileNumber" -SqlType ([System.Data.SqlDbType]::VarChar) -Value ""
      ParamConfig -Param "CompanyIdentifiedExpertSkill" -SqlType ([System.Data.SqlDbType]::VarChar) -Value ""
      ParamConfig -Param "CLEARDATE" -SqlType ([System.Data.SqlDbType]::Bit) -Value 1
      $a = ExecuteNonQuery
  
      $records | % {
        $record = $_
        $sme = $record."Company Identified Expert - Skill " -split "," 
        $sme | % {
          $skill = $_
          ClearParams
          ParamConfig -Param "RunDate" -SqlType ([System.Data.SqlDbType]::Date) -Value $runDate
          ParamConfig -Param "UserID" -SqlType ([System.Data.SqlDbType]::VarChar) -Value $record."User ID"
          ParamConfig -Param "UserFullName" -SqlType ([System.Data.SqlDbType]::VarChar) -Value $record."User Full Name"
          ParamConfig -Param "Division" -SqlType ([System.Data.SqlDbType]::VarChar) -Value $record."Division"
          ParamConfig -Param "UserEmail" -SqlType ([System.Data.SqlDbType]::VarChar) -Value $record."User Email"
          ParamConfig -Param "UserMobileNumber" -SqlType ([System.Data.SqlDbType]::VarChar) -Value $record."User Mobile Number"
          ParamConfig -Param "CompanyIdentifiedExpertSkill" -SqlType ([System.Data.SqlDbType]::VarChar) -Value $skill
          ParamConfig -Param "CLEARDATE" -SqlType ([System.Data.SqlDbType]::Bit) -Value 0
          $a = ExecuteNonQuery
        }
      }
    }

   $file | Move-Item -Destination "\\001sql02\e$\integrations\cornerstone\Archive\AldridgeSubjectMatterExperts" -Force
  }
}

function CompanyResources() {
  Get-ChildItem -Path "CompanyResources" | % {
    $file = $_
    $nameParts = $file.Name -split "_"
    $dateParts = $nameParts[-5..-1]
    $date = $dateParts[0]
    $runDate = Get-Date ([datetime]::parseexact($date, "yyyyMMdd", [System.Globalization.CultureInfo]::InvariantCulture)).Date -f "yyyy-MM-dd"
    $records = Import-Csv -Path $("CompanyResources/" + $file.Name)
    
    dslSql -ConnString "Data Source=sql.aldridge-electric.com;Integrated Security=SSPI;Initial Catalog=Aspn" {
      CmdConfig -Type "sp"  -Sql "cornerstone.sp_CompanyResources"
  
      ClearParams
      ParamConfig -Param "RunDate" -SqlType ([System.Data.SqlDbType]::Date) -Value $runDate
      ParamConfig -Param "UserID" -SqlType ([System.Data.SqlDbType]::VarChar) -Value ""
      ParamConfig -Param "UserFullName" -SqlType ([System.Data.SqlDbType]::VarChar) -Value ""
      ParamConfig -Param "Division" -SqlType ([System.Data.SqlDbType]::VarChar) -Value ""
      ParamConfig -Param "UserEmail" -SqlType ([System.Data.SqlDbType]::VarChar) -Value ""
      ParamConfig -Param "UserMobileNumber" -SqlType ([System.Data.SqlDbType]::VarChar) -Value ""
      ParamConfig -Param "CompanyResourceResponsibility" -SqlType ([System.Data.SqlDbType]::VarChar) -Value ""
      ParamConfig -Param "PhoneNumber" -SqlType ([System.Data.SqlDbType]::VarChar) -Value ""
      ParamConfig -Param "CLEARDATE" -SqlType ([System.Data.SqlDbType]::Bit) -Value 1
      $a = ExecuteNonQuery
  
      $records | % {
        $record = $_
        $responsiblilities = $record."Company Resource - Responsibility" -split "," 
        $responsiblilities | % {
          $responsibility = $_
          ClearParams
          ParamConfig -Param "RunDate" -SqlType ([System.Data.SqlDbType]::Date) -Value $runDate
          ParamConfig -Param "UserID" -SqlType ([System.Data.SqlDbType]::VarChar) -Value $record."User ID"
          ParamConfig -Param "UserFullName" -SqlType ([System.Data.SqlDbType]::VarChar) -Value $record."User Full Name"
          ParamConfig -Param "Division" -SqlType ([System.Data.SqlDbType]::VarChar) -Value $record."Division"
          ParamConfig -Param "UserEmail" -SqlType ([System.Data.SqlDbType]::VarChar) -Value $record."User Email"
          ParamConfig -Param "UserMobileNumber" -SqlType ([System.Data.SqlDbType]::VarChar) -Value $record."User Mobile Number"
          ParamConfig -Param "CompanyResourceResponsibility" -SqlType ([System.Data.SqlDbType]::VarChar) -Value $responsibility
          ParamConfig -Param "PhoneNumber" -SqlType ([System.Data.SqlDbType]::VarChar) -Value $record."User Phone Number"
          ParamConfig -Param "CLEARDATE" -SqlType ([System.Data.SqlDbType]::Bit) -Value 0
          $a = ExecuteNonQuery
        }
      }
    }

    $file | Move-Item -Destination "\\001sql02\e$\integrations\cornerstone\Archive\CompanyResources" -Force
  }
}
function ProfessionalCertificationsLicenses() {
  Get-ChildItem -Path "ProfessionalCertificationsLicenses" | % {
    $file = $_
    $nameParts = $file.Name -split "_"
    $dateParts = $nameParts[-5..-1]
    $date = $dateParts[0]
    $runDate = Get-Date ([datetime]::parseexact($date, "yyyyMMdd", [System.Globalization.CultureInfo]::InvariantCulture)).Date -f "yyyy-MM-dd"
    $records = Import-Csv -Path $("ProfessionalCertificationsLicenses/" + $file.Name)
    
    dslSql -ConnString "Data Source=sql.aldridge-electric.com;Integrated Security=SSPI;Initial Catalog=Aspn" {
      CmdConfig -Type "sp"  -Sql "cornerstone.sp_ProfessionalCertificationsLicenses"
  
      ClearParams
      ParamConfig -Param "RunDate" -SqlType ([System.Data.SqlDbType]::Date) -Value $runDate
      ParamConfig -Param "UserID" -SqlType ([System.Data.SqlDbType]::VarChar) -Value ""
      ParamConfig -Param "UserFullName" -SqlType ([System.Data.SqlDbType]::VarChar) -Value ""
      ParamConfig -Param "Division" -SqlType ([System.Data.SqlDbType]::VarChar) -Value ""
      ParamConfig -Param "UserEmail" -SqlType ([System.Data.SqlDbType]::VarChar) -Value ""
      ParamConfig -Param "UserMobileNumber" -SqlType ([System.Data.SqlDbType]::VarChar) -Value ""
      ParamConfig -Param "ProfessionalCertificationsLicensesCertificationLicense" -SqlType ([System.Data.SqlDbType]::VarChar) -Value ""
      ParamConfig -Param "ProfessionalCertificationsLicensesDateReceived" -SqlType ([System.Data.SqlDbType]::VarChar) -Value ""
      ParamConfig -Param "ProfessionalCertificationsLicensesDetails" -SqlType ([System.Data.SqlDbType]::VarChar) -Value ""
      ParamConfig -Param "CLEARDATE" -SqlType ([System.Data.SqlDbType]::Bit) -Value 1
      $a = ExecuteNonQuery
  
      $records | % {
        $record = $_
        $knt = 0
        $certifications = $record."Professional Certifications/Licenses - Certification/License" -split ","
        $certifications | % {
          $cert = $_
          $certDate = $($record."Professional Certifications/Licenses - Date Received" -split ",")
          $certDetails = $($record."Professional Certifications/Licenses - Details" -split ",")
          if($null -eq $certDate) { $certDate = "" }
          if($null -eq $certDetails) { $certDetails = "" }

          if($certDate.GetType().Name -eq "String") { $certDate = $certDate } else { $certDate = $certDate[$knt] }
          if($certDetails.GetType().Name -eq "String") { $certDetails = $certDetails } else { $certDetails = $certDetails[$knt] }

          ClearParams
          ParamConfig -Param "RunDate" -SqlType ([System.Data.SqlDbType]::Date) -Value $runDate
          ParamConfig -Param "UserID" -SqlType ([System.Data.SqlDbType]::VarChar) -Value $record."User ID"
          ParamConfig -Param "UserFullName" -SqlType ([System.Data.SqlDbType]::VarChar) -Value $record."User Full Name"
          ParamConfig -Param "Division" -SqlType ([System.Data.SqlDbType]::VarChar) -Value $record."Division"
          ParamConfig -Param "UserEmail" -SqlType ([System.Data.SqlDbType]::VarChar) -Value $record."User Email"
          ParamConfig -Param "UserMobileNumber" -SqlType ([System.Data.SqlDbType]::VarChar) -Value $record."User Mobile Number"
          ParamConfig -Param "ProfessionalCertificationsLicensesCertificationLicense" -SqlType ([System.Data.SqlDbType]::VarChar) -Value $cert
          ParamConfig -Param "ProfessionalCertificationsLicensesDateReceived" -SqlType ([System.Data.SqlDbType]::VarChar) -Value $certDate
          ParamConfig -Param "ProfessionalCertificationsLicensesDetails" -SqlType ([System.Data.SqlDbType]::VarChar) -Value $certDetails
          ParamConfig -Param "CLEARDATE" -SqlType ([System.Data.SqlDbType]::Bit) -Value 0
          $a = ExecuteNonQuery

          $knt++
        }
        
      }
    }

    $file | Move-Item -Destination "\\001sql02\e$\integrations\cornerstone\Archive\ProfessionalCertificationsLicenses" -Force
  }
}

function SelfIdentifiedAreasOfExpertise() {
  Get-ChildItem -Path "SelfIdentifiedAreasOfExpertise" | % {
    $file = $_
    $nameParts = $file.Name -split "_"
    $dateParts = $nameParts[-5..-1]
    $date = $dateParts[0]
    $runDate = Get-Date ([datetime]::parseexact($date, "yyyyMMdd", [System.Globalization.CultureInfo]::InvariantCulture)).Date -f "yyyy-MM-dd"
    $records = Import-Csv -Path $("SelfIdentifiedAreasOfExpertise/" + $file.Name)
    
    dslSql -ConnString "Data Source=sql.aldridge-electric.com;Integrated Security=SSPI;Initial Catalog=Aspn" {
      CmdConfig -Type "sp"  -Sql "cornerstone.sp_SelfIdentifiedAreasofExpertise"
  
      ClearParams
      ParamConfig -Param "RunDate" -SqlType ([System.Data.SqlDbType]::Date) -Value $runDate
      ParamConfig -Param "UserID" -SqlType ([System.Data.SqlDbType]::VarChar) -Value ""
      ParamConfig -Param "UserFullName" -SqlType ([System.Data.SqlDbType]::VarChar) -Value ""
      ParamConfig -Param "Division" -SqlType ([System.Data.SqlDbType]::VarChar) -Value ""
      ParamConfig -Param "UserEmail" -SqlType ([System.Data.SqlDbType]::VarChar) -Value ""
      ParamConfig -Param "UserMobileNumber" -SqlType ([System.Data.SqlDbType]::VarChar) -Value ""
      ParamConfig -Param "SelfIdentifiedAreasofExpertiseSkill" -SqlType ([System.Data.SqlDbType]::VarChar) -Value ""
      ParamConfig -Param "SelfIdentifiedAreasofExpertiseProficiency" -SqlType ([System.Data.SqlDbType]::VarChar) -Value ""
      ParamConfig -Param "CLEARDATE" -SqlType ([System.Data.SqlDbType]::Bit) -Value 1
      $a = ExecuteNonQuery
  

      $records | % {
        $record = $_
        $knt = 0
        $sies = $record."Self Identified Areas of Expertise - Skill".Split(",")
        
        $sies | % {
          $sie = $_
          $prof = $($record."Self Identified Areas of Expertise - Proficiency".Split(","))
          
          if($null -eq $prof) { $prof = "" }
          if($prof.GetType().Name -eq "String") {
            $prof = $prof
          } else {
            $prof = $prof[$knt]
          }
          ClearParams
          ParamConfig -Param "RunDate" -SqlType ([System.Data.SqlDbType]::Date) -Value $runDate
          ParamConfig -Param "UserID" -SqlType ([System.Data.SqlDbType]::VarChar) -Value $record."User ID"
          ParamConfig -Param "UserFullName" -SqlType ([System.Data.SqlDbType]::VarChar) -Value $record."User Full Name"
          ParamConfig -Param "Division" -SqlType ([System.Data.SqlDbType]::VarChar) -Value $record."Division"
          ParamConfig -Param "UserEmail" -SqlType ([System.Data.SqlDbType]::VarChar) -Value $record."User Email"
          ParamConfig -Param "UserMobileNumber" -SqlType ([System.Data.SqlDbType]::VarChar) -Value $record."User Mobile Number"
          ParamConfig -Param "SelfIdentifiedAreasofExpertiseSkill" -SqlType ([System.Data.SqlDbType]::VarChar) -Value $sie
          ParamConfig -Param "SelfIdentifiedAreasofExpertiseProficiency" -SqlType ([System.Data.SqlDbType]::VarChar) -Value $prof
          ParamConfig -Param "CLEARDATE" -SqlType ([System.Data.SqlDbType]::Bit) -Value 0
          $a = ExecuteNonQuery

          $knt++
        }
        
      }
    }

    $file | Move-Item -Destination "\\001sql02\e$\integrations\cornerstone\Archive\SelfIdentifiedAreasOfExpertise" -Force
  }
}


function AllMainTopics() {
    
    $files = Get-ChildItem -Path "AllMainTopics" 
    foreach($file in $files) {

        $SQLServer = "sql.aldridge-electric.com"
        $db = "ASPN"

        # Clear Table Data
        $sql = "delete FROM [ASPN].[cornerstone].[AllMainTopics]"
        Invoke-Sqlcmd -ServerInstance $SQLServer -Database $db -Query $sql        
        
        # Get File Name and Date of Export 

        $nameParts = $file.Name -split "_"
        $dateParts = $nameParts[-5..-1]
        $date = $dateParts[0]
        $runDate = Get-Date ([datetime]::parseexact($date, "yyyyMMdd", [System.Globalization.CultureInfo]::InvariantCulture)).Date -f "yyyy-MM-dd"

        # Import FTP File Data 

        $records = Import-Csv -Path $("AllMainTopics\" + $file.Name)
        foreach($record in $records) {
            $UserID = $record.'User - User ID' 
            $UserFullName = $record.'User - User Full Name' -replace("'","")
            if([string]::ISNullOrWhiteSpace($record.'User - User Last Hire Date')) {
                $UserLastHireDate = $null
            } else {
                $UserLastHireDate = [datetime]::parseexact($record.'User - User Last Hire Date', 'MM/dd/yyyy', $null).tostring('yyyy-MM-dd')
            }
            $UserPosition = $record.'User - Position' -replace("'","")
            $UserDivision = $record.'User - Division' -replace("'","")
            $UserLocation = $record.'User - Location' -replace("'","")
            $UserManagerUserFullName = $record.'User - Manager - User Full Name' -replace("'","")
            $TrainingTrainingTitle = $record.'Training - Training Title' -replace("'","")
            $TranscriptTranscriptStatus = $record.'Transcript - Transcript Status' -replace("'","")
            If([string]::ISNullOrWhiteSpace($record.'Transcript - Transcript Completed Date')) {
                $TranscriptTranscriptCompletedDate = $null
            } else {
                $TranscriptTranscriptCompletedDate = [datetime]::parseexact(@($record.'Transcript - Transcript Completed Date').substring(0,10), 'MM/dd/yyyy', $null).tostring('yyyy-MM-dd')
            }
            $TrainingTrainingType = $record.'Training - Training Type' -replace("'","")

            $sql = "INSERT INTO [ASPN].[cornerstone].[AllMainTopics] (
                    UserID,
                    UserFullName,  
                    UserLastHireDate,
                    UserPosition, 
                    UserDivision, 
                    UserLocation, 
                    UserManagerUserFullName,
                    TrainingTrainingTitle,
                    TranscriptTranscriptStatus,
                    TranscriptTranscriptCompletedDate,
                    TrainingTrainingType,
                    RunDate
            ) VALUES (
                    '$UserID',
                    '$UserFullName',
                    '$UserLastHireDate',
                    '$UserPosition', 
                    '$UserDivision', 
                    '$UserLocation', 
                    '$UserManagerUserFullName',
                    '$TrainingTrainingTitle',
                    '$TranscriptTranscriptStatus',
                    '$TranscriptTranscriptCompletedDate',
                    '$TrainingTrainingType',
                    '$runDate'
                    );"

            Invoke-Sqlcmd -ServerInstance $SQLServer -Database $db -Query $sql

        }
            
        # Copy imported file to Archive
        $file | Move-Item -Destination "\\001sql02\e$\integrations\cornerstone\Archive\AllMainTopics" -Force
    }
}

function EquipmentEvaluationStatus() {
    
    $files = Get-ChildItem -Path "EquipmentEvaluationStatus" 
    foreach($file in $files) {

        $SQLServer = "sql.aldridge-electric.com"
        $db = "ASPN"
        
        # Clear Table Data
        $sql = "delete FROM [ASPN].[cornerstone].[EquipmentEvaluationStatus]"
        Invoke-Sqlcmd -ServerInstance $SQLServer -Database $db -Query $sql

        # Get File Name and Date of Export 

        $nameParts = $file.Name -split "_"
        $dateParts = $nameParts[-5..-1]
        $date = $dateParts[0]
        $runDate = Get-Date ([datetime]::parseexact($date, "yyyyMMdd", [System.Globalization.CultureInfo]::InvariantCulture)).Date -f "yyyy-MM-dd"
        
        # Import FTP File Data
        
        $records = Import-Csv -Path $("EquipmentEvaluationStatus\" + $file.Name)
        foreach($record in $records) {
            $UserID = $record.'User - User ID' 
            $UserFullName = $record.'User - User Full Name' -replace("'","")
            if([string]::ISNullOrWhiteSpace($record.'User - User Last Hire Date')) {
                $UserLastHireDate = $null
            } else {
                $UserLastHireDate = [datetime]::parseexact($record.'User - User Last Hire Date', 'MM/dd/yyyy', $null).tostring('yyyy-MM-dd')
            }
            $UserPosition = $record.'User - Position' -replace("'","")
            $UserDivision = $record.'User - Division' -replace("'","")
            $UserManagerFullName = $record.'User - Manager - User Full Name' -replace("'","")
            $AssignmentTrainingTitle = $record.'Assignment Training - Assignment Training Title' -replace("'","")
            If([string]::ISNullOrWhiteSpace($record.'On the Job Training & Express Class - Date Observed')) {
                $OntheJobTrainingExpressClassDateObserved = $null
            } Else {
                $OntheJobTrainingExpressClassDateObserved = $record.'On the Job Training & Express Class - Date Observed'
            }
            $OntheJobTrainingExpressClassObservedTrainingResult = $record.'On the Job Training & Express Class - Observed Training Result' -replace("'","")
            $OntheJobTrainingExpressClassObservedTrainingComments = $record.'On the Job Training & Express Class - Observed Training Comments' -replace("'","")
            $UserLocation = $record.'User - Location' -replace("'","")
                                    
            $sql = "INSERT INTO [ASPN].[cornerstone].[EquipmentEvaluationStatus] (
                    UserID,
                    UserFullName,  
                    UserLastHireDate,
                    UserPosition, 
                    UserDivision, 
                    UserManagerFullName,
                    AssignmentTrainingTitle,
                    OntheJobTrainingExpressClassDateObserved,
                    OntheJobTrainingExpressClassObservedTrainingResult,
                    OntheJobTrainingExpressClassObservedTrainingComments,
                    UserLocation,
                    RunDate
            ) VALUES (
                    '$UserID',
                    '$UserFullName',
                    '$UserLastHireDate',
                    '$UserPosition', 
                    '$UserDivision', 
                    '$UserManagerFullName',
                    '$AssignmentTrainingTitle',
                    '$OntheJobTrainingExpressClassDateObserved',
                    '$OntheJobTrainingExpressClassObservedTrainingResult',
                    '$OntheJobTrainingExpressClassObservedTrainingComments',
                    '$UserLocation',
                    '$runDate'
                    );"
            #$sql
            Invoke-Sqlcmd -ServerInstance $SQLServer -Database $db -Query $sql

        }

        # Copy imported file to Archive

        $file | Move-Item -Destination "\\001sql02\e$\integrations\cornerstone\Archive\EquipmentEvaluationStatus" -Force
    }
}

AldridgeQualificationPrograms
AldridgeSubjectMatterExperts
CompanyResources
ProfessionalCertificationsLicenses
SelfIdentifiedAreasOfExpertise
AllMainTopics
EquipmentEvaluationStatus