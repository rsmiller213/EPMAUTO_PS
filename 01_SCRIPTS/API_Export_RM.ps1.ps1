# ===================================================================================
#   Author : 
#   Created On : 
#   Purpose : Export Automation
# ===================================================================================


function EPM_Find-Member ($tree, $alias) {  
    $subtree = ""
    if ("$($tree.alias)" -eq "$alias"){
        return $tree.name
    } else {
        $tree.children.ForEach({
            if ($subtree -eq ""){ $subtree = EPM_Find-Member $_ $alias}
        })
        return $subtree
    }
}
# -----------------------------------------------------------------------------------
#   STARTING TASKS
# -----------------------------------------------------------------------------------

$API_PASSFILE = "$PSScriptRoot\API_pw.epw"
$API_USER = "vseeryada@cherryroad.com"
$API_PREFIX = "epm"
$API_DOMAIN = "cmsepm3"
$API_DC = "us6"
$OUTPUT_PATH = "C:\Users\VSeeryada\OneDrive - CherryRoad Technologies\Data_01212021\1.Projects\1.CMS\4.Data\5.BAAS\TEST EXPORTS"

	# Only need to run once to create the password file then can re-comment this out.
    # Create a separate Shell script with $API_PASSFILE = "$PSScriptRoot\API_pw.epw" line + the following 3 lines to encrypt the password 
#$tempEncName = [System.Text.Encoding]::UTF8.GetBytes("$API_DOMAIN" + '.' + $API_USER + ':' + "<EnterPassword>")
#$tempEncAuth = [System.Convert]::ToBase64String($tempEncName)
#Set-Content "$API_PASSFILE" -Value $tempEncAuth -Force

$API_HEADER = @{"Authorization"="Basic " + (Get-Content -Path "$API_PASSFILE")}
$API_BASE_URL = "https://$API_PREFIX-test-$API_DOMAIN.epm.$API_DC.oraclecloud.com/HyperionPlanning/rest/v3"

$EXPORT_URI = "$API_BASE_URL/applications/CMSPLAN/plantypes/BAAS/exportdataslice"
$DIM_URI = "$API_BASE_URL/internal/applications/CMSPLAN/plantypes/BAAS/dimensions"

# https://epm-test-cmsepm3.epm.us6.oraclecloud.com/HyperionPlanning/rest/v3/internal/applications/CMSPLAN/plantypes/BAAS/dimensions/Object

# uncomment to test Conn should get a response
#$TestResponse = Invoke-RestMethod -Uri $API_BASE_URL -Headers $API_HEADER -Method GET -ContentType "application/json"
#Write-Host $TestResponse



# -----------------------------------------------------------------------------------
#   PROCESSING TASKS
# -----------------------------------------------------------------------------------


# Build The Initial Export to Loop through and Create Additional Exports
$InitialPayload = @"
        {
    "exportPlanningData": false,
    "gridDefinition": {
        "suppressMissingBlocks": true,
        "suppressMissingRows": true,
        "pov": {
            "dimensions": ["Years","Scenario","Fund","Version","Period","Layer","Purpose","Location","Object","LocalField"],                
            "members":  [["FY21"] ,["Budget"],["FD30"],["BAAS Send to State"],["Jul"],["No_BAASLine"] ,["No_Purpose"],["No_Location"],["No_Object"],["No_LocalField"]]
            
        },
        "columns": [{
                "dimensions": ["BAASDetails"],
                "members": [ ["BAASAmdDetailFlag"] ] 
                    }],
        "rows": [{
                "dimensions": ["PRC","AmendDim"],
                 "members": [ ["ILvl0Descendants(TPRC_BAAS)"],[ "ILvl0Descendants(Total Amendments)"] ]
                }]
        }
}
"@

# Store Results
$InitialResponse = Invoke-RestMethod -Uri $EXPORT_URI -Body $InitialPayload -Headers $API_HEADER -Method POST -ContentType "application/json" -TimeoutSec 6000 -UseBasicParsing
#$InitialResponse.Rows | ForEach-Object {
#  write-host ($_.headers)
#}


# Loop Throguh Results and pull out the PRC and AmendDim from the rows ($_.headers) to use in the secondary export
$InitialResponse.Rows | ForEach-Object {
	# For each department, run an export
    $ARR = $_.headers.split(" ")
    $MBRPRC = $ARR[0]
    $MBRAMDDIM = $ARR[1]
    write-host "$MBRPRC-$MBRAMDDIM"
	$ExportPayload = @"
	{
    "exportPlanningData": false,
    "gridDefinition": {
        "suppressMissingBlocks": true,
        "suppressMissingRows": true,
        "pov": {
            "dimensions": ["Years","Scenario","Fund","Version","Period","Purpose","Location","Object","LocalField","PRC","AmendDim"],                
            "members":  [["FY21"] ,["Budget"],["FD30"],["BAAS Send to State"],["Jul"],["No_Purpose"],["No_Location"],["No_Object"],["No_LocalField"],["$MBRPRC"],["$MBRAMDDIM"]]
            
        },
        "columns": [{
                "dimensions": ["BAASDetails"],
                "members": [ ["LEA number","Budget or Amendment","Budget or Amendment Number","Purpose Code", "Object Code","Site","LocalUse Field","Status of Budgeted Line",
                        "Budget Amount","Amendment/Change Amount", "Revised Budget Amount","Changed Item Justification"] ] 
                    }],
        "rows": [{
                "dimensions": ["Layer"],
                 "members": [ ["ILvl0Descendants(BAAS_MainLineEntry)"] ]
                }]
				
				
        }
}
"@

	# Filename for the export
	$OutFile = "$OUTPUT_PATH\EXP_$MBRPRC_$MBRAMDDIM.txt"


	# Execute the REST API
	$ExportResponse = Invoke-RestMethod -Uri $EXPORT_URI -Body $ExportPayload -Headers $API_HEADER -Method POST -ContentType "application/json" -TimeoutSec 6000 -UseBasicParsing
    Write-host $ExportResponse.rows.Count
    #Write results to file
	if ($ExportResponse.rows.count -gt 0) {

		try {

			$stream = [System.IO.StreamWriter]::new($OutFile,$false)

			$DIMPURPOSE = Invoke-RestMethod -Uri "$DIM_URI/Purpose" -Headers $API_HEADER -Method GET -ContentType "application/json" -TimeoutSec 6000 -UseBasicParsing
            $DIMOBJECT = Invoke-RestMethod -Uri "$DIM_URI/Object" -Headers $API_HEADER -Method GET -ContentType "application/json" -TimeoutSec 6000 -UseBasicParsing
            $DIMLOCATION = Invoke-RestMethod -Uri "$DIM_URI/Location" -Headers $API_HEADER -Method GET -ContentType "application/json" -TimeoutSec 6000 -UseBasicParsing
			
			# Get POV for Each Line 
			$arrTemp = @()
			$ExportResponse.pov | ForEach-Object {
				if (!@("No_Location","No_Object","No_Purpose").contains($_)) {
					$arrTemp += $_
				}
			}
			$pov = $arrTemp -join ("|,|")

			# Get Column Headers 
			$arrTemp = @()
			$ExportResponse.columns | ForEach-Object {
				$arrTemp += $_
			}
			#Write Header Line
			$stream.WriteLine("Years|,|Scenario|,|Fund|,|Version|,|Period|,|LocalField|,|PRC|,|AmendDim|,|Layer|,|$($arrTemp -join ("|,|"))")

            $ExportResponse.Rows | ForEach-Object {
				# "|" is the delimiter
				# Split each line 
				$line = @()
				$_.data | ForEach-Object {
					$line += $_
				}

				$BeginLine = "$($line[0])|,|$($line[1])|,|$($line[2])"

				# Grab Member Names
				$Purpose = EPM_Find-Member $DIMPURPOSE "$($line[3])"
				$Object = EPM_Find-Member $DIMOBJECT "$($line[4])"
				$Site = EPM_Find-Member $DIMLOCATION "$($line[5])"
				
				# Write Line to File 
				$stream.WriteLine( ("$pov|,|$(($_.headers) -join('|,|'))|,|$BeginLine|,|$Purpose|,|$Object|,|$Site|,|$($line[6])|,|$($line[7])|,|$($line[8])|,|$($line[9])|,|$($line[10])|,|$($line[11])"))
			
			}
		} catch {
			Write-Host "ERROR : $MBRPRC_$MBRAMDDIM : $_"
		} finally {
			if($stream.BaseStream){
                $stream.Close()
            }
		}
	}
}



<#
$DIMOBJECT = Invoke-RestMethod -Uri "$DIM_URI/Object" -Headers $API_HEADER -Method GET -ContentType "application/json" -TimeoutSec 6000 -UseBasicParsing
$TEST = EPM_Find-Member $DIMOBJECT "Other Admin Assignments"
Write-host "Output $TEST"#>
