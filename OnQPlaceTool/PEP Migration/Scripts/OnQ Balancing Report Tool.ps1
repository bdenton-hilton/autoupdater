# Point to the module folder
$basePath = Split-Path -Path (split-path -Path $PSCommandPath)
$templateFilePath = "$basePath\Assets\OnQ Balancing Report Templates.xlsx"
$modulePath = "$basePath\Modules\ImportExcel\7.8.10\ImportExcel.psd1"
$savedirectory = "$basePath\OnQ Balancing Reports"

Import-Module $modulePath

###################################################

function encyptedPlaintextPasswordToCredentials {
    param (
        $username,
        $encryptedpassword
    ) 
    return New-Object PSCredential ($username, (ConvertTo-SecureString $encryptedpassword -ErrorAction SilentlyContinue))
}

function Ensure-PathExists {
    param (
        [string]$path
    )

    if (-Not (Test-Path -Path $path)) {
        $null = New-Item -ItemType Directory -Path $path
        Write-Host "`nDirectory '$path' created successfully.`n" -ForegroundColor Green
    }
    else {
        Write-Host "`nDirectory '$path' already exists.`n" -ForegroundColor Green
    }
}

function Get-NewDate {
    param (
        [string]$dateString
    )

    # Convert the string to a DateTime object
    $date = [datetime]::ParseExact($dateString, "yyyy-MM-dd", $null)

    # Add 1 year and 1 day to the date
    $newDate = $date.AddYears(1).AddDays(1)

    # Return the new date in YYYY-MM-DD format
    return $newDate.ToString("yyyy-MM-dd")
}

function Find-ExcelCellCoordinates {
    param (
        [object]$worksheet,
        [string]$searchString,
        [int]$maxRows,
        [int]$maxCols,
        [int]$startRow
    )

    if ($startRow -lt 1) { $startRow = 1 }

    # Loop through each row and column to find the matching cell
    for ($row = $startRow; $row -le $maxRows; $row++) {
        for ($col = 1; $col -le $maxCols; $col++) {
            $cellValue = $worksheet.Cells[$row, $col].Value
            if ($cellValue -ceq $searchString) {
                return $row, $col
            }
        }
    }

    # Return -1, -1 if no match is found
    return -1, -1
}

function setSqlVariablesFromInncode {
    param (
        $inncode,
        $goLiveDate  
    )

    $bussinessDate = $goLiveDate
    $endDate = $(Get-NewDate -dateString $goLiveDate)

    $viewTotals = @"
GO

SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO

-- Static values
DECLARE @cPropertyID CHAR(5) = N'$inncode'
DECLARE @sdtCurBusDate SMALLDATETIME = '$bussinessDate'
    
BEGIN    
 SET implicit_transactions OFF    
 SET TRANSACTION ISOLATION LEVEL READ UNCOMMITTED    

 --FOR ROOM STATUS

   SET nocount ON  
  DECLARE  
    @tCurBusDate SMALLDATETIME
  
  SELECT  
    @tCurBusDate = @sdtCurBusDate
  SELECT  
    @tCurBusDate  = dbo.udf_DateOnly_Convert(@tCurBusDate) 

  CREATE TABLE #statusresults -- base entity is room_id  
  (  
    room_id   CHAR(6) NOT NULL,  
    crs_room_type  CHAR(6) NULL,  
    room_status  CHAR(3) NULL,
    stay_id INT NULL,
    property_id CHAR(5) NOT NULL,
    cur_bus_date SMALLDATETIME NULL,
    arrival_date smalldatetime NULL,
    departure_date smalldatetime NULL
  )  

  --FOR ARRIVALS

  CREATE TABLE #ARRIVALRESULTS    
(     
 facility_id     varchar(10) NULL,    
 propertyid     char(5),    
 property_name     varchar(50) NULL,    
 cur_bus_date    smalldatetime,    
 CSDTreport_request_date  smalldatetime NULL,    
 report_request_date   smalldatetime,    
 report_end_date    smalldatetime NULL,    
 uses_meta_rooms    bit NULL,    
 criteria_stamp    varchar(1000) NULL,     
 gtd_rooms_count    int NULL,    
 nongtd_rooms_count   int NULL,    
 gtd_preassigned_room_count int NULL,    
 nongtd_preassigned_room_count int NULL,    
 gtd_number_of_adults  int NULL,    
 nongtd_number_of_adults  int NULL,    
 gtd_number_of_children  int NULL,    
 nongtd_number_of_children int NULL,    
 sortfied     varchar(15) NULL, --goutham #8778    
 property_id     char(5) NULL,    
 guest_id     int NULL,    
 guest_lastname    varchar(30) NULL,    
 guest_firstname    varchar(30) NULL,     
 guest_middle_initial  char(1) NULL,    
 sortname     varchar(61) NULL,    
 business_name    varchar(60) NULL,    
 guarantor_ind    bit NULL,    
 group_code     char(18) NULL, -- VLS increased to 18    
 confirmation_num   char(10) NULL,    
 arrival_date    smalldatetime NULL,    
 arrival_time    smalldatetime NULL, --  Sri Kaza , Changed the datatype from char to smalldatetime    
 departure_date    smalldatetime NULL,    
 arrived_ind     bit NULL,    
 number_of_nights   int NULL,    
 number_of_adults   smallint NULL,    
 number_of_children   smallint NULL,    
 guarantee_type    char(2) NULL,    
 guarantee_credit_card  varchar(22) NULL,    
 crs_room_type    char(6) NULL,    
 room_id      char(6) NULL,    
 rate_plan_id    char(8) NULL,    
 is_guaranteed_ind   bit NULL,    
 status      varchar(15) NULL,    
 stay_group_id    int NULL,    
 stay_status     char(1) NULL,    
 stay_id      int NULL,    
 hsk_status     varchar(8) NULL,    
 effective_rate    money NULL,    
 linked_room_id    char(6) NULL,    
 guest_arrival_date   smalldatetime NULL,    
 guest_status    char(1) NULL,    
 shared_ind     bit NULL,    
 hhtier      char(1) NULL,    
 award_redemption_type  char(1) NULL,    
 in_house_ind    bit NULL DEFAULT 0,    
 honors_num     varchar(15) NULL,    
 folio_id     int NULL,    
 receipt_id     char (1) NULL,    
 credit_limit    money NULL,    
 mop       char(2) NULL,    
 guest_balance    money NULL,    
 auth_balance    money NULL,    
 group_id     int NULL,    
 walkin_ind     bit NULL,    
 reservation_date   smalldatetime NULL,    
 stay_length     int NULL,    
 compound_id     int NULL,    
 building_id     int NULL,    
 floor_num     smallint NULL,    
 rate_type_code    char(1) NULL,    
 market_type     char(1) NULL,    
 crs_prop_code    char(5) NULL, --goutham #8778    
 local_market_desc   varchar(60) NULL,    
 clean_ind bit    NULL,    
 occupied_ind bit   NULL,    
 hsk_occupied_discrepancy_ind bit NULL,    
 hsk_status_cd    int NULL,    
 checkin_source_code   varchar(8) NULL,    
 comment      varchar(1000) NULL,    
 dtbegindate     smalldatetime NULL,    
 Guest_title     varchar(10) NULL,    
 --Added by Jitendra        
 is_dk CHAR(1) NULL,       
 is_dci CHAR(1) NULL,      
 est_arrival_time SMALLDATETIME NULL,    
 auth_response_status CHAR (2) , -- Soumya Sharma 18/01/2018-- adding this column to match the number of columns fetched from cp_RPT_ARRIVALL     
 addon varchar(255) -- Sagar Sharma 11/08/2022-- adding this column to match the number of columns fetched from cp_RPT_ARRIVALL     
)  

  IF @@error <> 0  
  BEGIN  
    RAISERROR(63428, 11, 1)
	RAISERROR('cp_RPT_ARRIVREM:Error creating #statusresults', 11, 1)  
    RETURN  
  END 

  --room status pre-table input

    INSERT #statusresults  
  (  
    room_id,  
    crs_room_type,  
    room_status,
    stay_id,
    property_id,
    cur_bus_date
  )  
  SELECT  
    PI.room_id,  
    rt.crs_room_type,  
    (  
      CASE WHEN PI.out_of_order_item_id IS NOT NULL  
        THEN 'O O'  
        ELSE  
          CASE r.occupied_ind  
            WHEN 0 THEN 'V'  
            WHEN 1 THEN 'O'  
          END  
          + ' ' +  
          CASE r.hsk_status_cd  
            WHEN 1 THEN 'D'  
            WHEN 2 THEN 'P'  
            WHEN 3 THEN 'C'  
            WHEN 4 THEN 'R'  
          END  
      END  
    ),
    r.occupied_by_stay_id,
    PI.property_id,
    @tCurBusDate
  FROM  
    PHYSICAL_INVENTORY PI  
    JOIN property p ON  
      p.property_id = PI.property_id  
    JOIN vw_ROOM_NOT_DELETED r ON  
      r.room_id = PI.room_id  
      AND r.property_id = PI.property_id  
    JOIN ROOM_ROOM_TYPE rrt ON  
      rrt.room_id = r.room_id  
      AND rrt.property_id = r.property_id  
    JOIN ROOM_TYPE rt ON  
      rt.room_type_id = rrt.room_type_id  
      AND rt.property_id = rrt.property_id  
  WHERE  
    PI.property_id = @cPropertyID  
    AND PI.bus_date = @tCurBusDate

	  IF @@error <> 0  
  BEGIN  
    RAISERROR(62930, 11, 1)  
    RETURN  
  END  
  
  -- Update the stay info
  UPDATE #statusresults
  SET 
    arrival_date = er.arrival_date,
    departure_date = er.departure_date
  FROM 
    vw_EFFECTIVE_RATE er
  WHERE 
    er.stay_id = #statusresults.stay_id
    AND er.property_id = #statusresults.property_id
    AND er.bus_date = CASE  
      WHEN #statusresults.cur_bus_date <= er.arrival_date OR er.arrival_date = er.departure_date  
        THEN er.arrival_date  
      WHEN #statusresults.cur_bus_date < er.departure_date  
        THEN #statusresults.cur_bus_date  
      ELSE DATEADD(dd, -1, er.departure_date)  
    END

  IF @@error <> 0  
  BEGIN  
    RAISERROR(62931, 11, 1)  
    RETURN  
  END
  
  -- Remove duplicate entries where room_id is identical
  DELETE FROM #statusresults
  WHERE room_id IN (
      SELECT room_id
      FROM (
          SELECT room_id,
                 ROW_NUMBER() OVER (PARTITION BY room_id ORDER BY room_id) AS rn
          FROM #statusresults
      ) AS duplicates
      WHERE rn > 1
  )
  
  IF @@error <> 0  
  BEGIN  
    RAISERROR(62441, 11, 1)  
    RETURN  
  END 
  
  --arrivals pre-table input
  
  INSERT INTO #ARRIVALRESULTS    
EXEC cp_RPT_ARRIVALL   
 @cPropertyID = @cPropertyID,    
 @cbegindate = @sdtCurBusDate,    
 @cenddate = @sdtCurBusDate,    
 @cArrivalType = 'REMAINING',    
 @cCompanyName = '',    
 @cBeginRatePlanCode = NULL,    
 @cEndRatePlanCode = NULL,    
 @cRoomRate = NULL,    
 @nRoomAssignInd = 2,    
 @bMultiRoomInd = 0,    
 @bWalkInInd = 0,    
 @bSameDay = 0,    
 @nCRSPropId = 0,    
 @bAuthResponse = 0,    
 @nSortBy = 1
    
IF @@error <> 0    
 BEGIN    
  RAISERROR('cp_RPT_ARRIVREM:Error inserting #ARRIVALRESULTS', 11, 1)    
  RETURN    
 END    
    
-- Remove duplicate entries where shared_ind = 1
;WITH CTE AS (
    SELECT *,
           ROW_NUMBER() OVER (PARTITION BY stay_id 
                              ORDER BY 
                                  CASE WHEN guarantor_ind = 1 THEN 0 ELSE 1 END, 
                                  folio_id DESC) AS rn
    FROM #ARRIVALRESULTS
    WHERE shared_ind = 1
)
DELETE FROM CTE
WHERE rn > 1

IF @@error <> 0    
 BEGIN    
  RAISERROR('cp_RPT_ARRIVREM:Error deleting duplicates from #ARRIVALRESULTS', 11, 1)    
  RETURN    
 END
 
   -- Create the final output table
  CREATE TABLE #FINAL_RESULTS
  (
    crs_room_type char(6),
    view_totals_status varchar(20),
    stay_count int
  )
  
  --Room status results placement
    -- Insert Out of Order count
  INSERT INTO #FINAL_RESULTS (crs_room_type, view_totals_status, stay_count)
  SELECT crs_room_type, 'Out of Order', COUNT(*)
  FROM #statusresults
  WHERE room_status = 'O O'
  GROUP BY crs_room_type
  
  -- Insert Vacant Dirty count
  INSERT INTO #FINAL_RESULTS (crs_room_type, view_totals_status, stay_count)
  SELECT crs_room_type, 'Vacant Dirty', COUNT(*)
  FROM #statusresults
  WHERE room_status IN ('V D', 'V P')
  GROUP BY crs_room_type
  
  -- Insert Vacant Clean count
  INSERT INTO #FINAL_RESULTS (crs_room_type, view_totals_status, stay_count)
  SELECT crs_room_type, 'Vacant Clean', COUNT(*)
  FROM #statusresults
  WHERE room_status = 'V C'
  GROUP BY crs_room_type
  
  -- Insert Vacant Ready count
  INSERT INTO #FINAL_RESULTS (crs_room_type, view_totals_status, stay_count)
  SELECT crs_room_type, 'Vacant Ready', COUNT(*)
  FROM #statusresults
  WHERE room_status = 'V R'
  GROUP BY crs_room_type
  
  -- Insert In House count
  INSERT INTO #FINAL_RESULTS (crs_room_type, view_totals_status, stay_count)
  SELECT crs_room_type, 'In House', COUNT(*)
  FROM #statusresults
  WHERE room_status IN ('O D', 'O P', 'O C', 'O R') AND departure_date <> @tCurBusDate
  GROUP BY crs_room_type
  
  -- Insert Due Out count
  INSERT INTO #FINAL_RESULTS (crs_room_type, view_totals_status, stay_count)
  SELECT crs_room_type, 'Due Out', COUNT(*)
  FROM #statusresults
  WHERE room_status IN ('O D', 'O P', 'O C', 'O R') AND departure_date = @tCurBusDate
  GROUP BY crs_room_type
  
  -- Insert Day Use count
  INSERT INTO #FINAL_RESULTS (crs_room_type, view_totals_status, stay_count)
  SELECT crs_room_type, 'Day Use', COUNT(DISTINCT room_id)
  FROM #statusresults
  WHERE arrival_date = departure_date
  GROUP BY crs_room_type
  
  -- Insert Actual Total count
  INSERT INTO #FINAL_RESULTS (crs_room_type, view_totals_status, stay_count)
  SELECT crs_room_type, 'Actual Total', COUNT(*)
  FROM #statusresults
  GROUP BY crs_room_type
  
  -- Ensure all combinations are present with 0 counts if necessary
  INSERT INTO #FINAL_RESULTS (crs_room_type, view_totals_status, stay_count)
  SELECT DISTINCT crs_room_type, 'Out of Order', 0
  FROM #statusresults
  WHERE crs_room_type NOT IN (SELECT crs_room_type FROM #FINAL_RESULTS WHERE view_totals_status = 'Out of Order')
  
  INSERT INTO #FINAL_RESULTS (crs_room_type, view_totals_status, stay_count)
  SELECT DISTINCT crs_room_type, 'Vacant Dirty', 0
  FROM #statusresults
  WHERE crs_room_type NOT IN (SELECT crs_room_type FROM #FINAL_RESULTS WHERE view_totals_status = 'Vacant Dirty')
  
  INSERT INTO #FINAL_RESULTS (crs_room_type, view_totals_status, stay_count)
  SELECT DISTINCT crs_room_type, 'Vacant Clean', 0
  FROM #statusresults
  WHERE crs_room_type NOT IN (SELECT crs_room_type FROM #FINAL_RESULTS WHERE view_totals_status = 'Vacant Clean')
  
  INSERT INTO #FINAL_RESULTS (crs_room_type, view_totals_status, stay_count)
  SELECT DISTINCT crs_room_type, 'Vacant Ready', 0
  FROM #statusresults
  WHERE crs_room_type NOT IN (SELECT crs_room_type FROM #FINAL_RESULTS WHERE view_totals_status = 'Vacant Ready')
  
  INSERT INTO #FINAL_RESULTS (crs_room_type, view_totals_status, stay_count)
  SELECT DISTINCT crs_room_type, 'In House', 0
  FROM #statusresults
  WHERE crs_room_type NOT IN (SELECT crs_room_type FROM #FINAL_RESULTS WHERE view_totals_status = 'In House')
  
  INSERT INTO #FINAL_RESULTS (crs_room_type, view_totals_status, stay_count)
  SELECT DISTINCT crs_room_type, 'Due Out', 0
  FROM #statusresults
  WHERE crs_room_type NOT IN (SELECT crs_room_type FROM #FINAL_RESULTS WHERE view_totals_status = 'Due Out')

    INSERT INTO #FINAL_RESULTS (crs_room_type, view_totals_status, stay_count)
  SELECT DISTINCT crs_room_type, 'Day Use', 0
  FROM #statusresults
  WHERE crs_room_type NOT IN (SELECT crs_room_type FROM #FINAL_RESULTS WHERE view_totals_status = 'Arrival = Departure')
  
  INSERT INTO #FINAL_RESULTS (crs_room_type, view_totals_status, stay_count)
  SELECT DISTINCT crs_room_type, 'Actual Total', 0
  FROM #statusresults
  WHERE crs_room_type NOT IN (SELECT crs_room_type FROM #FINAL_RESULTS WHERE view_totals_status = 'Actual Total') 

  --Add arrival results

  -- Insert GTD Arrivals count
INSERT INTO #FINAL_RESULTS (crs_room_type, view_totals_status, stay_count)
SELECT crs_room_type, 'GTD Arrivals', COUNT(*)
FROM #ARRIVALRESULTS
WHERE is_guaranteed_ind = 1
GROUP BY crs_room_type

-- Insert Non-GTD Arrivals count
INSERT INTO #FINAL_RESULTS (crs_room_type, view_totals_status, stay_count)
SELECT crs_room_type, 'Non-GTD Arrivals', COUNT(*)
FROM #ARRIVALRESULTS
WHERE is_guaranteed_ind = 0
GROUP BY crs_room_type

-- Insert Pending count
INSERT INTO #FINAL_RESULTS (crs_room_type, view_totals_status, stay_count)
SELECT crs_room_type, 'Pending', COUNT(*)
FROM #ARRIVALRESULTS
WHERE status <> 'NA'
GROUP BY crs_room_type

-- Insert Pre-Assigned count
INSERT INTO #FINAL_RESULTS (crs_room_type, view_totals_status, stay_count)
SELECT crs_room_type, 'Pre-Assigned', COUNT(*)
FROM #ARRIVALRESULTS
WHERE room_id IS NOT NULL
GROUP BY crs_room_type

-- Ensure all combinations are present with 0 counts if necessary
INSERT INTO #FINAL_RESULTS (crs_room_type, view_totals_status, stay_count)
SELECT DISTINCT crs_room_type, 'GTD Arrivals', 0
FROM #ARRIVALRESULTS
WHERE crs_room_type NOT IN (SELECT crs_room_type FROM #FINAL_RESULTS WHERE view_totals_status = 'GTD Arrivals')

INSERT INTO #FINAL_RESULTS (crs_room_type, view_totals_status, stay_count)
SELECT DISTINCT crs_room_type, 'Non-GTD Arrivals', 0
FROM #ARRIVALRESULTS
WHERE crs_room_type NOT IN (SELECT crs_room_type FROM #FINAL_RESULTS WHERE view_totals_status = 'Non-GTD Arrivals')

INSERT INTO #FINAL_RESULTS (crs_room_type, view_totals_status, stay_count)
SELECT DISTINCT crs_room_type, 'Pending', 0
FROM #ARRIVALRESULTS
WHERE crs_room_type NOT IN (SELECT crs_room_type FROM #FINAL_RESULTS WHERE view_totals_status = 'Pending')

INSERT INTO #FINAL_RESULTS (crs_room_type, view_totals_status, stay_count)
SELECT DISTINCT crs_room_type, 'Pre-Assigned', 0
FROM #ARRIVALRESULTS
WHERE crs_room_type NOT IN (SELECT crs_room_type FROM #FINAL_RESULTS WHERE view_totals_status = 'Pre-Assigned')

-- Create a temporary table to hold all unique crs_room_type values
CREATE TABLE #ALL_ROOM_TYPES (crs_room_type CHAR(6))

-- Insert unique crs_room_type values from #statusresults
INSERT INTO #ALL_ROOM_TYPES (crs_room_type)
SELECT DISTINCT crs_room_type FROM #statusresults

-- Insert unique crs_room_type values from #ARRIVALRESULTS
INSERT INTO #ALL_ROOM_TYPES (crs_room_type)
SELECT DISTINCT crs_room_type FROM #ARRIVALRESULTS
WHERE crs_room_type NOT IN (SELECT crs_room_type FROM #ALL_ROOM_TYPES)

-- Ensure all crs_room_type values are present in #FINAL_RESULTS with 0 counts if necessary
INSERT INTO #FINAL_RESULTS (crs_room_type, view_totals_status, stay_count)
SELECT DISTINCT art.crs_room_type, 'GTD Arrivals', 0
FROM #ALL_ROOM_TYPES art
WHERE art.crs_room_type NOT IN (SELECT crs_room_type FROM #FINAL_RESULTS WHERE view_totals_status = 'GTD Arrivals')

INSERT INTO #FINAL_RESULTS (crs_room_type, view_totals_status, stay_count)
SELECT DISTINCT art.crs_room_type, 'Non-GTD Arrivals', 0
FROM #ALL_ROOM_TYPES art
WHERE art.crs_room_type NOT IN (SELECT crs_room_type FROM #FINAL_RESULTS WHERE view_totals_status = 'Non-GTD Arrivals')

INSERT INTO #FINAL_RESULTS (crs_room_type, view_totals_status, stay_count)
SELECT DISTINCT art.crs_room_type, 'Pending', 0
FROM #ALL_ROOM_TYPES art
WHERE art.crs_room_type NOT IN (SELECT crs_room_type FROM #FINAL_RESULTS WHERE view_totals_status = 'Pending')

INSERT INTO #FINAL_RESULTS (crs_room_type, view_totals_status, stay_count)
SELECT DISTINCT art.crs_room_type, 'Pre-Assigned', 0
FROM #ALL_ROOM_TYPES art
WHERE art.crs_room_type NOT IN (SELECT crs_room_type FROM #FINAL_RESULTS WHERE view_totals_status = 'Pre-Assigned')

-- Drop the temporary table
DROP TABLE #ALL_ROOM_TYPES

IF @@error <> 0  
  BEGIN  
    RAISERROR('Error inserting into #FINAL_RESULTS', 11, 1)  
    RETURN  
  END 

    -- Select the final results sorted by crs_room_type
  SELECT crs_room_type,
         view_totals_status,
         stay_count
  FROM #FINAL_RESULTS
  ORDER BY crs_room_type, view_totals_status
  
  IF @@error <> 0  
  BEGIN  
    RAISERROR('Error selecting #FINAL_RESULTS', 11, 1)  
    RETURN  
  END  
  
END
"@

    $ADDONFULL = @"
GO

DECLARE @return_value int

EXEC @return_value = [dbo].[cp_RPT_ADDON_FULFILMENT]
    @cPropertyID = N'$inncode',
    @dtStartDate = '$bussinessDate',
    @dtEndDate = '$endDate',
    @vAddon_code = NULL,
    @vArrivalTierStamp = NULL,
    @vHHonorsTierStamp = NULL,
    @nSortParam = 0

SELECT 'Return Value' = @return_value

GO
"@

    $RMDETMNT = @"
GO

DECLARE @return_value int

EXEC @return_value = [dbo].[cp_RPT_RMDETMNT]
    @cPropertyID = N'$inncode',
    @sdtArrival = '2000-01-01',
    @sdtDeparture = '$bussinessDate',
    @vStamp = NULL,
    @vWhereClause = NULL

SELECT 'Return Value' = @return_value

"@

    $ADVDPARR = @"
GO

DECLARE	@return_value int

EXEC	@return_value = [dbo].[cp_RPT_ADVDPARR]
		@cPropertyID = N'$inncode',
		@cBeginDate = '$bussinessDate',
		@cEndDate = '$endDate',
		@cIncludeRefund = N'B',
		@cIncludeAllowAdj = N'B',
		@nSort = 0,
		@vStamp1 = NULL,
		@vWhereClause1 = N'WHERE 1=1'

SELECT	'Return Value' = @return_value

GO
"@

    $HK_ALLSUM = @"
GO

DECLARE	@return_value int

EXEC	@return_value = [dbo].[cp_HK_ALL_SUM_RPTS]
		@pcPropertyID = N'$inncode',
		@vStamp = NULL,
		@vWhereClause = NULL

SELECT	'Return Value' = @return_value

GO
"@

    $ARAGENUM = @"
GO

DECLARE	@return_value int

EXEC	@return_value = [dbo].[cp_RPT_AR_AGING]
		@cPropertyID = N'$inncode',
		@sdtCurDate = '$bussinessDate',
		@nARAccountID = NULL,
		@nHotelTotalsOnly = NULL,
		@cCalculateAgeBy = NULL

SELECT	'Return Value' = @return_value

GO
"@

    $GPBALSUM = @"
GO

DECLARE	@return_value int

EXEC	@return_value = [dbo].[cp_RPT_GPBALSUM]
		@cPropertyid = N'$inncode',
		@sdtCurDate = '$bussinessDate',
		@cGroupcode = NULL,
		@cBalFilter = NULL

SELECT	'Return Value' = @return_value

GO
"@

    $RMBALSUM = @"
GO

DECLARE	@return_value int

EXEC	@return_value = [dbo].[cp_RPT_RMBALSUM]
		@cPropertyId = N'$inncode',
		@dtBusDate = '$bussinessDate',
		@nIncludeZeroBalance = 1,
		@nCompoundId = NULL,
		@nBuildingId = NULL,
		@nFloor_Num = NULL,
		@nSortBy = 0,
		@nCRSPropId = NULL

SELECT	'Return Value' = @return_value

GO
"@

    $OOBALSUM = @"
GO

DECLARE	@return_value int

EXEC	@return_value = [dbo].[cp_FOLIO_EXPIRED_WITH_BALANCE]
		@cPropertyID = N'$inncode',
		@dtReportDate = '$bussinessDate'

SELECT	'Return Value' = @return_value

GO
"@

    $DPAUDDAA = @"

GO

DECLARE	@return_value int

EXEC	@return_value = [dbo].[cp_RPT_DPAUDDAA]
		@cPropertyID = N'$inncode',
		@sdtStartDate = '$bussinessDate',
		@sdtEndDate = '$bussinessDate',
		@nSort = NULL,
		@cPrintTaxDetails = N'N',
		@cPrintAllowancesOnly = N'N',
		@vStamp1 = NULL,
		@vStamp2 = NULL,
		@vStamp3 = NULL,
		@vStamp4 = NULL,
		@vStamp5 = NULL,
		@vStamp6 = NULL,
		@vStamp7 = NULL,
		@vWhereClause1 = NULL,
		@vWhereClause2 = NULL,
		@vWhereClause3 = NULL,
		@vWhereClause4 = NULL,
		@vWhereClause5 = NULL,
		@vWhereClause6 = NULL,
		@vWhereClause7 = NULL

SELECT	'Return Value' = @return_value

GO

"@

    $RECRCP_GL = @"
GO

DECLARE @return_value int

-- Drop the temporary table if it already exists
IF OBJECT_ID('tempdb..#TempResults') IS NOT NULL
BEGIN
    DROP TABLE #TempResults
END

-- Create a temporary table to store the results
CREATE TABLE #TempResults (
    property_id0 VARCHAR(255),
    accounting_id VARCHAR(255),
    gl_account_id VARCHAR(255),
    accounting_id_desc VARCHAR(255),
    Yesterday_balance DECIMAL(18, 2),
    Today_debits DECIMAL(18, 2),
    Today_Credits DECIMAL(18, 2),
    acct_type VARCHAR(50),
    entry_type VARCHAR(50),
    entry_desc VARCHAR(255),
    facility_id VARCHAR(255),
    property_id char(5),
    property_name VARCHAR(255),
    cur_date DATE,
    time TIME,
    cur_bus_date DATE,
    YesterdayBal VARCHAR(255),
    TodayDebit VARCHAR(255),
    TodayCredit VARCHAR(255),
    Balance VARCHAR(255),
    NetOutstanding VARCHAR(255),
    Todaysoutstanding VARCHAR(255),
    Title VARCHAR(255)
)

-- Insert the results into the temporary table
INSERT INTO #TempResults
EXEC [dbo].[cp_RPT_RECRCP_GL]
    @cPropertyID = N'$inncode',
    @dtBusDate = '$bussinessDate'

-- Select all columns except the first one
SELECT 
    gl_account_id,
    accounting_id_desc,
    Yesterday_balance,
    Today_debits,
    Today_Credits,
    acct_type,
    entry_type,
    entry_desc,
    facility_id,
    property_id,
    property_name,
    cur_date,
    time,
    cur_bus_date,
    YesterdayBal,
    TodayDebit,
    TodayCredit,
    Balance,
    NetOutstanding,
    Todaysoutstanding,
    Title
FROM #TempResults

-- Drop the temporary table
DROP TABLE #TempResults

SELECT 'Return Value' = @return_value

GO
"@

    $HABALSUM = @"
GO

USE [hpms3]
GO

DECLARE @cPropertyID CHAR(5) = '$inncode'  -- Replace with the desired property ID
DECLARE @sdtCurDate SMALLDATETIME = '$bussinessDate'  -- Replace with the desired current date

SET implicit_transactions OFF
SET TRANSACTION isolation level READ uncommitted

DECLARE @dtCurDate SMALLDATETIME
SELECT @dtCurDate = @sdtCurDate
SELECT @dtCurDate = dbo.udf_DateOnly_Convert(@dtCurDate)
SELECT @sdtCurDate = @dtCurDate

--PROPERTY INFO ##########
CREATE TABLE #property_info
(
    facility_id VARCHAR(10) NULL,
    property_id CHAR(5) NOT NULL,
    property_name VARCHAR(50) NULL,
    cur_bus_date SMALLDATETIME NULL
)

CREATE TABLE #guest_info
(
    property_id CHAR(5) NOT NULL,
    house_acct_id INT NOT NULL,
    account_name VARCHAR(60) NOT NULL,
    last_posted_date DATETIME NULL,
    active_ind BIT NOT NULL,
    folio_id INT NOT NULL,
    folio_status CHAR(1) NULL,
    balance MONEY NULL
)

INSERT INTO #property_info
SELECT
    CONVERT(VARCHAR(10), p.facility_id) AS facility_id,
    p.property_id,
    p.property_name,
    @dtCurDate AS cur_bus_date
FROM
    property p
WHERE
    p.property_id = @cPropertyID

IF @@ERROR <> 0
BEGIN
    RAISERROR(60141, 11, 1)
    RETURN
END

--GUEST INFO ############
INSERT INTO #guest_info
SELECT
    ha.property_id,
    ha.house_acct_id,
    ha.account_name,
    NULL AS last_posted_date,
    ha.active_ind,
    f.folio_id,
    f.folio_status,
    NULL AS balance
FROM
    folio f
    JOIN property p ON p.property_id = f.property_id
    JOIN PROPERTY_EXTENSION pe ON pe.property_id = p.property_id
    JOIN house_account ha ON ha.house_acct_id = f.house_acct_id AND ha.property_id = f.property_id
WHERE
    f.property_id = @cPropertyID
    AND f.house_acct_id IS NOT NULL
    AND (f.folio_status IS NULL OR f.folio_status <> 'C')
    AND (ha.visible_ind = 1 OR ha.house_acct_id = pe.vat_house_acct_id)
    AND (ha.type IS NULL OR ha.type <> 'AD')

IF @@ERROR <> 0
BEGIN
    RAISERROR(62155, 11, 1)
    RETURN
END

UPDATE #guest_info
SET
    last_posted_date = dt.max_posted_datetime,
    balance = dt.bal_sum
FROM
    #guest_info gi
    JOIN (
        SELECT
            gi.folio_id,
            MAX(posted_datetime) AS max_posted_datetime,
            ISNULL(SUM(entry_amount), 0) AS bal_sum
        FROM
            vw_LEDGER vl
            JOIN #guest_info gi ON gi.folio_id = vl.folio_id
        WHERE
            vl.bus_date <= @dtCurDate
        GROUP BY
            gi.folio_id
    ) dt ON gi.folio_id = dt.folio_id

DECLARE @AdvBalSumResults TABLE
(
    property_id CHAR(5) NULL,
    property_name VARCHAR(60) NULL,
    facility_id VARCHAR(10) NULL,
    cur_bus_date SMALLDATETIME NULL,
    selection VARCHAR(255) NULL,
    sort_param VARCHAR(30) NULL,
    last_posted_date SMALLDATETIME NULL,
    arrival_date SMALLDATETIME NULL,
    conf_num VARCHAR(18) NULL,
    guest_name VARCHAR(60) NULL,
    valVAT MONEY NULL,
    valCharge MONEY NULL,
    valTotal MONEY NULL,
    guest_num INT NULL,
    guest_type CHAR(1) NULL,
    data_id INT NULL,
    sort_column VARCHAR(60) NULL,
    todays_date SMALLDATETIME NULL
)

INSERT INTO @AdvBalSumResults
EXEC cp_RPT_ADVDPBALSUM @cPropertyId, NULL, '', NULL

DECLARE @AdvanceDepositTotal MONEY
SELECT
    @AdvanceDepositTotal = CASE WHEN SUM(t3.valTotal) IS NULL THEN 0.00 ELSE SUM(t3.valTotal) END
FROM
    @AdvBalSumResults t3
WHERE
    t3.guest_type = 'H'

SELECT
    pi.facility_id,
    pi.property_id,
    pi.property_name,
    pi.cur_bus_date,
    gi.house_acct_id,
    gi.account_name,
    gi.last_posted_date,
    gi.active_ind,
    gi.folio_id,
    gi.folio_status,
    gi.balance,
    @AdvanceDepositTotal AS AdvanceDepositTotal
FROM
    #guest_info gi
    RIGHT OUTER JOIN #property_info pi ON gi.property_id = pi.property_id
ORDER BY
    gi.account_name

-- Drop temporary tables
DROP TABLE #guest_info
DROP TABLE #property_info

IF @@ERROR <> 0 RAISERROR(62157, 11, 1)
GO
"@


    $reports = [PSCustomObject]@{
        viewTotals = $viewTotals
        ADDONFULL  = $ADDONFULL
        RMDETMNT   = $RMDETMNT
        ADVDPARR   = $ADVDPARR
        HK_ALLSUM  = $HK_ALLSUM
        ARAGENUM   = $ARAGENUM
        GPBALSUM   = $GPBALSUM
        RMBALSUM   = $RMBALSUM
        OOBALSUM   = $OOBALSUM
        DPAUDDAA   = $DPAUDDAA
        RECRCP_GL  = $RECRCP_GL
        HABALSUM   = $HABALSUM
    }

    return $reports
}

function Get-ValidDate {
    while ($true) {
        # Prompt the user to enter a date
        $inputDate = $(Write-Host -NoNewLine) + $(Write-Host "`nPlease enter the Go-Live date (the date after Migration) in YYYY-MM-DD format: " -ForegroundColor Cyan -NoNewLine; Read-Host)

        # Replace different delimiters with hyphens
        $inputDate = $inputDate -replace '[./]', '-'

        # Split the input date into components
        $dateParts = $inputDate -split '-'

        # Check if the input has three parts (year, month, day)
        if ($dateParts.Length -ne 3) {
            Write-Host "Invalid format. Please enter the date in YYYY-MM-DD format." -ForegroundColor Red
            continue
        }

        # Extract year, month, and day
        $year = $dateParts[0]
        $month = $dateParts[1]
        $day = $dateParts[2]

        # Correct two-digit year
        if ($year.Length -eq 2) {
            $year = "20$year"
        }

        # Convert to integers
        $year = [int]$year
        $month = [int]$month
        $day = [int]$day

        # Validate month and day
        if ($month -lt 1 -or $month -gt 12) {
            Write-Host "Invalid month. Please enter a valid month (01-12)." -ForegroundColor Red
            continue
        }

        if ($day -lt 1 -or $day -gt 31) {
            Write-Host "Invalid day. Please enter a valid day (01-31)." -ForegroundColor Red
            continue
        }

        # Create a DateTime object
        try {
            $date = [datetime]::new($year, $month, $day)
        }
        catch {
            Write-Host "Invalid date. Please enter a valid date." -ForegroundColor Red
            continue
        }

        # Return the valid date as a string in YYYY-MM-DD format
        return $date.ToString("yyyy-MM-dd")
    }
}


####################################################

function SafeParseDouble {
    param (
        [string]$value
    )
    if ([string]::IsNullOrWhiteSpace($value)) {
        return 0.0
    }
    $result = 0.0
    [double]::TryParse($value, [ref]$result) | Out-Null
    return $result
}

function SafeParseNumberandRound {
    param (
        [string]$value
    )
    if ([string]::IsNullOrWhiteSpace($value)) {
        return $value
    }
    $result = 0.0
    if ([double]::TryParse($value, [ref]$result)) {
        return [math]::Round($result, 2)
    }
    else {
        return $value
    }
}

function Get-FormattedTime {
    param (
        [string]$dateTimeString
    )

    # Parse the input string to a DateTime object
    $dateTime = [datetime]::Parse($dateTimeString)

    # Format the DateTime object to the desired format
    $formattedTime = $dateTime.ToString("h:mm:ss tt")

    return $formattedTime
}

function sumList {
    param (
        $list
    )
    $total_sum = 0.0
    foreach ($value in $list) {
        $parsed_value = SafeParseNumberandRound $value
        if ($parsed_value -is [double]) {
            $total_sum += $parsed_value
        }
    }
    return $total_sum
}

function ConvertToTitleCase {
    param (
        [string]$inputString
    )
    if ([string]::IsNullOrWhiteSpace($inputString)) {
        return $inputString
    }
    $cultureInfo = [System.Globalization.CultureInfo]::CurrentCulture
    $textInfo = $cultureInfo.TextInfo
    return $textInfo.ToTitleCase($inputString.ToLower())
}

function Split-TimeFromDateTime {
    param (
        [string]$dateTimeString
    )
  
    # Split the string by spaces and return the time part
    $parts = $dateTimeString -split " "
    return $parts[1] + " " + $parts[2]
}

function SafeGetDateString {
    param (
        [string]$date
    )
    try {
        return (Get-Date $date).ToString("M/d/yyyy")
    }
    catch {
        return ""
    }
}

######################################################

function ADVDPARR.RPT {
    param (
        [psobject]$csvData,
        [string]$templatePath,
        [string]$outputDir
    )

    # Load the Excel file
    $excel = Open-ExcelPackage -Path $templatePath

    # Select the "Tester" worksheet
    $worksheet = $excel.Workbook.Worksheets["ADVDPARR.RPT"]
  
    $bussinessDate = SafeGetDateString -date $csvData.cur_bus_date[0]
  
    $advancedDate = (Get-Date $($csvData.cur_bus_date[0])).AddYears(1).AddDays(1).ToString("M/d/yyyy")

    $worksheet.Cells[1, 4].Value = Get-Date -Format "M/d/yyyy"
    $worksheet.Cells[2, 4].Value = $bussinessDate
    $worksheet.Cells[3, 4].Value = Get-Date -Format "h:mm tt"

    $worksheet.Cells[1, 6].Value = $csvData.property_id[0] + " - " + $csvData.property_name[0] + " - " + $csvData.facility_id[0]
    $worksheet.Cells[3, 6].Value = "Date Range : " + $bussinessDate + " TO " + $advancedDate


    $offset = 0
    $runningDeposit = 0
    $runningPosted = 0
    $runningRefund = 0
    $runningAdjAllow = 0
    $runningBalance = 0

    # Loop through the data and insert each entry into the specified row
    for ($i = 0; $i -lt $csvData.Count; $i++) {
        $row = $i * 2 + (19 + $offset)
        $col = 1

        # Assign each expression to a single variable with error handling
        $guestName = "$($csvData[$i].guest_lastname_or_acct_name)$(if ($($csvData[$i].guest_firstname).Length -lt 1) {} else { "/$($csvData[$i].guest_firstname)" })"
        $arrivalDate = SafeGetDateString $csvData[$i].arrival_date
        $dueDate = SafeGetDateString $csvData[$i].due_date
        $depositType = ($($csvData[$i].deposit_type_desc) -split ' ')[0]
        $paymentDesc = $csvData[$i].payment_desc
        $confirmationNum = $csvData[$i].confirmation_num
        $dueAmount = [Math]::Round($csvData[$i].due_amount, 2)
        $postedAmount = [Math]::Abs([Math]::Round((SafeParseDouble $csvData[$i].payments), 2))
        $adjAllowanceAmt = [Math]::Abs([Math]::Round((SafeParseDouble $csvData[$i].adj_amt) + (SafeParseDouble $csvData[$i].allowance_amt), 2))
        $refundAmt = [Math]::Round((SafeParseDouble $csvData[$i].refund_amt), 2)
        $balanceMnt = [Math]::Round($csvData[$i].posted_amount, 2)
  
        $runningDeposit += $dueAmount
        $runningPosted += $postedAmount
        $runningRefund += $refundAmt
        $runningAdjAllow += $adjAllowanceAmt
        $runningBalance += $balanceMnt

        # Define the expressions to be entered into the worksheet
        $expressions = @(
            $guestName,
            $arrivalDate,
            $dueDate,
            $depositType,
            $paymentDesc,
            $confirmationNum,
            $dueAmount,
            $postedAmount,
            $refundAmt,
            $adjAllowanceAmt,
            $balanceMnt
        )

        if ($depositType.Length -le 0) {
            $offset += -2
            $finalRow = $row + $offset
        }
        else {
            # Insert the expressions into the worksheet
            foreach ($expression in $expressions) {
                $worksheet.Cells[$row, $col].Value = $expression
                $col++
                $finalRow = $row
            }
        }
    }

    $finalRow += 2
    $col = 1

    $expressions = @(
        "TOTAL",
        "",
        "",
        "",
        "",
        "",
        $runningDeposit,
        $runningPosted,
        $runningRefund,
        $runningAdjAllow,
        $runningBalance
    )
    foreach ($expression in $expressions) {
        $worksheet.Cells[$finalRow, $col].Value = $expression
        $col++
    }

    $worksheet.Cells[($finalRow + 3), 1].Value = "END OF REPORT"

    $filename = "OnQ - Adv Deposit Summary By Arrival Date"

    # Define the new file name
    $filename = "$filename - $($csvData.property_id[0]) - $($(Get-date $csvData.cur_bus_date[0]).ToString("M-d-yyyy")).xlsx"

    $fullOutputPath = Join-Path -Path $outputDir -ChildPath $filename

    # Save the entire Excel file with the new name
    Close-ExcelPackage $excel -SaveAs $fullOutputPath
  
    # Reopen the saved Excel file
    $newExcel = Open-ExcelPackage -Path $fullOutputPath
  
    # Get the list of all worksheet names
    $worksheetNames = $newExcel.Workbook.Worksheets.Name
  
    foreach ($sheetName in $worksheetNames) {
        if ($sheetName -ne "ADVDPARR.RPT") {
            $newExcel.Workbook.Worksheets.Delete($sheetName)
        }
    }

    Close-ExcelPackage $newExcel -SaveAs $fullOutputPath

    Write-Host "Data has been successfully inserted into the Template worksheet and saved as '$filename'."
} #done

function RMDETMNT.RPT {
    param (
        [psobject]$csvData,
        [string]$templatePath,
        [string]$outputDir
    )

    # Load the Excel file
    $excel = Open-ExcelPackage -Path $templatePath

    # Select the "Tester" worksheet
    $worksheet = $excel.Workbook.Worksheets["RMDETMNT.RPT"]
  
    $bussinessDate = SafeGetDateString -date $csvData.cur_bus_date[0]

    $propertyID = $csvData.property_id[0]
    $saveDate = $(Get-Date $csvData.cur_bus_date[0]).ToString("M-d-yyyy")

    $worksheet.Cells[1, 4].Value = Get-Date -Format "M/d/yyyy"
    $worksheet.Cells[2, 4].Value = $bussinessDate
    $worksheet.Cells[3, 4].Value = Get-Date -Format "h:mm tt"

    $worksheet.Cells[4, 5].Value = $csvData.property_id[0]
    $worksheet.Cells[4, 8].Value = $csvData.property_id[0]

    $script:propertyIDHeader = $csvData.property_id[0] + " - " + $csvData.property_name[0] + " - " + $csvData.facility_id[0]

    $worksheet.Cells[1, 6].Value = $csvData.property_id[0] + " - " + $csvData.property_name[0] + " - " + $csvData.facility_id[0]
    $worksheet.Cells[3, 6].Value = "FOR " + $bussinessDate 

    # Remove entries where room_id is blank
    $csvData = $csvData | Where-Object { $_.room_id -ne "" } | Where-Object { $null -ne $_.room_id } | Sort-Object -Property { $_.sort_sequence_num }
 
    # Group by room_id and process each group

    $row = 12
    $col = 1

    if ($csvData.count -gt 0) {
        $groupedData = $csvData | Group-Object -Property room_id | ForEach-Object {
            $room_id = $_.Name
            $crs_room_type = ($_.Group | Select-Object -First 1).crs_room_type
            $reasonComments = $_.Group | ForEach-Object {
                "$($_.reason.Trim()); $($_.comment.Trim())"
            } | Sort-Object -Unique
            $earliestFromDate = (Get-Date ($_.Group | Sort-Object -Property from_date | Select-Object -First 1).from_date).ToString("M/d/yyyy")
            $latestToDate = (Get-Date ($_.Group | Sort-Object -Property to_date -Descending | Select-Object -First 1).to_date).ToString("M/d/yyyy")
            $latestToTime = try { (Get-Date ($_.Group | Sort-Object -Property return_time -Descending | Select-Object -First 1).return_time).ToString("h:mm:ss tt") }catch { "" }
            $return_status = ($_.Group | Select-Object -First 1).'return status'
            $oper = ($_.Group | Select-Object -First 1).employee_id
  
            [PSCustomObject]@{
                room_id        = $room_id
                crs_room_type  = $crs_room_type
                blank          = ""
                info           = "___________________________"
                reason_comment = ($reasonComments -join "; ")
                from_date      = $earliestFromDate
                to_date        = $latestToDate
                to_time        = $latestToTime
                return_status  = $return_status
                oper           = $oper
            }
        }

        # Iterate through each PSCustomObject and write each value one by one


        foreach ($item in $groupedData) {
            $col = 1
            foreach ($value in  $item.PSObject.Properties.Value) {
                $worksheet.Cells[$row, $col].Value = $value
                $col++ 
            }
            $row++
        }
    }
    else {
        Write-Host "No Out of Order Entries." -ForegroundColor Blue
        $row++
    }

    $finalRow = $row
    $worksheet.Cells[($finalRow + 3), 3].Value = 'TOTAL MAINTENANCE:'
    $worksheet.Cells[($finalRow + 3), 4].Value = $groupedData.Count

    $worksheet.Cells[($finalRow + 7), 1].Value = "END OF REPORT"

    $filename = "OnQ - Out Of Order"

    # Define the new file name
    $filename = "$filename - $propertyID - $saveDate.xlsx"

    $fullOutputPath = Join-Path -Path $outputDir -ChildPath $filename

    # Save the entire Excel file with the new name
    Close-ExcelPackage $excel -SaveAs $fullOutputPath
  
    # Reopen the saved Excel file
    $newExcel = Open-ExcelPackage -Path $fullOutputPath
  
    # Get the list of all worksheet names
    $worksheetNames = $newExcel.Workbook.Worksheets.Name
  
    # Loop through all worksheets and delete those that are not "HK_ALLSUM.RPT"
    foreach ($sheetName in $worksheetNames) {
        if ($sheetName -ne "RMDETMNT.RPT") {
            $newExcel.Workbook.Worksheets.Delete($sheetName)
        }
    }

    Close-ExcelPackage $newExcel -SaveAs $fullOutputPath

    Write-Host "Data has been successfully inserted into the Template worksheet and saved as '$filename'."
} #done

function HK_ALLSUM.RPT {
    param (
        [psobject]$csvData,
        [string]$templatePath,
        [string]$outputDir
    )


    # Load the Excel file
    $excel = Open-ExcelPackage -Path $templatePath

    # Select the "Tester" worksheet
    $worksheet = $excel.Workbook.Worksheets["HK_ALLSUM.RPT"]
  
    $bussinessDate = SafeGetDateString -date $csvData.cur_bus_date[0]
  
    $worksheet.Cells[1, 4].Value = Get-Date -Format "M/d/yyyy"
    $worksheet.Cells[2, 4].Value = $bussinessDate
    $worksheet.Cells[3, 4].Value = Get-Date -Format "h:mm tt"

    $worksheet.Cells[3, 6].Value = $csvData.property_id[0]
    $worksheet.Cells[3, 9].Value = $csvData.property_id[0]

    $worksheet.Cells[1, 6].Value = $csvData.property_id[0] + " - " + $csvData.property_name[0]

    $csvData = $csvData | Where-Object { $_.room_id -ne "" } | Where-Object { $null -ne $_.room_id }


    $statusRow = 8
    $statusColumn = 1
    foreach ($status in ($csvData.Title | Get-Unique )) {
        $worksheet.Cells[$statusRow, $statusColumn].Value = $status
        $worksheet.Cells[$statusRow, $statusColumn].Style.HorizontalAlignment = "Right"
        $worksheet.Cells[$statusRow, ($statusColumn + 1)].Value = "Rooms Summary"

        $roomcount = $($csvData | Where-Object { $_.Title -eq $status }).count
        if ($null -eq $roomCount) { $roomCount = 1 }
        $roomRow = $statusRow + 1
        for ($i = 0; $i -lt $roomCount; $i++) {
            $entry = ($csvData | Where-Object { $_.Title -eq $status })[$i]
            $value = $entry.room_id.TrimEnd() + $(if ($csvData.in_room_ind -like "False") {}else { "*" })
            $worksheet.Cells[($roomRow + $i), 1].Value = $value
        }
        $totalRow = $statusRow + $roomCount + 1
        $worksheet.Cells[$totalRow, 1].Value = "Total"
        $worksheet.Cells[$totalRow, 2].Value = $status
        $worksheet.Cells[$totalRow, 4].Value = $roomCount
        $worksheet.Cells[$totalRow, 4].Style.HorizontalAlignment = "Left"
        $statusRow += $roomCount + 2
    }
  
    $noteRow = $totalRow + 2
    $worksheet.Cells[$noteRow, 1].Value = "NOTE * = ATTENDANT IS IN ROOM"
    $worksheet.Cells[($noteRow + 1), 1].Value = "END OF REPORT"

    $filename = "OnQ - All Room Status Summary Report"

    # Define the new file name
    $filename = "$filename - $($csvData.property_id[0]) - $($(Get-Date $csvData.cur_bus_date[0]).ToString("M-d-yyyy")).xlsx"

    $fullOutputPath = Join-Path -Path $outputDir -ChildPath $filename

    # Save the entire Excel file with the new name
    Close-ExcelPackage $excel -SaveAs $fullOutputPath
  
    # Reopen the saved Excel file
    $newExcel = Open-ExcelPackage -Path $fullOutputPath
  
    # Get the list of all worksheet names
    $worksheetNames = $newExcel.Workbook.Worksheets.Name
  
    # Loop through all worksheets and delete those that are not "HK_ALLSUM.RPT"
    foreach ($sheetName in $worksheetNames) {
        if ($sheetName -ne "HK_ALLSUM.RPT") {
            $newExcel.Workbook.Worksheets.Delete($sheetName)
        }
    }

    Close-ExcelPackage $newExcel -SaveAs $fullOutputPath

    Write-Host "Data has been successfully inserted into the Template worksheet and saved as '$filename'."
} #done

function ARAGENUM.RPT {
    param (
        [psobject]$csvData,
        [string]$templatePath,
        [string]$outputDir
    )

    # Load the Excel file
    $excel = Open-ExcelPackage -Path $templatePath

    # Select the "Tester" worksheet
    $worksheet = $excel.Workbook.Worksheets["ARAGENUM.RPT"]
  
    $bussinessDate = SafeGetDateString -date $csvData.report_date[0]
  
    $worksheet.Cells[2, 4].Value = Get-Date -Format "M/d/yyyy"
    $worksheet.Cells[3, 4].Value = $bussinessDate
    $worksheet.Cells[4, 4].Value = Get-Date -Format "h:mm tt"

    $worksheet.Cells[2, 6].Value = $csvData.property_id[0] + " - " + $csvData.property_name[0]

    $worksheet.Cells[4, 6].Value = "as of: " + $bussinessDate

    $csvData = $csvData | Where-Object { $_.ar_acct_code -ne "" }

    $entriesCount = $csvData.Count
    $startRow = 11
    $startCol = 1
    $displayedHeaders = @('ar_acct_code', 'ar_acct_desc', 'acct_type_code', 'acct_type_desc', 'total_current_ar', 'total_31to60_ar', 'total_61to90_ar', 'total_91to120_ar', 'total_121to150_ar', 'total_Over150_ar', 'grand_balance')
    foreach ($header in $displayedHeaders) {
        for ($i = 0; $i -lt $entriesCount; $i++) {
            if ($header -eq 'acct_type_code' -or $header -eq 'acct_type_desc') {
                $value = $null
            }
            else {
                $value = SafeParseNumberandRound -value $csvData.$header[$i]
                if ($value -eq 0 -or $value -eq '0') {
                    $value = $null
                }
            }
            $worksheet.Cells[($startRow + $i), $startCol].Value = $value
        }
        $startCol++
    }

    $totals = @('total_current_ar', 'total_31to60_ar', 'total_61to90_ar', 'total_91to120_ar', 'total_121to150_ar', 'total_Over150_ar', 'grand_balance')
    $TotalsRow = ($startRow + $entriesCount)

    $worksheet.Cells[$TotalsRow, 1].Value = "TOTAL"
    $worksheet.Cells[$TotalsRow, 1].Style.Font.Bold = $true

   
    foreach ($total in $totals) {
        $value = sumList -list $csvData.$total
        if ($value -ne 0 -or $value -ne '0') {
            $index = $totals.IndexOf($total)
            $worksheet.Cells[$TotalsRow, (5 + $index)].Value = $value
            $worksheet.Cells[$TotalsRow, (5 + $index)].Style.Font.Bold = $true
        }
    }

 
    $worksheet.Cells[($TotalsRow + 2), 1].Value = "END OF REPORT"

    $startRow = $worksheet.Dimension.Start.Row
    $endRow = $worksheet.Dimension.End.Row
  
    for ($row = $startRow; $row -le $endRow; $row++) {
        $worksheet.Row($row).Height = 12.75
    }
  
    $filename = "OnQ - Aged Receivables By Account Number"

    # Define the new file name
    $filename = "$($filename) - $($csvData.property_id[0]) - $($(Get-Date $csvData.report_date[0]).ToString("M-d-yyyy")).xlsx"

    $fullOutputPath = Join-Path -Path $outputDir -ChildPath $filename

    # Save the entire Excel file with the new name
    Close-ExcelPackage $excel -SaveAs $fullOutputPath
  
    # Reopen the saved Excel file
    $newExcel = Open-ExcelPackage -Path $fullOutputPath
  
    # Get the list of all worksheet names
    $worksheetNames = $newExcel.Workbook.Worksheets.Name
  
    # Loop through all worksheets and delete those that are not "HK_ALLSUM.RPT"
    foreach ($sheetName in $worksheetNames) {
        if ($sheetName -ne "ARAGENUM.RPT") {
            $newExcel.Workbook.Worksheets.Delete($sheetName)
        }
    }

    Close-ExcelPackage $newExcel -SaveAs $fullOutputPath

    Write-Host "Data has been successfully inserted into the Template worksheet and saved as '$filename'."
} #i think this is good

function GPBALSUM.RPT {
    param (
        [psobject]$csvData,
        [string]$templatePath,
        [string]$outputDir
    )

 
    # Load the Excel file
    $excel = Open-ExcelPackage -Path $templatePath

    # Select the "Tester" worksheet
    $worksheet = $excel.Workbook.Worksheets["GPBALSUM.RPT"]
  
    $bussinessDate = SafeGetDateString -date $csvData.cur_bus_date[0]
  
    $worksheet.Cells[1, 4].Value = Get-Date -Format "M/d/yyyy"
    $worksheet.Cells[2, 4].Value = $bussinessDate
    $worksheet.Cells[4, 4].Value = Get-Date -Format "h:mm tt"

    $worksheet.Cells[1, 6].Value = $csvData.property_id[0] + " - " + $csvData.property_name[0] + " - " + $csvData.facility_id[0]

    $csvData = $csvData | Where-Object { $_.group_code -ne "" } | Where-Object { $null -ne $_.group_code }

    $entriesCount = $csvData.Count
    $startRow = 9
    $startCol = 1
    $headers = @(
        "group_code",
        "group_name",
        "blank1",
        "contact_person",
        "blank2",
        "definite_ind",
        "begin_date",
        "end_date",
        "Receipt",
        "folio_bal",
        "mop_type",
        "special_rate_plan"
    )

    foreach ($header in $headers) {
        for ($i = 0; $i -lt $entriesCount; $i++) {
            if ($header -eq "mop_type") {
                switch ($csvData.$header[$i]) {
                    "CS" { $value = "CASH" }
                    "CC" { $value = "CREDIT CARD" }
                    "DY" { $value = "DEPOSITORY PAYMENT" }
                    "DB" { $value = "DIRECT BILL" }
                    "CH" { $value = "CHECK" }
                    default { $value = "NOT GUARENTEED" }
                } 
            }
            else {
                switch ($header) {
                    "blank1" { $value = $null }
                    "blank2" { $value = $null }
                    "begin_date" { $value = SafeGetDateString -date $csvData.$header[$i] }
                    "end_date" { $value = SafeGetDateString -date $csvData.$header[$i] }
                    "definite_ind" { $value = if ($csvData.$header[$i] -eq "TRUE") { "ACTIVE" } else { "NOT ACTIVE" } }
                    default { $value = SafeParseNumberandRound -value $csvData.$header[$i] }
                } 
            }
            $spaceRows = 4
            if ($i -gt 0) {
                if ($csvData."folio_id"[($i - 1)] -eq $csvData."folio_id"[($i)]) {
                    $spaceRows = 2
                }
            }
            else {
                $previousEntryRow = $startRow
            }
            $worksheet.Cells[($previousEntryRow + $spaceRows), $startCol].Value = $value
            $previousEntryRow = $previousEntryRow + $spaceRows
        }
        $startCol++
    }


    $worksheet.Cells[($previousEntryRow + 4), 1].Value = "   TOTAL GROUP BALANCES :"
    $worksheet.Cells[($previousEntryRow + 4), 5].Value = sumList -list $csvData.folio_bal
    $worksheet.Cells[($previousEntryRow + 4), 6].Value = "(EXCLUDES ADV DEP)"
    $worksheet.Cells[($previousEntryRow + 6), 1].Value = "  *GROUPS WITH ADV. DEP :"
    $worksheet.Cells[($previousEntryRow + 6), 5].Value = sumList -list $csvData.AdvanceDepositTotal
    $worksheet.Cells[($previousEntryRow + 6), 6].Value = "(ALSO INCLUDED IN ADV DEP TRAY)"
    $worksheet.Cells[($previousEntryRow + 8), 1].Value = "       TOTAL GROUP TRAY :"
    $worksheet.Cells[($previousEntryRow + 8), 5].Value = ($(sumList -list $csvData.folio_bal) + $(sumList -list $csvData.AdvanceDepositTotal))
    $worksheet.Cells[($previousEntryRow + 13), 1].Value = "END OF REPORT"

    $startRow = $worksheet.Dimension.Start.Row
    $endRow = $worksheet.Dimension.End.Row
  
    for ($row = $startRow; $row -le $endRow; $row++) {
        $worksheet.Row($row).Height = 12.75
    }
  
    $filename = "OnQ - Group Master Balance Summary Report"

    # Define the new file name
    $filename = "$($filename) - $($csvData.property_id[0]) - $($(Get-Date $csvData.cur_bus_date[0]).ToString("M-d-yyyy")).xlsx"

    $fullOutputPath = Join-Path -Path $outputDir -ChildPath $filename

    # Save the entire Excel file with the new name
    Close-ExcelPackage $excel -SaveAs $fullOutputPath
  
    # Reopen the saved Excel file
    $newExcel = Open-ExcelPackage -Path $fullOutputPath
  
    # Get the list of all worksheet names
    $worksheetNames = $newExcel.Workbook.Worksheets.Name
  
    # Loop through all worksheets and delete those that are not "HK_ALLSUM.RPT"
    foreach ($sheetName in $worksheetNames) {
        if ($sheetName -ne "GPBALSUM.RPT") {
            $newExcel.Workbook.Worksheets.Delete($sheetName)
        }
    }

    Close-ExcelPackage $newExcel -SaveAs $fullOutputPath

    Write-Host "Data has been successfully inserted into the Template worksheet and saved as '$filename'."
} #i think this is good

function RMBALSUM.RPT {
    param (
        [psobject]$csvData,
        [string]$templatePath,
        [string]$outputDir
    )

    # Load the Excel file
    $excel = Open-ExcelPackage -Path $templatePath

    # Select the "Tester" worksheet
    $worksheet = $excel.Workbook.Worksheets["RMBALSUM.RPT"]
  
    $bussinessDate = SafeGetDateString -date $csvData.cur_bus_date[0]
  
    $worksheet.Cells[1, 4].Value = Get-Date -Format "M/d/yyyy"
    $worksheet.Cells[3, 4].Value = $bussinessDate
    $worksheet.Cells[5, 4].Value = Get-Date -Format "h:mm tt"
    $worksheet.Cells[5, 6].Value = "FOR " + $bussinessDate

    #$worksheet.Cells[8, 6].Value = $(if ($csvData.adv_dep_ind -notcontains "FALSE") { "An advance deposit has been detected when none were expected. Please reach out to Brady Denton and Re-Run this report from OnQ" }else { $null })
    $script:propertyIDHeader = $csvData.property_id[0] + " - " + $csvData.property_name[0] + " - " + $csvData.facility_id[0]
    $worksheet.Cells[1, 6].Value = $csvData.property_id[0] + " - " + $csvData.property_name[0] + " - " + $csvData.facility_id[0]

    $csvData = $csvData | Where-Object { $_.stay_id -ne "" } | Where-Object { $null -ne $_.stay_id }

    $entriesCount = $csvData.Count
    $startRow = 26
    $startCol = 1
    $headers = @(
        "room_id",
        "share",
        "guest_status_short_desc",
        "guest_firstname",
        "guest_arrival_date",
        "guest_departure_date",
        "receipt_A_BAL",
        "receipt_A_MOP",
        "receipt_B_BAL",
        "receipt_B_MOP",
        "receipt_C_BAL",
        "receipt_C_MOP",
        "entry_amount"
    )

    foreach ($header in $headers) {
        for ($i = 0; $i -lt $entriesCount; $i++) {
            switch ($header) {
                "room_id" { $value = $csvData.$header[$i] }
                "share" { $value = if ($csvData.$header[$i] -eq "TRUE") { "Y" } }
                "guest_firstname" { $value = $csvData.guest_lastname[$i] + "/" + $csvData.$header[$i] }
                "guest_arrival_date" { $value = SafeGetDateString -date $csvData.$header[$i] }
                "guest_departure_date" { $value = SafeGetDateString -date $csvData.$header[$i] }
                default { $value = SafeParseNumberandRound -value $csvData.$header[$i] }
            } 
      
            if ($header -eq "room_id" -and $csvData.stay_id[$i] -eq $csvData.stay_id[($i - 1)] -and $csvData.room_id[$i] -eq $csvData.room_id[($i - 1)]) {
                $value = $null
            }
            $entryRow = $startRow + ($i * 2)
            $worksheet.Cells[($entryRow), $startCol].Value = $value
            switch ($header) {
                "receipt_A_BAL" {
                    $worksheet.Cells[($entryRow), $startCol].Style.Numberformat.Format = "$#,##0.00_);($#,##0.00)"
                    $worksheet.Cells[($entryRow), $startCol].Style.HorizontalAlignment = "Right"
                }
                "receipt_B_BAL" {
                    $worksheet.Cells[($entryRow), $startCol].Style.Numberformat.Format = "$#,##0.00_);($#,##0.00)"
                    $worksheet.Cells[($entryRow), $startCol].Style.HorizontalAlignment = "Right"
                }
                "receipt_C_BAL" {
                    $worksheet.Cells[($entryRow), $startCol].Style.Numberformat.Format = "$#,##0.00_);($#,##0.00)"
                    $worksheet.Cells[($entryRow), $startCol].Style.HorizontalAlignment = "Right"
                }
                "entry_amount" {
                    $worksheet.Cells[($entryRow), $startCol].Style.Numberformat.Format = "$#,##0.00_);($#,##0.00)"
                    $worksheet.Cells[($entryRow), $startCol].Style.HorizontalAlignment = "Right"
                }
                default {}
            }
        }
        $startCol++
    }

    $receipt_A = sumList -list $csvData.receipt_A_BAL
    $receipt_B = sumList -list $csvData.receipt_B_BAL
    $receipt_C = sumList -list $csvData.receipt_C_BAL

    $balanceRow = $entryRow + 3
    $worksheet.Cells[$balanceRow, 6].Value = "BALANCE :"
    $worksheet.Cells[$balanceRow, 7].Value = $receipt_A
    $worksheet.Cells[$balanceRow, 7].Style.Numberformat.Format = "$#,##0.00_);($#,##0.00)"
    $worksheet.Cells[$balanceRow, 9].Value = $receipt_B
    $worksheet.Cells[$balanceRow, 9].Style.Numberformat.Format = "$#,##0.00_);($#,##0.00)"
    $worksheet.Cells[$balanceRow, 11].Value = $receipt_C
    $worksheet.Cells[$balanceRow, 11].Style.Numberformat.Format = "$#,##0.00_);($#,##0.00)"
    $worksheet.Cells[$balanceRow, 13].Value = sumList -list $csvData.entry_amount
    $worksheet.Cells[$balanceRow, 13].Style.Numberformat.Format = "$#,##0.00_);($#,##0.00)"

    $balanceRow = $entryRow + 9
    $worksheet.Cells[($balanceRow), 3].Value = "TOTAL RM BALANCES :"
    $worksheet.Cells[($balanceRow), 3].Style.HorizontalAlignment = "Right"
    $worksheet.Cells[($balanceRow), 4].Value = $receipt_C + $receipt_B + $receipt_A
    $worksheet.Cells[($balanceRow), 4].Style.HorizontalAlignment = "Right"
    $worksheet.Cells[($balanceRow), 4].Style.Numberformat.Format = "$#,##0.00_);($#,##0.00)"
    $worksheet.Cells[($balanceRow), 5].Value = "(EXCLUDES ADV DEP)"
    $worksheet.Cells[($balanceRow), 5].Style.HorizontalAlignment = "Left"


    $worksheet.Cells[($balanceRow + 2), 3].Value = "*GSTS WITH ADV. DEP :"
    $worksheet.Cells[($balanceRow + 2), 3].Style.HorizontalAlignment = "Right"
    $worksheet.Cells[($balanceRow + 2), 4].Value = 0
    $worksheet.Cells[($balanceRow + 2), 4].Style.HorizontalAlignment = "Right"
    $worksheet.Cells[($balanceRow + 2), 4].Style.Numberformat.Format = "$#,##0.00_);($#,##0.00)"
    $worksheet.Cells[($balanceRow + 2), 5].Value = "(ALSO INCLUDED IN ADV DEP TRAY)"
    $worksheet.Cells[($balanceRow + 2), 5].Style.HorizontalAlignment = "Left"


    $worksheet.Cells[($balanceRow + 4), 3].Value = "TOTAL ROOM TRAY :"
    $worksheet.Cells[($balanceRow + 4), 3].Style.HorizontalAlignment = "Right"
    $worksheet.Cells[($balanceRow + 4), 4].Value = sumList -list $csvData.entry_amount
    $worksheet.Cells[($balanceRow + 4), 4].Style.HorizontalAlignment = "Right"
    $worksheet.Cells[($balanceRow + 4), 4].Style.Numberformat.Format = "$#,##0.00_);($#,##0.00)"

    $worksheet.Cells[($balanceRow + 6), 1].Value = "END OF REPORT"

    $startRow = $worksheet.Dimension.Start.Row
    $endRow = $worksheet.Dimension.End.Row
  
    for ($row = $startRow; $row -le $endRow; $row++) {
        $worksheet.Row($row).Height = 12.75
    }
  
    $filename = "OnQ - Room Balance Summary"

    # Define the new file name
    $filename = "$($filename) - $($csvData.property_id[0]) - $($(Get-Date $csvData.cur_bus_date[0]).ToString("M-d-yyyy")).xlsx"

    $fullOutputPath = Join-Path -Path $outputDir -ChildPath $filename

    # Save the entire Excel file with the new name
    Close-ExcelPackage $excel -SaveAs $fullOutputPath
  
    # Reopen the saved Excel file
    $newExcel = Open-ExcelPackage -Path $fullOutputPath
  
    # Get the list of all worksheet names
    $worksheetNames = $newExcel.Workbook.Worksheets.Name
  
    # Loop through all worksheets and delete those that are not "HK_ALLSUM.RPT"
    foreach ($sheetName in $worksheetNames) {
        if ($sheetName -ne "RMBALSUM.RPT") {
            $newExcel.Workbook.Worksheets.Delete($sheetName)
        }
    }

    Close-ExcelPackage $newExcel -SaveAs $fullOutputPath

    Write-Host "Data has been successfully inserted into the Template worksheet and saved as '$filename'."
} #i think this is good

function OOBALSUM.RPT {
    param (
        [psobject]$csvData,
        [string]$templatePath,
        [string]$outputDir
    )

    # Load the Excel file
    $excel = Open-ExcelPackage -Path $templatePath

    # Select the "Tester" worksheet
    $worksheet = $excel.Workbook.Worksheets["OOBALSUM.RPT"]
  
    $bussinessDate = SafeGetDateString -date $csvData.Cur_Bus_Date[0]
    $saveDate = $(Get-Date $csvData.Cur_Bus_Date[0]).ToString("M-d-yyyy")
    $propertyID = $csvData.Property_Id[0]
  
    $worksheet.Cells[2, 4].Value = Get-Date -Format "M/d/yyyy"
    $worksheet.Cells[3, 4].Value = $bussinessDate
    $worksheet.Cells[4, 4].Value = Get-Date -Format "h:mm tt"

    $worksheet.Cells[2, 6].Value = $csvData.Property_Id[0] + "-" + $csvData.Property_Name[0]
  
    $csvData = $csvData | Where-Object { $_.SortOrder -ne "" } | Where-Object { $null -ne $_.SortOrder }
 
    $headers = @(
        "room_id_or_account",
        "blank1",
        "blank2",
        "name",
        "blank3",
        "blank4",
        "blank5",
        "folio_balance",
        "arrival_date",
        "departure_date"
    )


    $startCol = 1
    $startRow = 11
    $currentRow = $startRow
    $oobAccountTypes = $csvData.SortOrder | Select-Object -Unique
    $categoryOpeners = @()
    $categoryClosers = @()

    $checkedoutOOB = 0
    $noshowOOB = 0
    $groupOOB = 0
    $haOOB = 0
    $saOOB = 0

    if ($csvData.SortOrder.Count -gt 0) {
        foreach ($accountType in $oobAccountTypes) {
            switch ($accountType) {
                "A" {
                    if ($csvData | Where-Object { $_.status -ne "Canceled" } | Where-Object { $_.SortOrder -eq "A" }) {
                        $currentRow += 2
                        $worksheet.Cells[$currentRow, 1].Value = "CHECKED OUT ROOMS"
                        $currentRow += 2
                        foreach ($entry in ($csvData | Where-Object { $_.status -ne "Canceled" } | Where-Object { $_.SortOrder -eq "A" })) {
                            foreach ($header in $headers) {
                                $index = ($headers.IndexOf($header)) + 1
                                switch ($header) {
                                    "arrival_date" { $value = SafeGetDateString -date $entry.$header }
                                    "departure_date" { $value = SafeGetDateString -date $entry.$header }
                                    "folio_balance" {
                                        $value = SafeParseNumberandRound -value $entry.$header 
                                        $worksheet.Cells[$currentRow, $index].Style.Numberformat.Format = "$#,##0.00_);($#,##0.00)"
                                    }
                                    "room_id_or_account" { $value = $entry.$header }
                                    "name" { $value = $entry.$header }
                                    default { $value = $null }
                                }
                                $worksheet.Cells[$currentRow, $index].Value = $value
                            } 
                            $currentRow++
                        }
                        $currentRow += 2
                        $worksheet.Cells[$currentRow, 1].Value = "CHECKED OUT ROOMS BALANCE:"
                        $checkedoutOOB = sumList -list $($csvData | Where-Object { $_.status -ne "Canceled" } | Where-Object { $_.SortOrder -eq "A" }).folio_balance
                        $worksheet.Cells[$currentRow, 8].Value = $checkedoutOOB
                        $worksheet.Cells[$currentRow, 8].Style.Numberformat.Format = "$#,##0.00_);($#,##0.00)"
                    }
                    if ($csvData | Where-Object { $_.status -eq "Canceled" } | Where-Object { $_.SortOrder -eq "A" }) {
                        $currentRow += 2
                        $worksheet.Cells[$currentRow, 1].Value = "NO SHOWS/CANCELS"
                        $currentRow += 2
                        foreach ($entry in ($csvData | Where-Object { $_.status -eq "Canceled" } | Where-Object { $_.SortOrder -eq "A" })) {
                            foreach ($header in $headers) {
                                $index = ($headers.IndexOf($header)) + 1
                                switch ($header) {
                                    "arrival_date" { $value = SafeGetDateString -date $entry.$header }
                                    "departure_date" { $value = SafeGetDateString -date $entry.$header }
                                    "folio_balance" {
                                        $value = SafeParseNumberandRound -value $entry.$header 
                                        $worksheet.Cells[$currentRow, $index].Style.Numberformat.Format = "$#,##0.00_);($#,##0.00)"
                                    }
                                    "room_id_or_account" { $value = $entry.$header }
                                    "name" { $value = $entry.$header }
                                    default { $value = $null }
                                }
                                $worksheet.Cells[$currentRow, $index].Value = $value
                            } 
                            $currentRow++
                        }
                        $currentRow += 2
                        $worksheet.Cells[$currentRow, 1].Value = "NO SHOWS/CANCELS RECORDS BALANCE:"
                        $noshowOOB = sumList -list $($csvData | Where-Object { $_.status -eq "Canceled" } | Where-Object { $_.SortOrder -eq "A" }).folio_balance
                        $worksheet.Cells[$currentRow, 8].Value = $noshowOOB
                        $worksheet.Cells[$currentRow, 8].Style.Numberformat.Format = "$#,##0.00_);($#,##0.00)"
                    }
                }
                "B" {
                    $currentRow += 2
                    $worksheet.Cells[$currentRow, 1].Value = "CLOSED GROUP MASTERS WITH BALANCES:"
                    $currentRow += 2
                    foreach ($entry in ($csvData | Where-Object { $_.SortOrder -eq "B" })) {
                        foreach ($header in $headers) {
                            $index = ($headers.IndexOf($header)) + 1
                            switch ($header) {
                                "arrival_date" { $value = SafeGetDateString -date $entry.$header }
                                "departure_date" { $value = SafeGetDateString -date $entry.$header }
                                "folio_balance" {
                                    $value = SafeParseNumberandRound -value $entry.$header 
                                    $worksheet.Cells[$currentRow, $index].Style.Numberformat.Format = "$#,##0.00_);($#,##0.00)"
                                }
                                "room_id_or_account" { $value = $entry.$header }
                                "name" { $value = $entry.$header }
                                default { $value = $null }
                            }
                            $worksheet.Cells[$currentRow, $index].Value = $value
                      
                        } 
                    }
                    $currentRow += 2
                    $worksheet.Cells[$currentRow, 1].Value = "CLOSED GROUP MASTER BALANCE TOTALS:"
                    $groupOOB = sumList -list $($csvData | Where-Object { $_.SortOrder -eq "B" }).folio_balance
                    $worksheet.Cells[$currentRow, 8].Value = $groupOOB
                    $worksheet.Cells[$currentRow, 8].Style.Numberformat.Format = "$#,##0.00_);($#,##0.00)"
                }
                "C" {
                    $currentRow += 2
                    $worksheet.Cells[$currentRow, 1].Value = "CLOSED HOUSE ACCOUNT WITH BALANCES:"
                    $currentRow += 2
                    foreach ($entry in ($csvData | Where-Object { $_.SortOrder -eq "C" })) {
                        foreach ($header in $headers) {
                            $index = ($headers.IndexOf($header)) + 1
                            switch ($header) {
                                "arrival_date" { $value = SafeGetDateString -date $entry.$header }
                                "departure_date" { $value = SafeGetDateString -date $entry.$header }
                                "folio_balance" {
                                    $value = SafeParseNumberandRound -value $entry.$header 
                                    $worksheet.Cells[$currentRow, $index].Style.Numberformat.Format = "$#,##0.00_);($#,##0.00)"
                                }
                                "room_id_or_account" { $value = $entry.$header }
                                "name" { $value = $entry.$header }
                                default { $value = $null }
                            }
                            $worksheet.Cells[$currentRow, $index].Value = $value
                     
                        } 
                        $currentRow++
                    }
                    $currentRow += 2
                    $worksheet.Cells[$currentRow, 1].Value = "CLOSED HOUSE ACCOUNT BALANCE TOTALS:"
                    $haOOB = sumList -list $($csvData | Where-Object { $_.SortOrder -eq "C" }).folio_balance 
                    $worksheet.Cells[$currentRow, 8].Value = $haOOB
                    $worksheet.Cells[$currentRow, 8].Style.Numberformat.Format = "$#,##0.00_);($#,##0.00)"
                }
                "D" {
                    $currentRow += 2
                    $worksheet.Cells[$currentRow, 1].Value = "SYSTEM ACCOUNTS WITH BALANCES:"
                    $currentRow += 2
                    foreach ($entry in ($csvData | Where-Object { $_.SortOrder -eq "D" })) {
                        foreach ($header in $headers) {
                            $index = ($headers.IndexOf($header)) + 1
                            switch ($header) {
                                "arrival_date" { $value = SafeGetDateString -date $entry.$header }
                                "departure_date" { $value = SafeGetDateString -date $entry.$header }
                                "folio_balance" {
                                    $value = SafeParseNumberandRound -value $entry.$header 
                                    $worksheet.Cells[$currentRow, $index].Style.Numberformat.Format = "$#,##0.00_);($#,##0.00)"
                                }
                                "room_id_or_account" { $value = $entry.$header }
                                "name" { $value = $entry.$header }
                                default { $value = $null }
                            }
                            $worksheet.Cells[$currentRow, $index].Value = $value
                        } 
                        $currentRow++
                    }
                    $currentRow += 2
                    $worksheet.Cells[$currentRow, 1].Value = "SYSTEM ACCOUNTS WITH BALANCES TOTAL:"
                    $saOOB = sumList -list $($csvData | Where-Object { $_.SortOrder -eq "D" }).folio_balance 
                    $worksheet.Cells[$currentRow, 8].Value = $saOOB
                    $worksheet.Cells[$currentRow, 8].Style.Numberformat.Format = "$#,##0.00_);($#,##0.00)"
                }
            }
        }
    }
    else { Write-Host "No entries for Reconciliation" -ForegroundColor Blue }
    $currentRow += 2
    $totalOOB = $checkedoutOOB + $noshowOOB + $groupOOB + $haOOB + $saOOB
    $worksheet.Cells[$currentRow, 1].Value = "OUT OF BALANCE SUMMARY TOTAL:"
    $worksheet.Cells[$currentRow, 8].Value = $totalOOB
    $worksheet.Cells[$currentRow, 8].Style.Numberformat.Format = "$#,##0.00_);($#,##0.00)"
    $currentRow ++
    $worksheet.Cells[$currentRow, 1].Value = "END OF REPORT"

    $filename = "OnQ - Reconciliation"

    # Define the new file name
    $filename = "$($filename) - $propertyID - $saveDate.xlsx"

    $fullOutputPath = Join-Path -Path $outputDir -ChildPath $filename

    # Save the entire Excel file with the new name
    Close-ExcelPackage $excel -SaveAs $fullOutputPath
  
    # Reopen the saved Excel file
    $newExcel = Open-ExcelPackage -Path $fullOutputPath
  
    # Get the list of all worksheet names
    $worksheetNames = $newExcel.Workbook.Worksheets.Name
  
    # Loop through all worksheets and delete those that are not "HK_ALLSUM.RPT"
    foreach ($sheetName in $worksheetNames) {
        if ($sheetName -ne "OOBALSUM.RPT") {
            $newExcel.Workbook.Worksheets.Delete($sheetName)
        }
    }

    Close-ExcelPackage $newExcel -SaveAs $fullOutputPath

    Write-Host "Data has been successfully inserted into the Template worksheet and saved as '$filename'."
}#done

function DPAUDDAA.RPT {
    param (
        [psobject]$csvData,
        [string]$templatePath,
        [string]$outputDir
    )

    # Load the Excel file
    $excel = Open-ExcelPackage -Path $templatePath

    # Select the "Tester" worksheet
    $worksheet = $excel.Workbook.Worksheets["DPAUDDAA.RPT"]

    $bussinessDate = $(if (![string]::IsNullOrWhiteSpace($csvData.entry_date)) { SafeGetDateString -date $csvData.entry_date[0] }else { $($($script:goLiveDate | Get-Date).ToString("M/d/yyyy")) })
 
    $worksheet.Cells[1, 4].Value = Get-Date -Format "M/d/yyyy"
    $worksheet.Cells[2, 4].Value = Get-Date -Format "h:mm tt"

    $worksheet.Cells[1, 6].Value = "$inncode"
    $worksheet.Cells[2, 6].Value = "ACCOUNT DETAIL REPORT FOR " + $bussinessDate

    $responseRecevied = $(if ($csvData) { $true }else { $false })

    $csvData = $csvData | Where-Object { $_.entry_id -ne "" } | Where-Object { $null -ne $_.entry_id }
    $csvData = $csvData | ForEach-Object {
        $_.PSObject.Properties | ForEach-Object {
            if ($_.Value -is [string]) {
                $_.Value = $_.Value.Trim()
            }
        }
        $_
    }


    $postingHeaders = @("index", "room_id", "employee_id", "entry_date", "entry_time", "entry_description", "revenue_amount", "adjustment_amount", "trans_id", "entry_ref_id", "conf_num")
    $accountHeaders = @("LINE ITEMS:", "DATE", "", "REVENUE", "ADJUSTMENTS", "", "NET")

    $accountIDs = $csvData.accounting_id | Select-Object -Unique

    $startRow = 15
    $entryRow = $startRow
 
    if (![string]::IsNullOrWhiteSpace($csvData.entry_date)) {
        foreach ($accountID in $accountIDs) {
            $accountPostings = $csvData | Where-Object { $_.accounting_id.Trim() -eq $accountID }
            $worksheet.Cells[($entryRow), 1].Value = ($accountPostings.category_type_desc | Select-Object -Unique) + " - " + ($accountPostings.accounting_id_desc | Select-Object -Unique) + " - (" + ($accountPostings.category_id | Select-Object -Unique) + ")"
            $worksheet.Cells[($entryRow), 1].Style.HorizontalAlignment = "Left"
            $entryRow += 2
            $headerColumn = 3
            foreach ($header in $accountHeaders) {
                switch ($header) {
                    "LINE ITEMS:" { $worksheet.Cells[($entryRow), $headerColumn].Style.HorizontalAlignment = "Center" }
                    "DATE" { $worksheet.Cells[($entryRow), $headerColumn].Style.HorizontalAlignment = "Left" }
                    "REVENUE" { $worksheet.Cells[($entryRow), $headerColumn].Style.HorizontalAlignment = "Right" }
                    "ADJUSTMENTS" { $worksheet.Cells[($entryRow), $headerColumn].Style.HorizontalAlignment = "Right" }
                    "NET" { $worksheet.Cells[($entryRow), $headerColumn].Style.HorizontalAlignment = "Right" }
                    Default { $null }
                }
                $worksheet.Cells[($entryRow), $headerColumn].Value = $header
                $headerColumn ++
            }
            $entryRow += 1
            $headerColumn = 3
            foreach ($header in $accountHeaders) {
                $worksheet.Cells[($entryRow), $headerColumn].Style.HorizontalAlignment = "Right"           
                switch ($header) {
                    "LINE ITEMS:" { $value = $(if ($null -eq $accountPostings.Count) { 1 }else { $accountPostings.Count }) }
                    "DATE" { $value = $bussinessDate }
                    "REVENUE" {
                        $value = $(sumList -list $accountPostings.revenue_amount)
                        $worksheet.Cells[($entryRow), $headerColumn].Style.Numberformat.Format = "$#,##0.00_);($#,##0.00)"
                    }
                    "ADJUSTMENTS" {
                        $value = sumList -list $accountPostings.adjustment_amount
                        $worksheet.Cells[($entryRow), $headerColumn].Style.Numberformat.Format = "$#,##0.00_);($#,##0.00)"
                    }
                    "NET" {
                        $value = (sumList -list $accountPostings.revenue_amount) - (sumList -list $accountPostings.adjustment_amount)
                        $worksheet.Cells[($entryRow), $headerColumn].Style.Numberformat.Format = "$#,##0.00_);($#,##0.00)"
                    }
                    Default { $value = $null }
                }
                $worksheet.Cells[($entryRow), $headerColumn].Value = $value
                $headerColumn ++
            }
            $entryRow += 1
            $worksheet.Cells[($entryRow), 1].Value = "ROOM #"
            $worksheet.Cells[($entryRow), 1].Style.HorizontalAlignment = "Left"
            $entryRow += 2
            $worksheet.Cells[($entryRow), 1].Value = "GROUP #"
            $worksheet.Cells[($entryRow), 1].Style.HorizontalAlignment = "Left"
            $entryRow += 1
            $worksheet.Cells[($entryRow), 1].Value = "AR ACCT #"
            $worksheet.Cells[($entryRow), 1].Style.HorizontalAlignment = "Left"
            $entryRow += 2
            $worksheet.Cells[($entryRow), 5].Value = "DESCRIPTION"
            $worksheet.Cells[($entryRow), 5].Style.HorizontalAlignment = "Left"
            $entryRow += 1
            $headerColumn = 1
            foreach ($header in $postingHeaders) {
                switch ($header) {
                    "index" {
                        $align = "Center"
                        $value = "ITEM"
                    }
                    "room_id" {
                        $align = "Center"
                        $value = "HOUSE ACCT #"
                    }
                    "employee_id" {
                        $align = "Left"
                        $value = "USER ID"
                    }
                    "entry_date" {
                        $align = "Left"
                        $value = "DATE"
                    }
                    "entry_time" {
                        $align = "Left"
                        $value = "TIME"
                    }
                    "entry_description" {
                        $align = "Left"
                        $value = $null
                    }
                    "revenue_amount" {
                        $align = "Center"
                        $value = "AMOUNT"
                    }
                    "adjustment_amount" {
                        $align = "Center"
                        $value = "ADJUSTMENT"
                    }
                    "trans_id" {
                        $align = "Left"
                        $value = "TRANS ID"
                    }
                    "entry_ref_id" {
                        $align = "Right"
                        $value = "POS"
                    }
                    "conf_num" {
                        $align = "Left"
                        $value = "CONF #"
                    }              
                }
                $worksheet.Cells[($entryRow), $headerColumn].Style.HorizontalAlignment = $align
                $worksheet.Cells[($entryRow), $headerColumn].Value = $value
                $headerColumn ++
            }
            $entryRow += 1
            $index = 1
            foreach ($posting in $accountPostings) {
                $headerColumn = 1
                foreach ($header in $postingHeaders) {
                    switch ($header) {
                        "index" {
                            $align = "Center"
                            $value = $index
                        }
                        "room_id" {
                            $align = "Center"
                            $value = $posting.$header
                        }
                        "employee_id" {
                            $align = "Left"
                            $value = $posting.$header
                        }
                        "entry_date" {
                            $align = "Center"
                            $value = SafeGetDateString -date $posting.$header
                        }
                        "entry_time" {
                            $align = "Center"
                            $value = Get-FormattedTime -dateTimeString $posting.entry_date
                        }
                        "entry_description" {
                            $align = "Left"
                            $value = $posting.$header
                        }
                        "revenue_amount" {
                            $align = "Center"
                            $value = SafeParseNumberandRound -value $posting.$header
                        }
                        "adjustment_amount" {
                            $align = "Center"
                            $value = SafeParseNumberandRound -value $posting.$header
                            $worksheet.Cells[($entryRow), $headerColumn].Style.Numberformat.Format = "$#,##0.00_);($#,##0.00)"
                        }
                        "trans_id" {
                            $align = "Center"
                            $value = $posting.$header
                        }
                        "entry_ref_id" {
                            $align = "Right"
                            $value = $posting.$header
                        }
                        "conf_num" {
                            $align = "Left"
                            $value = $posting.$header
                        }              
                    }
                    $worksheet.Cells[($entryRow), $headerColumn].Style.HorizontalAlignment = $align
                    $worksheet.Cells[($entryRow), $headerColumn].Value = $value
                    $headerColumn ++
                }
                $index += 1
                $entryRow += 1
            }
            $entryRow += 1       
            $worksheet.Cells[($entryRow), 1].Value = "TOTALS:"
            $worksheet.Cells[($entryRow), 1].Style.HorizontalAlignment = "Left"
            $value = ($accountPostings.accounting_id_desc | Select-Object -Unique) + "  Sub Total Net Rev:"
            $worksheet.Cells[($entryRow), 6].Value = $value
            $worksheet.Cells[($entryRow), 6].Style.HorizontalAlignment = "Right"
            $value = (sumList -list $accountPostings.revenue_amount) - (sumList -list $accountPostings.adjustment_amount)
            $worksheet.Cells[($entryRow), 8].Value = $value
            $worksheet.Cells[($entryRow), 8].Style.Numberformat.Format = "$#,##0.00_);($#,##0.00)"
            $worksheet.Cells[($entryRow), 8].Style.Font.UnderLine = [OfficeOpenXml.Style.ExcelUnderLineType]::Double
            $entryRow += 4
        }
    }
    elseif ($responseRecevied) { $worksheet.Cells[9, 1].Value = "This page left intentionally blank." }else { $worksheet.Cells[9, 1].Value = "No response received from OnQ server, please re-run this report on Hotel Hardware." }

    $worksheet.Cells[$entryRow, 1].Value = "END OF REPORT"


    $startRow = $worksheet.Dimension.Start.Row
    $endRow = $worksheet.Dimension.End.Row
    for ($row = $startRow; $row -le $endRow; $row++) {
        $worksheet.Row($row).Height = 12.75
    }

    $filename = "OnQ - Account Detail"

    # Define the new file name
    $filename = "$($filename) - $inncode - $($(Get-Date $bussinessDate).ToString("M-d-yyyy")).xlsx"

    $fullOutputPath = Join-Path -Path $outputDir -ChildPath $filename

    # Save the entire Excel file with the new name
    Close-ExcelPackage $excel -SaveAs $fullOutputPath
  
    # Reopen the saved Excel file
    $newExcel = Open-ExcelPackage -Path $fullOutputPath
  
    # Get the list of all worksheet names
    $worksheetNames = $newExcel.Workbook.Worksheets.Name
  
    # Loop through all worksheets and delete those that are not "HK_ALLSUM.RPT"
    foreach ($sheetName in $worksheetNames) {
        if ($sheetName -ne "DPAUDDAA.RPT") {
            $newExcel.Workbook.Worksheets.Delete($sheetName)
        }
    }

    Close-ExcelPackage $newExcel -SaveAs $fullOutputPath

    Write-Host "Data has been successfully inserted into the Template worksheet and saved as '$filename'."
}# make this handle no entries

function ADDONFULL.RPT {
    param (
        [psobject]$csvData,
        [string]$templatePath,
        [string]$outputDir
    )

 
    # Load the Excel file
    $excel = Open-ExcelPackage -Path $templatePath

    # Select the "Tester" worksheet
    $worksheet = $excel.Workbook.Worksheets["ADDONFULL.RPT"]
  
    $bussinessDate = SafeGetDateString -date $csvData.STARTDATE[0]
    $endDate = SafeGetDateString -date $csvData.ENDDATE[0]
  
    $worksheet.Cells[2, 3].Value = Get-Date -Format "M/d/yyyy"
    $worksheet.Cells[3, 3].Value = $bussinessDate
    $worksheet.Cells[4, 3].Value = Get-Date -Format "h:mm tt"

    #$worksheet.Cells["E2:N2"].Merge()
    #$worksheet.Cells["E3:N3"].Merge()
    $worksheet.Cells[2, 5].Value = $csvData.PROPERTYID[0] + "-" + $csvData.PropertyName[0] + "-" + $csvData.FacilityID[0]
    $worksheet.Cells[3, 5].Value = "ADD-ON FULFILLMENT REPORT FOR " + $bussinessDate + " - " + $endDate
    $savedate = $(Get-Date $csvData.STARTDATE[0]).ToString("M-d-yyyy")

    #$worksheet.Cells["C7:O7"].Merge()
    $worksheet.Cells[7, 3].Value = "DATE (" + $bussinessDate + " - " + $endDate + "); ADD ON ([ALL]); ARRIVAL TYPE ([ALL]); HONORS TIER ([ALL]);"
    $csvData = $csvData | Where-Object { $_.ROOMNUM -ne "" }

    foreach ($row in $csvData) {
        foreach ($property in $row.PSObject.Properties) {
            if ($property.Value -is [string]) {
                $property.Value = $property.Value.Trim()
            }
        }
    }

    $headers = @(
        "GUESTNAME",
        "CONFIRMATIONNUM",
        "HHINFO",
        "ArrivalTime",
        "DCI",
        "DK",
        "ROOMNUM",
        "GST",
        "NUM_OF_NIGHTS",
        "NUM_OF_Addon",
        "ADDON",
        "POSTING_CADENCE",
        "FULLFILMENT_Date",
        "UNIT_PRICE",
        "TOTAL_PRICE",
        "FULFILLMENTS COMMENTS"
    )

    $entryRow = 11
    foreach ($entry in $csvData) {
        $index = 2
        foreach ($header in $headers) {
            $worksheet.Cells[$entryRow, $index].Value = $(if ($header -eq "FULFILLMENTS COMMENTS") { $null }else { if ($entry.$header -eq "-1") { $null }else { $(SafeParseNumberandRound -value $entry.$header) } })
            if ($header -eq "FULFILLMENTS COMMENTS") { $worksheet.Cells[$entryRow, $index].Style.Border.Bottom.Style = [OfficeOpenXml.Style.ExcelBorderStyle]::Thin }
            if ($worksheet.Cells[$entryRow, $index].Style.Border.Top.Style -ne [OfficeOpenXml.Style.ExcelBorderStyle]::None) {
                $worksheet.Cells[$entryRow, $index].Style.Border.Top.Style = $worksheet.Cells[$entryRow, $index].Style.Border.Top.Style
            }
            $index += 1
        }
        $entryRow += 1
    }

    $bottomOfEntries = $entryRow - 1
    $entryRow += 1
    $summaryRow = $entryRow
    $worksheet.Cells[$summaryRow, 2].Value = "SUMMARY"
    $worksheet.Cells[$summaryRow, 2].Style.Font.Bold = $true
    $summaryRow += 1
    $worksheet.Cells[$summaryRow, 2].Value = "ADD-ON"
    $worksheet.Cells[$summaryRow, 2].Style.Font.Bold = $true
    $worksheet.Cells[$summaryRow, 3].Value = "REVENUE"
    $worksheet.Cells[$summaryRow, 3].Style.Font.Bold = $true

    $listRow = $summaryRow + 1
    try { $addons = $csvData.addons | Where-Object { $_ -ne $null -and $_.Trim() -ne "" } | ForEach-Object { $_.Trim() } | Sort-Object | Get-Unique }catch { $addons = $false }

    if ($addons) {
        foreach ($addon in $addons) {
            $amountParsed = 0
            $worksheet.Cells[$listRow, 2].Value = $addon
            $amounts = $csvData | Where-Object { $_.addons -eq $addon }
            $amountParsed = SafeParseNumberandRound -value $amounts[0].amount
            $worksheet.Cells[$listRow, 3].Value = sumList -list $amountParsed
            $totals += $amountParsed
            $listRow += 1
        }
    }

    $worksheet.Cells[$listRow, 2].Value = "TOTAL"
    $worksheet.Cells[$listRow, 2].Style.Font.Bold = $true
    $worksheet.Cells[$listRow, 3].Value = if ($addons) { $totals } else { 0 }

 
    $worksheet.Cells[$entryRow, 2, $entryRow, 3].Style.Border.Top.Style = [OfficeOpenXml.Style.ExcelBorderStyle]::Thin
    $worksheet.Cells[$listRow, 2, $listRow, 3].Style.Border.Bottom.Style = [OfficeOpenXml.Style.ExcelBorderStyle]::Thin
    $worksheet.Cells[$entryRow, 2, $listRow, 2].Style.Border.Left.Style = [OfficeOpenXml.Style.ExcelBorderStyle]::Thin
    $worksheet.Cells[$entryRow, 3, $listRow, 3].Style.Border.Right.Style = [OfficeOpenXml.Style.ExcelBorderStyle]::Thin

    $listRow += 1
    $worksheet.Cells[$listRow, 2].Value = "END OF REPORT"

    $startRow = $worksheet.Dimension.Start.Row
    $endRow = $worksheet.Dimension.End.Row
    for ($row = $startRow; $row -le $endRow; $row++) {
        $worksheet.Row($row).Height = 15
    }

    $startRow = 11
    $endRow = $bottomOfEntries
    for ($row = $startRow; $row -le $endRow; $row++) {
        $worksheet.Row($row).Height = 19.5
    }
  
    $worksheet.Row(10).Height = 18
    $worksheet.Row(1).Height = 32.25

    $filename = "OnQ - Add On Fulfillment"

    # Define the new file name
    $filename = "$($filename) - $inncode - $( $savedate).xlsx"

    $fullOutputPath = Join-Path -Path $outputDir -ChildPath $filename

    # Save the entire Excel file with the new name
    Close-ExcelPackage $excel -SaveAs $fullOutputPath
  
    # Reopen the saved Excel file
    $newExcel = Open-ExcelPackage -Path $fullOutputPath
  
    # Get the list of all worksheet names
    $worksheetNames = $newExcel.Workbook.Worksheets.Name
  
    # Loop through all worksheets and delete those that are not "HK_ALLSUM.RPT"
    foreach ($sheetName in $worksheetNames) {
        if ($sheetName -ne "ADDONFULL.RPT") {
            $newExcel.Workbook.Worksheets.Delete($sheetName)
        }
    }

    Close-ExcelPackage $newExcel -SaveAs $fullOutputPath

    Write-Host "Data has been successfully inserted into the Template worksheet and saved as '$filename'."
}#i think this is good

function RECRCP_GL.RPT {
    param (
        [psobject]$csvData,
        [string]$templatePath,
        [string]$outputDir
    )

    # Load the Excel file
    $excel = Open-ExcelPackage -Path $templatePath

    # Select the "Tester" worksheet
    $worksheet = $excel.Workbook.Worksheets["RECRCP.RPT"]
  
    $bussinessDate = SafeGetDateString -date $csvData.cur_bus_date[0]
  
    $bussinessDate = $($($script:goLiveDate | Get-Date).ToString("M/d/yyyy"))

    $worksheet.Cells[1, 4].Value = Get-Date -Format "M/d/yyyy"
    $worksheet.Cells[2, 4].Value = $bussinessDate
    $worksheet.Cells[3, 4].Value = Get-Date -Format "h:mm tt"
    $worksheet.Cells[3, 6].Value = "FOR " + $bussinessDate

    if ($null -eq $script:propertyIDHeader) { $script:propertyIDHeader = $inncode }

    $worksheet.Cells[1, 6].Value = $script:propertyIDHeader

    $csvData = $csvData | Where-Object { $_.entry_type -ne "" } | Where-Object { $null -ne $_.entry_type }
    $netYesterday = 0
    $netDebits = 0
    $netCredit = 0

    $startRow = 1
    $entryTypes = $csvData.entry_type | Get-Unique
    foreach ($ledger in $entryTypes) {
        $entries = $null
        $entries = $csvData | Where-Object { $_.entry_type -eq "$ledger" }
        $accountingIDDescs = $entries.accounting_id_desc | Get-Unique
        foreach ($accountingDesc in $accountingIDDescs) {
            $entry = $csvData | Where-Object { $_.accounting_id_desc -eq "$accountingDesc" } | Where-Object { $_.entry_type -eq "$ledger" }
            $row, $col = Find-ExcelCellCoordinates -searchString $($accountingDesc.Trim()) -worksheet $worksheet -maxRows 150 -maxCols 10 -startRow $($startRow - 20)
            $worksheet.Cells[$row, 5].Value = if ($entry.Yesterday_balance -eq "0.00") { $null } else { $(SafeParseNumberandRound $entry.Yesterday_balance) }
            $worksheet.Cells[$row, 7].Value = $(SafeParseNumberandRound $entry.Today_debits)
            $worksheet.Cells[$row, 9].Value = $(SafeParseNumberandRound $entry.Today_Credits)
            $startRow = $row
        }
    }

    foreach ($ledger in $entryTypes) {

        $row = switch ($ledger) {
            "1TRAY" { 33 }
            "2CC" { 49 }
            "3DB" { 71 }
            "4YO" { 87 }
            default { $null }
        }
        if ($null -ne $row) {
            $entries = $null
            $entries = $csvData | Where-Object { $_.entry_type -eq "$ledger" }
            $netEntryYesterday = sumList -list $entries.Yesterday_balance
            $netEntryDebits = sumList -list $entries.Today_debits
            $netEntryCredits = sumList -list $entries.Today_Credits
            $netYesterday += $netEntryYesterday
            $netDebits += $netEntryDebits
            $netCredit += $netEntryCredits
            $worksheet.Cells[$row, 5].Value = $netEntryYesterday
            $worksheet.Cells[$row, 7].Value = $netEntryDebits
            $worksheet.Cells[$row, 9].Value = $netEntryCredits
            $worksheet.Cells[$row, 11].Value = ($netEntryYesterday + $netEntryDebits - $netEntryCredits)
        }
    }

    $row = 88
    $worksheet.Cells[$row, 5].Value = $netYesterday
    $worksheet.Cells[$row, 7].Value = $netDebits
    $worksheet.Cells[$row, 9].Value = $netCredits
    $worksheet.Cells[$row, 11].Value = ($netYesterday + $netDebits - $netCredits)

    $todaysTotalOutstanding = 0

    foreach ($ledger in $($csvData | Where-Object { $_.entry_type -eq "5NET" })) {
        if ($ledger.accounting_id_desc -ne "+ Yesterdays Outstanding") {
            switch ( $ledger.accounting_id_desc.Substring(0, 1)) {
                "+" { $todaysTotalOutstanding += $(SafeParseNumberandRound $ledger.Yesterday_balance) }
                "-" { $todaysTotalOutstanding -= $(SafeParseNumberandRound $ledger.Yesterday_balance) }
            }
        }
        else {
            $todaysTotalOutstanding += $netYesterday
            $worksheet.Cells[103, 5].Value = $netYesterday
        }
    }

    $worksheet.Cells[108, 5].Value = $todaysTotalOutstanding
  
    $filename = "OnQ - Detailed Receivables Recap"

    # Define the new file name
    $filename = "$($filename) - $($csvData.property_id[0]) - $($($script:goLiveDate | Get-Date).ToString("M-d-yyyy")).xlsx"

    $fullOutputPath = Join-Path -Path $outputDir -ChildPath $filename

    # Save the entire Excel file with the new name
    Close-ExcelPackage $excel -SaveAs $fullOutputPath
  
    # Reopen the saved Excel file
    $newExcel = Open-ExcelPackage -Path $fullOutputPath
  
    # Get the list of all worksheet names
    $worksheetNames = $newExcel.Workbook.Worksheets.Name
  
    # Loop through all worksheets and delete those that are not "HK_ALLSUM.RPT"
    foreach ($sheetName in $worksheetNames) {
        if ($sheetName -ne "RECRCP.RPT") {
            $newExcel.Workbook.Worksheets.Delete($sheetName)
        }
    }

    Close-ExcelPackage $newExcel -SaveAs $fullOutputPath

    Write-Host "Data has been successfully inserted into the Template worksheet and saved as '$filename'."
} #i think this is good

function HABALSUM.RPT {
    param (
        [psobject]$csvData,
        [string]$templatePath,
        [string]$outputDir
    )


    # Load the Excel file
    $excel = Open-ExcelPackage -Path $templatePath

    # Select the "Tester" worksheet
    $worksheet = $excel.Workbook.Worksheets["HABALSUM.RPT"]

    $bussinessDate = SafeGetDateString -date $csvData.cur_bus_date[0]
  
    $worksheet.Cells[1, 4].Value = Get-Date -Format "M/d/yyyy"
    $worksheet.Cells[2, 4].Value = $bussinessDate
    $worksheet.Cells[3, 4].Value = Get-Date -Format "h:mm tt"

    $worksheet.Cells[1, 6].Value = $csvData.property_id[0] + " - " + $csvData.property_name[0] + " - " + $csvData.facility_id[0]

    $csvData = $csvData | Where-Object { $_.house_acct_id -ne "" } | Where-Object { $null -ne $_.house_acct_id } | Sort-Object -Property { [int]$_.house_acct_id }

    $entriesCount = $csvData.Count
    $startRow = 6
    $startCol = 1
    $headers = @(
        "house_acct_id",
        "blank1",
        "account_name",
        "blank2",
        "blank3",
        "blank4",
        "balance",
        "folio_id",
        "last_posted_date"
    )

    foreach ($header in $headers) {
        for ($i = 0; $i -lt $entriesCount; $i++) {
            $value = $null
            switch ($header) {
                "blank1" { $value = $null }
                "blank2" { $value = $null }
                "blank3" { $value = $null }
                "blank4" { $value = $null }
                "last_posted_date" { $value = SafeGetDateString -date $csvData.$header[$i] }
                "account_name" { $value = $csvData.$header[$i] }
                "house_acct_id" { $value = [int]$csvData.$header[$i] }
                "folio_id" { $value = [int]$csvData.$header[$i] }
                "balance" { $value = SafeParseNumberandRound -value $csvData.$header[$i] }
                default { $value = SafeParseNumberandRound -value $csvData.$header[$i] }
            } 
  
            $spaceRows = 3
            if ($i -eq 0) {
                $previousEntryRow = $startRow
            }
            $worksheet.Cells[($previousEntryRow + $spaceRows), $startCol].Value = $value
            $previousEntryRow = $previousEntryRow + $spaceRows
        }
        $startCol++
    }


    $worksheet.Cells[($previousEntryRow + 4), 1].Value = "   TOTAL  HOUSE ACCT BALANCES :"
    $worksheet.Cells[($previousEntryRow + 4), 5].Value = sumList -list $csvData.balance
    $worksheet.Cells[($previousEntryRow + 4), 6].Value = "(EXCLUDES ADV DEP)"
    $worksheet.Cells[($previousEntryRow + 6), 1].Value = "  *HOUSE ACCT WITH ADV. DEP :"
    $worksheet.Cells[($previousEntryRow + 6), 5].Value = sumList -list $csvData.AdvanceDepositTotal
    $worksheet.Cells[($previousEntryRow + 6), 6].Value = "(ALSO INCLUDED IN ADV DEP TRAY)"
    $worksheet.Cells[($previousEntryRow + 8), 1].Value = "       TOTAL HOUSE ACCT TRAY :"
    $worksheet.Cells[($previousEntryRow + 8), 5].Value = ($(sumList -list $csvData.balance) + $(sumList -list $csvData.AdvanceDepositTotal))
    $worksheet.Cells[($previousEntryRow + 13), 1].Value = "END OF REPORT"

    $startRow = $worksheet.Dimension.Start.Row
    $endRow = $worksheet.Dimension.End.Row
  
    for ($row = $startRow; $row -le $endRow; $row++) {
        $worksheet.Row($row).Height = 12.75
    }
  
    $filename = "OnQ - House Account Balance Summary Report"

    # Define the new file name
    $filename = "$($filename) - $($csvData.property_id[0]) - $($(Get-Date $csvData.cur_bus_date[0]).ToString("M-d-yyyy")).xlsx"

    $fullOutputPath = Join-Path -Path $outputDir -ChildPath $filename

    # Save the entire Excel file with the new name
    Close-ExcelPackage $excel -SaveAs $fullOutputPath
  
    # Reopen the saved Excel file
    $newExcel = Open-ExcelPackage -Path $fullOutputPath
  
    # Get the list of all worksheet names
    $worksheetNames = $newExcel.Workbook.Worksheets.Name
  
    # Loop through all worksheets and delete those that are not "HK_ALLSUM.RPT"
    foreach ($sheetName in $worksheetNames) {
        if ($sheetName -ne "HABALSUM.RPT") {
            $newExcel.Workbook.Worksheets.Delete($sheetName)
        }
    }

    Close-ExcelPackage $newExcel -SaveAs $fullOutputPath

    Write-Host "Data has been successfully inserted into the Template worksheet and saved as '$filename'."
} #i think this is good

function viewTotals.RPT {
    param (
        [psobject]$csvData,
        [string]$templatePath,
        [string]$outputDir
    )

    # Load the Excel file
    $excel = Open-ExcelPackage -Path $templatePath

    # Select the "Tester" worksheet
    $worksheet = $excel.Workbook.Worksheets["OnQ View Totals Extract"]

    $csvData = $csvData | Where-Object { $null -ne $_.crs_room_type }

    $roomTypes = $csvData.crs_room_type | Get-Unique

    $i = 2
    foreach ($roomType in $roomTypes) {
        $worksheet.Cells[$i, 1].Value = $roomType
        $totals = $csvData | Where-Object { $roomType -eq $_.crs_room_type }
        foreach ($total in $totals) {
            $row, $col = Find-ExcelCellCoordinates -worksheet $worksheet -searchString $total.view_totals_status -maxRows 2 -maxCols 15 -startRow 1
            $worksheet.Cells[$i, $col].Value = [int]$total.stay_count
            $worksheet.Cells[$i, 14].Value = $goLiveDate
            $worksheet.Cells[$i, 15].Value = Get-Date -Format 'M/d/yyyy HH:mm:ss'
            $worksheet.Cells[$i, 16].Value = $inncode
        }
        $i++
    }

    $filename = "OnQ - View Totals"

    # Define the new file name
    $filename = "$($filename) - $($inncode) - $($(Get-Date $goLiveDate).ToString("M-d-yyyy")).xlsx"

    $fullOutputPath = Join-Path -Path $outputDir -ChildPath $filename

    # Save the entire Excel file with the new name
    Close-ExcelPackage $excel -SaveAs $fullOutputPath
  
    # Reopen the saved Excel file
    $newExcel = Open-ExcelPackage -Path $fullOutputPath
  
    # Get the list of all worksheet names
    $worksheetNames = $newExcel.Workbook.Worksheets.Name
  
    # Loop through all worksheets and delete those that are not "HK_ALLSUM.RPT"
    foreach ($sheetName in $worksheetNames) {
        if ($sheetName -ne "View Totals") {
            if ($sheetName -ne "OnQ View Totals Extract") {
                $newExcel.Workbook.Worksheets.Delete($sheetName)
            }
        }
    }

    Close-ExcelPackage $newExcel -SaveAs $fullOutputPath

    Write-Host "Data has been successfully inserted into the Template worksheet and saved as '$filename'."


}

######################################################

##########################
##     Start Script     ##
##########################


$databaseName = "HPMS3"
$query = @"
USE HPMS3;

SELECT 
    p.property_id,
    p.property_name,
    p.cur_bus_date,
    a.audit_marker_code,
    a.description,
    a.audit_procedure,
    a.date_completed,
    a.start_time,
    a.end_time,
    a.real_end_time
FROM dbo.PROPERTY p
CROSS JOIN dbo.AUDIT_MARKER a
WHERE a.audit_marker_code = 'DTBMP';
"@

# Run the SQL command
$sqlresult = Invoke-Sqlcmd -ServerInstance $env:COMPUTERNAME -Database $databaseName -Query $query
$goLiveDate = ($sqlresult.cur_bus_date | Get-Date).ToString("yyyy-MM-dd")


$inncode = $env:COMPUTERNAME -replace "SERVER", ""
$script:propertyIDHeader = $null

$OutputFolderPath = $saveDirectory
Ensure-PathExists -path $OutputFolderPath

Write-Host 'Export functions are running. Please wait...' -BackgroundColor DarkRed -ForegroundColor White # or bg magenta
$reports = setSqlVariablesFromInncode -inncode $inncode -goLiveDate $goLiveDate
$keys = $reports.PSObject.Properties.Name

foreach ($key in $keys) {

    $response = $null
    $functionName = "$key.RPT"
    $query = $($reports.$key)
    Write-Host "`n"
    Write-Host "Running report $functionName" -BackgroundColor DarkRed -ForegroundColor White

    $response = invoke-sqlcmd -ServerInstance $env:COMPUTERNAME -database "hpms3" -query $query -QueryTimeout 180
    $response | Format-Table
    $filteredResponse = $response | Where-Object { $_.entry_id -ne "" }
    $trimmedResponse = $filteredResponse | ForEach-Object {
        $_.PSObject.Properties | ForEach-Object {
            if ($_.Value -is [string]) {
                $_.Value = $_.Value.Trim()
            }
        }
        $_
    }
    & $functionName -csvData $trimmedResponse -templatePath $templateFilePath -outputDir $OutputFolderPath      
}

Write-host "Reports saved to $OutputFolderPath"-ForegroundColor White
Write-Host "`n`n"
Write-Host "Script Completed" -ForegroundColor White -BackgroundColor DarkGreen
exit
