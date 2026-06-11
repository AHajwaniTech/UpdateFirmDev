<?php
/**
 * ============================================================================
 * FIRM COST CHECK DETAIL REPORT GENERATOR
 * ============================================================================
 * 
 * @author      KEANT Technologies
 * @description Generates Excel reports for Firm cost transactions from RMAACABHS
 *              database. Processes remittance transactions (RMSTRANCDE='1A')
 *              and creates detailed financial breakdowns per client code.
 * 
 * @version     1.0
 * @created     2026-01-26
 * 
 * WORKFLOW:
 * 1. Accepts comma-separated client codes and optional date parameter
 * 2. Loops through each client code (VENDORNUM)
 * 3. Queries distinct PYALORGCD for each firm
 * 4. Retrieves transaction details: account info, payments, remittances, costs
 * 5. Generates Excel worksheet with:
 *    - Client account details (names, account numbers)
 *    - Transaction breakdown (dates, codes, descriptions)
 *    - Financial data (payment amounts, costs requested/paid, amounts paid to firm)
 *    - Subtotals per client code
 *    - Grand totals
 * 6. Saves Excel files to client-specific directories
 * 7. Sends email notifications (if enabled), but we have commented it out for now.
 * 8. Returns status array with success/failure results
 * 
 * CHANGELOG:
 * ----------------------------------------------------------------------------
 * Version | Date       | Author              | Description
 * ----------------------------------------------------------------------------
 * 1.0     | 2026-01-26 | KEANT Technologies  | Initial version with developer
 *         |            |                     | documentation and parameter-based
 *         |            |                     | date override functionality
 * 1.1     |2026-05-14  | AH15042026  AH        | Added  Validation for stop 
 *												producing zero data files
 * ----------------------------------------------------------------------------
 * 1.2     | 2026-05-19 | LK20052026          | Added AS400 ODBC connection            
 *         |            |                     | replaced PDO $conndb2 query with                    
 *         |            |                     | ODBC AS400 direct connection    
 *-----------------------------------------------------------------------------
 */
require_once('PHP_XLSXWriter/xlsxwriter.class.php');
// require_once(__DIR__ . '/Report/PHP_XLSXWriter/xlsxwriter.class.php');
// use XLSXWriter;

// LK20052026 - start
// Define log file path
if (!defined('FEE_LOG_FILE')){define('FEE_LOG_FILE', '/home/pipewayweb/log/W_FC_FF_Logs_POC');}

/**
 * Write log message to file with timestamp
 * @param string $message Log message
 * @param string $level Log level (INFO, ERROR, SUCCESS)
 */
function writeLogcost($message, $level = 'INFO') {
    // AH08052026 START
    // WRAPPER_LOG_ONLY flag — when defined as true in the wrapper, suppress all
    // INFO-level log lines from child scripts. Only ERROR level and explicit wrapper
    // START/END calls (which use writeLogcost directly in wrapper scope) pass through.
    // Set WRAPPER_LOG_ONLY = false in wrapper to restore full verbose logging.
    if(defined('WRAPPER_LOG_ONLY') && WRAPPER_LOG_ONLY === true && $level === 'INFO' && strpos($message, '[WRAPPER]') === false) {
        return;
    }
    // AH08052026 END

    $timestamp = date('Y-m-d H:i:s');
    $logMessage = "[{$timestamp}] [{$level}] {$message}" . PHP_EOL;

    $result = file_put_contents(FEE_LOG_FILE, $logMessage, FILE_APPEND);

    if ($result === false) {
        echo "LOG WRITE FAILED\n";
        if(function_exists('error_get_last')) {
            print_r(error_get_last());
        }
    }
}

// AS400 ODBC connection constants — update DSN/user/pass as per /etc/odbc.ini
// define('AS400_DSN',     'DB2');              // DSN name configured in /etc/odbc.ini
// define('AS400_HOST',    '192.168.21.202');     // AS400 hostname or IP
// // define('AS400_LIB',     'SABINESH','AACALIB');         // Default library/schema
// // FIX 1: Library list ko properly array format me define kiya taaki syntax error na aaye
// define('AS400_LIBS',    ['AACALIB', 'SABINESH']); 
// define('AS400_USER',    'AHAJWANI');        // AS400 user profile
// define('AS400_PASS',    'Ahajwani@1210');       // AS400 password



if (!defined('AS400_DSN'))  { define('AS400_DSN',  'DB2'); }
if (!defined('AS400_HOST')) { define('AS400_HOST', '192.168.21.8'); }
if (!defined('AS400_USER')) { define('AS400_USER', 'TESTADMIN'); }
if (!defined('AS400_PASS')) { define('AS400_PASS', 'x1ns8@p5d7'); }

/**
 * Connect to AS400 via ODBC
 * @return resource ODBC connection resource
 * @throws Exception if connection fails
 */
function connectAS400cost() {
    $conn = @odbc_connect(AS400_DSN, AS400_USER, AS400_PASS);
    if (!$conn) {
        $err = odbc_errormsg();
        writeLogcost('AS400 ODBC connection failed: ' . $err, 'ERROR');
        throw new Exception('AS400 ODBC connection failed: ' . $err);
    }
    return $conn;
}

/**
 * Execute a query on AS400 via ODBC and return results
 * mimics PDO fetchAll(PDO::FETCH_OBJ) — returns array of stdClass objects
 * so all existing $row->COLUMN references work without any change
 * @param resource $conn  ODBC connection resource
 * @param string   $query SQL query string (no parameters needed — billdate injected as string)
 * @return array          Array of stdClass objects (same as PDO FETCH_OBJ)
 */
function odbc_fetch_all_objcost($conn, $query) {
    $stmt = @odbc_exec($conn, $query);
    if (!$stmt) {
        $err = odbc_errormsg($conn);
        writeLogcost('AS400 ODBC query failed: ' . $err . ' | Query: ' . $query, 'ERROR');
          if (is_resource($stmt)) {
        odbc_free_result($stmt);
    				}
        return [];
    }
    $results = [];
    while ($row = odbc_fetch_array($stmt)) {
        // Trim trailing spaces — AS400 pads CHAR fields with spaces
        $cleanRow = [];

		foreach ($row as $col => $val) {
			$cleanRow[$col] = is_string($val) ? trim($val) : $val;
		}

		$results[] = $cleanRow;
    }
    odbc_free_result($stmt);
    return $results;
}
// LK20052026 - end

function firmCostCheckDetailToMyDownload($path, $id, $reportName, $code_name, $userType, $userReportName, $outputName, $reportDescription, $mailNotification, $sftpId, $mode, $run_by, $reportBasePath, $reportDate = null)
{

	//AH03062026 :Start - Add company tracking arrays to match FirmFee
	$generatedCompanies = array();
	$existingCompanies  = array();
	$failedCompanies    = array();
	//AH03062026 : End

	$report_start_time = date("H:i:s");
	$date = new DateTime();
	//AH15052026 : CHnaged File Name 
	//$fileName = filterReportName($date->format('Y-m-d'), $reportName);

	$paths = explode(",", $path);
	$dataPresents = array();
	$noDataPresents = array();
	$new_status = 3;
	$msg = 'No Data Present';
	$status_msg = 'Failed';
	$sftpStatus = 0;
	$status = 0;
	$FileSizeKB = 0;

	// ========================================================================
	// DATE PARAMETER HANDLING
	// ========================================================================
	
	/**
	 * Check if reportDate parameter is provided
	 * If provided, use it; otherwise auto-calculate based on current day
	 */
	if ($reportDate !== null && !empty($reportDate)) {
		// User provided a specific report date
		$queryDate = $reportDate;
	} else {
		// Auto-calculate date: Monday = -3 days, Other days = -1 day
		$currentDay = date('D');

		//AH01062026 : Start
		// if($currentDay == 'Mon'){
		// 	// $queryDate = date('Ymd', strtotime('-3 day'));
		// 	$queryDate = date('Y-m-d', strtotime('-3 day'));
		// } else {
		// 	// $queryDate = date('Ymd', strtotime('-1 day'));
		// 	$queryDate = date('Y-m-d', strtotime('-1 day'));
		// }
		$queryDate = date('Y-m-d');
		//AH01062026 : End
	}


	foreach ($paths as $companyPath) {
		$companyPath = str_replace(["'", " "], "", $companyPath);
		$RemitAmount = '0';
		$PaymentAmount = '0';
		$FeePaidToFirm = '0';
		$FeeRequestedByFirm = '0';
		$AmountPaidToFirm = '0';
		// if ($results['numRows'] > 0)
		$companyName = createDirgetCompanyName($companyPath, $reportBasePath);
		$fileName = $companyName . '_' . $reportName . '_' . $queryDate . '.xlsx'; //AH15052026 : CHnaged File Name
		echo "Processing Company: " . $companyName . "\n";
		writeLogcost("Processing Company: " . $companyName); //AH02062026
		$companyStatus = getCompanyStatus($companyName);

		// ====================================================================
		// QUERY 1: GET DISTINCT CLIENT CODES (PYALORGCD) FOR THIS FIRM
		// ====================================================================
		
		/**
		 * ORIGINAL QUERY (AUTO-CALCULATED DATE) - COMMENTED OUT
		 * This query automatically calculates the date based on current day
		 * Monday: -3 days, Other days: -1 day
		 */
		/*
		$query = "SELECT DISTINCT PYALORGCD
			from RMAACABHS
			WHERE RMSTRANCDE='1A'
			AND CAST(DTPRLT AS DATE) >= (case when weekday(CURRENT_DATE()) = 0 then DATE_SUB(CURRENT_DATE(),INTERVAL 3 DAY) else DATE_SUB(CURRENT_DATE(), INTERVAL 1 DAY) end)
			AND VENDORNUM IN ('" . $companyName . "') group by PYALORGCD ORDER BY  PYALORGCD";
		*/

		/**
		 * NEW QUERY (PARAMETER-BASED DATE)
		 * Uses $queryDate variable which can be:
		 * - Provided via $reportDate parameter, OR
		 * - Auto-calculated based on current day
		 */
		 // LK20052026 - start
    	// Query to get PYALORGCD — replaced PDO $conndb2 with AS400 ODBC connection
		

		// $query = "SELECT DISTINCT PYALORGCD
		// 	from RMAACABHS
		// 	WHERE RMSTRANCDE='1A'
		// 	-- AND CAST(DTPRLT AS DATE) >= '" . $queryDate . "'
		// 	AND VENDORNUM IN ('" . $companyName . "') group by PYALORGCD ORDER BY  PYALORGCD";

		//Updated the query from MySQL to ODBC format and modified the syntax as per ODBC/AS400 compatibility. LK20052026

		$query = "SELECT DISTINCT PYALORGCD
            FROM SABINESH.RMAACABLF
            WHERE RMSTRANCDE='1A'
            AND DTPRLT >= '" . $queryDate . "'
            AND VENDORNUM IN ('" . $companyName . "') 
            GROUP BY PYALORGCD 
            ORDER BY PYALORGCD";

           try {
					$as400conn = connectAS400cost();
					$queryResult = odbc_fetch_all_objcost($as400conn, $query);
					odbc_close($as400conn);
					$results = array('numRows' => count($queryResult),'results' => $queryResult);
					writeLogcost('AS400 client query returned ' .count($queryResult) .' records.','INFO'); // AH05062026 - fixed: writeLog() undefined, correct is writeLogcost()
				} catch (Exception $e) {
					writeLogcost('Failed to fetch clients from AS400: ' .$e->getMessage(),'ERROR'); // AH05062026 - fixed: writeLog() undefined, correct is writeLogcost()
					$results = array('numRows' => 0,'results' => array());
				}
				
		// $results = getResult($query); LK05052026 end

		if ($results['numRows'] > 0) {
			$writer = new XLSXWriter();
			$excelPrefix = getExcelPrefix13();
			$writer->writeSheetHeader($excelPrefix['sheetName'], $excelPrefix['headers'], $excelPrefix['style']);
			$blankrow = array('', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '');
			$writer->writeSheetRow($excelPrefix['sheetName'], $blankrow);
			$hasData = false;
			foreach ($results['results'] as $result) {
				// ============================================================
				// QUERY 2: GET DETAILED TRANSACTION DATA FOR EACH CLIENT CODE
				// ============================================================
				
				/**
				 * This query retrieves all transaction details for a specific
				 * client code (PYALORGCD) including account info, transaction
				 * dates, amounts, fees, and firm file numbers
				 * 
				 * Note: This query also uses $queryDate variable for consistency
				 */
                // LK20052026 - start
				// $query = "WITH FIRMCOSTFEE AS (
				// 	select BLAAINNM, LEFT(BLAAINNM, 1) as TYPERT, CUROFFCRCD AS 'Client_Code', VENDORNUM AS 'Firm',
				// 	RMSACCTNUM AS 'Acct_Number',RMSCORPNM1 AS 'Last_Name', RMSCORPNM2 AS 'First_Name',
				// 	RMSTRANCDE AS 'TR_CD', RMSTRANDTE AS 'Transaction_Date', RMSTRANDSC AS 'Transaction_Description',
				// 	IF (RMSTRANCDE = '1A' , INVCLIENT , 0.00) AS 'Cost_Amount', COLLAM AS 'Payment_Amount', DUECLIENT AS 'Remit_Amount',
				// 	FEESFR AS 'Fee_Requested_by_Firm', FEES AS 'Fee_Paid_to_Firm', INVOICENO AS 'Firm_Invoice_No',
				// 	CHECKNO AS 'Firm_Check_Number', BLPYFRCK AS 'AACA_Check_Number',BLPYFRTO AS 'Amount_Paid_to_Firm',
				// 	DTPRLT AS 'AACA_Check_Date', RMSFILENUM, PYALORGCD, PAIDDATE  AS 'Firm_Invoice_Date'
				// 	from RMAACABHS
				// 	WHERE RMSTRANCDE='1A'
				// 	AND CAST(DTPRLT AS DATE)  >= '" . $queryDate . "'
				// 	AND VENDORNUM ='" . $companyName . "' AND PYALORGCD='" . $result['PYALORGCD'] . "')SELECT Client_Code as 'Client Code', Acct_Number as 'Acct No.', Last_Name as 'Last Name ', First_Name as 'First Name',
				// 	TR_CD as 'TR CD', Transaction_Date as 'Transaction Date', Transaction_Description as 'Transaction Description',
				// 	Payment_Amount AS 'Payment Amount' ,Remit_Amount AS 'Remit Amount',
				// 	Fee_Requested_by_Firm AS 'Fee Requested By Firm',
				// 	Fee_Paid_to_Firm as 'Fee Paid To Firm',Firm_Invoice_No as 'Firm Invoice No', Firm_Check_Number as 'Firm Check Number',
				// 	Amount_Paid_to_Firm as 'Amount Paid To Firm' ,
				// 	Firm_Invoice_Date as 'Firm Invoice Date', P.FILELOCATN AS 'Firm File No',PYALORGCD
				// 	from FIRMCOSTFEE
				// 	LEFT JOIN (
				// 	SELECT DISTINCT RMSFILENUM, FILELOCATN
				// 	FROM RMSPMASTER
				// 	) AS P
				// 	ON FIRMCOSTFEE.RMSFILENUM = P.RMSFILENUM
				// 	ORDER BY  PYALORGCD,Client_Code,Firm_Invoice_Date, Firm_Invoice_No";

		//Updated the query from MySQL to ODBC format and modified the syntax as per ODBC/AS400 compatibility. LK20052026
			writeLogfee('Company: ' . $companyName .' | PYALORGCD: ' . $result['PYALORGCD'], 'INFO');
			$query = "WITH FIRMCOSTFEE AS (
                    SELECT 
                        BLAAINNM,
                        SUBSTR(BLAAINNM,1,1) AS TYPERT,
                        CUROFFCRCD AS Client_Code,
                        VENDORNUM AS Firm,
                        RMSACCTNUM AS Acct_Number,
                        RMSCORPNM1 AS Last_Name,
                        RMSCORPNM2 AS First_Name,
                        RMSTRANCDE AS TR_CD,
                        RMSTRANDTE AS Transaction_Date,
                        RMSTRANDSC AS Transaction_Description,
                        CASE 
                            WHEN RMSTRANCDE = '1A' THEN INVCLIENT
                            ELSE 0.00
                        END AS Cost_Amount,
                        COLLAM AS Payment_Amount,
                        DUECLIENT AS Remit_Amount,
                        FEESFR AS Fee_Requested_by_Firm,
                        FEES AS Fee_Paid_to_Firm,
                        INVOICENO AS Firm_Invoice_No,
                        CHECKNO AS Firm_Check_Number,
                        BLPYFRCK AS AACA_Check_Number,
                        BLPYFRTO AS Amount_Paid_to_Firm,
                        DTPRLT AS AACA_Check_Date,
                        RMSFILENUM,
                        PYALORGCD,
                        PAIDDATE AS Firm_Invoice_Date
                    FROM SABINESH.RMAACABLF
                    WHERE RMSTRANCDE = '1A'
                    AND DTPRLT >= '" . $queryDate . "'
                    AND VENDORNUM = '" . $companyName . "'
                    AND PYALORGCD = '" . $result['PYALORGCD'] . "'
                )
                SELECT
                    Client_Code AS \"Client Code\",
                    Acct_Number AS \"Acct No.\",
                    Last_Name AS \"Last Name\",
                    First_Name AS \"First Name\",
                    TR_CD AS \"TR CD\",
                    Transaction_Date AS \"Transaction Date\",
                    Transaction_Description AS \"Transaction Description\",
                    Payment_Amount AS \"Payment Amount\",
                    Remit_Amount AS \"Remit Amount\",
                    Fee_Requested_by_Firm AS \"Fee Requested By Firm\",
                    Fee_Paid_to_Firm AS \"Fee Paid To Firm\",
                    Firm_Invoice_No AS \"Firm Invoice No\",
                    Firm_Check_Number AS \"Firm Check Number\",
                    Amount_Paid_to_Firm AS \"Amount Paid To Firm\",
                    Firm_Invoice_Date AS \"Firm Invoice Date\",
                    P.FILELOCATN AS \"Firm File No\",
                    PYALORGCD
                FROM FIRMCOSTFEE
                LEFT JOIN (
                    SELECT DISTINCT RMSFILENUM, FILELOCATN
                    FROM RMSAACAOBJ.RMSPMASTER
                ) AS P ON FIRMCOSTFEE.RMSFILENUM = P.RMSFILENUM
                ORDER BY PYALORGCD, Client_Code, Firm_Invoice_Date, Firm_Invoice_No";
				// $results = getResult($query); LK05052026
				
    			// — replaced PDO $conndb2 with AS400 ODBC connection
				try {
						$as400conn = connectAS400cost();
						$queryResult = odbc_fetch_all_objcost($as400conn, $query);
						odbc_close($as400conn);
						$results = array('numRows' => count($queryResult),'results' => $queryResult);
						writeLogcost('Company: ' . $companyName . ' | Client Query Returned: ' . count($queryResult) . ' records','INFO');
					} catch (Exception $e) {
						writeLogcost('Failed to fetch data from AS400: ' .$e->getMessage(),'ERROR');
						$results = array('numRows' => 0,'results' => array());
					}
					// LK20052026 - end

				if ($results['numRows'] > 0) {
					$hasData = true;

					$PaymentAmountwise = '0';
					$RemitAmountwise = '0';
					$FeePaidToFirmwise = '0';
					$FeeRequestedByFirmwise = '0';
					$AmountPaidToFirmwise = '0';
					foreach ($results['results'] as $resultRow) {

						$PaymentAmount += $resultRow['Payment Amount'];
						$PaymentAmountwise += $resultRow['Payment Amount'];

						$RemitAmount += $resultRow['Remit Amount'];
						$RemitAmountwise += $resultRow['Remit Amount'];

						$FeeRequestedByFirm += $resultRow['Fee Requested By Firm'];
						$FeeRequestedByFirmwise += $resultRow['Fee Requested By Firm'];

						$FeePaidToFirm += $resultRow['Fee Paid To Firm'];
						$FeePaidToFirmwise += $resultRow['Fee Paid To Firm'];

						$AmountPaidToFirm += $resultRow['Amount Paid To Firm'];
						$AmountPaidToFirmwise += $resultRow['Amount Paid To Firm'];
						$writer->writeSheetRow($excelPrefix['sheetName'], $resultRow);
					}
					$newarray1 = array();
					$newarray1 = [$result['PYALORGCD'], '', '', '', '', '', '', $PaymentAmountwise, $RemitAmountwise, $FeeRequestedByFirmwise, $FeePaidToFirmwise, '', '', $AmountPaidToFirmwise, '', '', ''];
					$writer->writeSheetRow($excelPrefix['sheetName'], $newarray1, $excelPrefix['style1']);
					$blankrow = array('', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '');
					$writer->writeSheetRow($excelPrefix['sheetName'], $blankrow);
				} else {
					writeLogcost('No data present for company: ' . $companyName,'INFO');

					// ifDataNotPresent($companyStatus, $companyName, $reportName, $report_start_time, $sftpStatus, $FileSizeKB, $path, $run_by);
					// array_push($DataPresents, array(
					// 	'paths' => $path,
					// 	'filename' => $fileName,
					// 	'clientcode' => $companyName
					// ));
				}
			}
			$newarray = array();
			$newarray = ['TOTAL', '', '', '', '', '', '', $PaymentAmount, $RemitAmount, $FeeRequestedByFirm, $FeePaidToFirm, '', '', $AmountPaidToFirm, '', '', ''];
			
		//AH13052026: Start			
		// 	if ($hasData && floatval($AmountPaidToFirm) > 0) {
		// 	$writer->writeSheetRow($excelPrefix['sheetName'], $newarray, $excelPrefix['style1']);
		// 	$writer->writeToFile(str_replace(__FILE__, $reportBasePath . $companyPath . '/' . $fileName, __FILE__));
		// 	if (file_exists(str_replace(__FILE__, $reportBasePath . $companyPath . '/' . $fileName, __FILE__))) {
		// 		$new_status = 2;
		// 		// $$FileSizeKB = 0;
		// 		$FileSizeKB = getfileSize($reportBasePath . $companyPath . '/' . $fileName);
		// 		// VK26JAN2026 mailNotifaction($mailNotification, $companyPath, $companyName, $userType, $userReportName, $reportDescription);


		// 		array_push($dataPresents, array(
		// 			'paths' => $companyPath,
		// 			'filename' => $fileName,
		// 			'clientcode' => $companyName
		// 		));
		// 		// VK26JAN2026 ifDataPresent($companyStatus, $companyName, $reportName, $report_start_time, $sftpStatus, $FileSizeKB, $companyPath, $run_by,$userType);
		// 	echo "Cost report generated successfully for company: " . $companyName . "\n"; // VK26JAN2026	
		// 	}
		// 	} // end if ($hasData)
		// } else {
		// 	// VK26JAN2026 ifDataNotPresent($companyStatus, $companyName, $reportName, $report_start_time, $sftpStatus, $FileSizeKB, $companyPath, $run_by,$userType);
		// 	echo "cost : No data present for company: " . $companyName . "\n"; // VK26JAN2026
		$allZero =
		floatval($PaymentAmount) == 0 &&
		floatval($RemitAmount) == 0 &&
		floatval($FeeRequestedByFirm) == 0 &&
		floatval($FeePaidToFirm) == 0 &&
		floatval($AmountPaidToFirm) == 0;

		// Generate XLS only when:
		// 1. Detail rows exist
		// 2. At least one financial total is non-zero
		if ($hasData && !$allZero) {
			$writer->writeSheetRow($excelPrefix['sheetName'], $newarray, $excelPrefix['style1']);
			//AH15052026: Start - Check if file already exists before writing
			$filePath = str_replace(__FILE__,$reportBasePath . $companyPath . '/' . $fileName,__FILE__);
			if (file_exists($filePath)) {
				echo "Report already exists for company: " . $companyName . "\n";
				$existingCompanies[] = $companyName; //AH03062026
				} else //AH15052026: End - If file doesn't exist, write it
				{
			$writer->writeToFile(
				str_replace(__FILE__, $reportBasePath . $companyPath . '/' . $fileName, __FILE__)
			);
			if (file_exists(
				str_replace(__FILE__, $reportBasePath . $companyPath . '/' . $fileName, __FILE__)
			)) {
				$new_status = 2;
				$FileSizeKB = getfileSize(
					$reportBasePath . $companyPath . '/' . $fileName
				);
				array_push($dataPresents, array(
					'paths' => $companyPath,
					'filename' => $fileName,
					'clientcode' => $companyName
				));
			echo "Cost report generated successfully for company: " . $companyName . "\n";
			$generatedCompanies[] = $companyName; //AH03062026
					}
				}
			} else {
			echo "cost : No data present for company: " . $companyName . "\n";
		
			array_push($noDataPresents, array(
				'paths' => $companyPath,
				'filename' => $fileName,
				'clientcode' => $companyName
			));
		}
		//AH13052026: End
		if (!empty($dataPresents)) {
			$new_status = 2;
			$msg = 'Report generated successfully';
			$status_msg = 'generated';
			$status = 1;
		} else {
			$new_status = 3;
			$msg = 'No Data Present';
			$status_msg = 'Failed';
			$status = 0;
		}
	}
	}
	writeLogcost('Firm Cost Report Completed','INFO');
	return array(
    'status' => $status,
    'status_msg' => $status_msg,
    'new_status' => $new_status,

    'generated_companies' => $generatedCompanies,
    'existing_companies'  => $existingCompanies,
    'failed_companies'    => $failedCompanies
);
}

function getExcelPrefix13()
{

	$header = array(
		'Client Code' => 'string',
		'Acct No.' => 'string',
		'Last Name' => 'string',
		'First Name' => 'string',
		'TR CD' => 'string',
		'Transaction Date' => 'string',
		'Transaction Description' => 'string',
		'Payment Amount' => 'dollar',
		'Remit Amount' => 'dollar',
		'Fee Requested By Firm' => 'dollar',
		'Fee Paid To Firm' => 'dollar',
		'Firm Invoice No' => 'string',
		'Firm Check Number' => 'string',
		'Amount Paid To Firm' => 'dollar',
		'Firm Invoice Date' => 'string',
		'Firm File No' => 'string',
		'PYALORGCD' => 'string'
	);

	$style = array(
		'font-style' => 'bold',
		'fill' => '#eee',
		'halign' => 'center',
		'border' => 'left, right, top, bottom',
		'widths' => [20, 20, 20, 20, 20, 20, 20, 30, 20, 20, 20, 20, 20, 20, 20, 20, 20]
	);
	$style1 = array(
		'font-style' => 'bold',
		'fill' => '#eee',
		'font-size' => '10.5',
		'height' => '16.5'
	);
	$sheetName = 'FirmcostCheckDetailToMyDownload';

	return (['headers' => $header, 'style' => $style, 'style1' => $style1, 'sheetName' => $sheetName]);
}