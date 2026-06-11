<?php
/**
 * ============================================================================
 * FIRM FEE CHECK DETAIL REPORT GENERATOR
 * ============================================================================
 * 
 * @author      KEANT Technologies
 * @description Generates Excel reports for Firm Fee transactions from RMAACABHS
 *              database. Processes remittance transactions (RMSTRANCDE='50'-'59')
 *              and creates detailed financial breakdowns per client code.
 * 
 * @version     1.0
 * @created     2026-01-26
 * 
 * WORKFLOW:
 * 1. Accepts comma-separated client codes and optional date parameter
 * 2. Loops through each client code (VENDORNUM)
 * 3. Queries distinct PYALORGCD for each firm
 * 4. Retrieves transaction details: account info, payments, remittances, fees
 * 5. Generates Excel worksheet with:
 *    - Client account details (names, account numbers)
 *    - Transaction breakdown (dates, codes, descriptions)
 *    - Financial data (payment amounts, fees requested/paid, amounts paid to firm)
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
 * ----------------------------------------------------------------------------
 */

require_once('PHP_XLSXWriter/xlsxwriter.class.php');
// use XLSXWriter;

function firmFeeCheckDetailToMyDownload($path, $id, $reportName, $code_name, $userType, $userReportName, $outputName, $reportDescription, $mailNotification, $sftpId, $mode, $run_by, $reportBasePath, $reportDate = null)
{
	$report_start_time = date("H:i:s");
	$date = new DateTime();
	$fileName = filterReportName($date->format('Y-m-d'), $reportName);
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
		if($currentDay == 'Mon'){
			$queryDate = date('Ymd', strtotime('-3 day'));
		} else {
			$queryDate = date('Ymd', strtotime('-1 day'));
		}
	}


	foreach ($paths as $companyPath) {
		$companyPath = str_replace(["'", " "], "", $companyPath);
		$RemitAmount = '0';
		$PaymentAmount = '0';
		$FeeRequestedByFirm = '0';
		$FeePaidToFirm = '0';
		$AmountPaidToFirm = '0';
		$writer = new XLSXWriter();
		$companyName = createDirgetCompanyName($companyPath, $reportBasePath);
		echo "Processing Company: " . $companyName . "\n";
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
			WHERE RMSTRANCDE >= '50' AND RMSTRANCDE <= '59' 
			AND CAST(DTPRLT AS DATE) >= (case when weekday(CURRENT_DATE()) = 0 then DATE_SUB(CURRENT_DATE(),INTERVAL 3 DAY) else DATE_SUB(CURRENT_DATE(), INTERVAL 1 DAY) end)
			AND VENDORNUM IN ('" . $companyName . "') group by PYALORGCD ORDER BY  PYALORGCD";
		*/
		
		/**
		 * NEW QUERY (PARAMETER-BASED DATE)
		 * Uses $queryDate variable which can be:
		 * - Provided via $reportDate parameter, OR
		 * - Auto-calculated based on current day
		 */
		$query = "SELECT DISTINCT PYALORGCD
			from RMAACABHS
			WHERE RMSTRANCDE >= '50' AND RMSTRANCDE <= '59' 
			AND CAST(DTPRLT AS DATE) >= '" . $queryDate . "'
			AND VENDORNUM IN ('" . $companyName . "') group by PYALORGCD ORDER BY  PYALORGCD";

		$results = getResult($query);
		if ($results['numRows'] > 0) {
			$excelPrefix = getExcelPrefix14();
			$writer->writeSheetHeader($excelPrefix['sheetName'], $excelPrefix['headers'], $excelPrefix['style']);
			$blankrow = array('', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '');
			$writer->writeSheetRow($excelPrefix['sheetName'], $blankrow);
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
				$query = "WITH FIRMCOSTFEE AS (
					select BLAAINNM, LEFT(BLAAINNM, 1) as TYPERT, CUROFFCRCD AS 'Client_Code',VENDORNUM as 'Firm',
					RMSACCTNUM AS 'Acct_Number',RMSCORPNM1 AS 'Last_Name', RMSCORPNM2 AS 'First_Name',
					RMSTRANCDE AS 'TR_CD', RMSTRANDTE as  'Transaction_Date', RMSTRANDSC AS 'Transaction_Description',
					IF (RMSTRANCDE = '1A' , INVCLIENT , 0.00) AS 'Cost_Amount', COLLAM AS 'Payment_Amount', DUECLIENT AS 'Remit_Amount',
					FEESFR AS 'Fee_Requested_by_Firm', FEES AS 'Fee_Paid_to_Firm', INVOICENO AS 'Firm_Invoice_No',
					CHECKNO AS 'Firm_Check_Number', BLPYFRCK AS 'AACA_Check_Number',
					IF (RMSTRANCDE = '51' , (BLPYFRTO - (0 - DUECLIENT)) , (BLPYFRTO - ( COLLAM - DUECLIENT ))) AS 'Amount_Paid_to_Firm',
					DTPRLT AS 'AACA_Check_Date', RMSFILENUM, PYALORGCD, PAIDDATE  AS 'Firm_Invoice_Date'
					from RMAACABHS
					WHERE RMSTRANCDE >= '50' AND RMSTRANCDE <= '59' 
					AND CAST(DTPRLT AS DATE) >= '" . $queryDate . "'
					AND VENDORNUM ='" . $companyName . "' AND PYALORGCD='" . $result['PYALORGCD'] . "')SELECT Client_Code as 'Client Code', Acct_Number as 'Acct No.', Last_Name as 'Last Name ', First_Name as 'First Name',
					TR_CD as 'TR CD', Transaction_Date as 'Transaction Date', Transaction_Description as 'Transaction Description',
					Payment_Amount AS 'Payment Amount' ,Remit_Amount AS 'Remit Amount',
					Fee_Requested_by_Firm AS 'Fee Requested By Firm',
					Fee_Paid_to_Firm as 'Fee Paid To Firm',Firm_Invoice_No as 'Firm Invoice No', Firm_Check_Number as 'Firm Check Number',
					Amount_Paid_to_Firm as 'Amount Paid To Firm' ,
					Firm_Invoice_Date as 'Firm Invoice Date', P.FILELOCATN AS 'Firm File No',PYALORGCD
					from FIRMCOSTFEE
					left join RMSPMASTER AS P
					ON FIRMCOSTFEE.RMSFILENUM = P.RMSFILENUM
					ORDER BY  PYALORGCD,Client_Code,Firm_Invoice_Date, Firm_Invoice_No";
				$results = getResult($query);

				if ($results['numRows'] > 0) {
					$AmountPaidToFirmwise = '0';
					$PaymentAmountwise = '0';
					$RemitAmountwise = '0';
					$FeePaidToFirmwise = '0';
					$FeeRequestedByFirmwise = '0';
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
			$writer->writeSheetRow($excelPrefix['sheetName'], $newarray, $excelPrefix['style1']);
			$writer->writeToFile(str_replace(__FILE__, $reportBasePath . $companyPath . '/' . $fileName, __FILE__));
			if (file_exists(str_replace(__FILE__, $reportBasePath . $companyPath . '/' . $fileName, __FILE__))) {
				$FileSizeKB = getfileSize($reportBasePath . $companyPath . '/' . $fileName);
				// VK26JAN2026 mailNotifaction($mailNotification, $companyPath, $companyName, $userType, $userReportName, $reportDescription);


				array_push($dataPresents, array(
					'paths' => $companyPath,
					'filename' => $fileName,
					'clientcode' => $companyName
				));
				
				// VK26JAN2026 ifDataPresent($companyStatus, $companyName, $reportName, $report_start_time, $sftpStatus, $FileSizeKB, $companyPath, $run_by,$userType);
				echo "Fee report generated successfully for company: " . $companyName . "\n"; // VK26JAN2026
			}
		} else {
			// VK26JAN2026 ifDataNotPresent($companyStatus, $companyName, $reportName, $report_start_time, $sftpStatus, $FileSizeKB, $companyPath, $run_by,$userType);
			echo "fee : No data present for company: " . $companyName . "\n"; // VK26JAN2026
			array_push($noDataPresents, array(
				'paths' => $companyPath,
				'filename' => $fileName,
				'clientcode' => $companyName
			));
		}
		if (!empty($dataPresents)) {
			$new_status = 2;
			$status_msg = 'generated';
			$status = 1;
		} else {
			$new_status = 3;
			$status_msg = 'Failed';
			$status = 0;
		}
	}

	return array('status' => $status, 'status_msg' => $status_msg, 'new_status' => $new_status);
}
function getExcelPrefix14()
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
	$sheetName = 'FirmFeeCheckDetailToMyDownload';

	return (['headers' => $header, 'style' => $style, 'style1' => $style1, 'sheetName' => $sheetName]);
}
