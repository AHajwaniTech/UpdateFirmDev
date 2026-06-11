<?php
/**
 * ============================================================================
 * FIRM COST & FEE CHECK DETAIL REPORTS - DUAL WRAPPER SCRIPT
 * ============================================================================
 * 
 * @author      KEANT Technologies
 * @description Unified wrapper script to execute both Firm Cost Check Detail 
 *              and Firm Fee Check Detail report generation for multiple client 
 *              codes in a single execution cycle.
 * 
 *              REPORT 1: Firm Cost Check Detail
 *              - Processes court cost transactions (RMSTRANCDE='1A')
 *              - Generates Excel reports with cost-related financial data
 * 
 *              REPORT 2: Firm Fee Check Detail  
 *              - Processes firm fee transactions (RMSTRANCDE='50' to '59')
 *              - Generates Excel reports with remittance and fee data
 * 
 *              Both reports query RMAACABHS database for transactions from the
 *              previous business day (Monday: -3 days, Other days: -1 day) and
 *              create detailed Excel worksheets with client account details,
 *              transaction breakdowns, and financial summaries per client code.
 * 
 * @version     4.0
 * @created     2026-01-26
 * 
 * COMMAND LINE USAGE:
 *   php Wrapper_Firm_Cost_Fee.php [REPORT_TYPE] [DATE]
 * 
 *   Parameters:
 *     REPORT_TYPE - Which report to run: 'cost', 'fee', or 'both' (default: both)
 *     DATE        - Report date in YYYYMMDD format (optional, auto-calculates if omitted)
 * 
 *   Examples:
 *     php Wrapper_Firm_Cost_Fee.php                    # Run both reports, auto-date
 *     php Wrapper_Firm_Cost_Fee.php both               # Run both reports, auto-date
 *     php Wrapper_Firm_Cost_Fee.php cost 20250116      # Cost report for Jan 16, 2025
 *     php Wrapper_Firm_Cost_Fee.php fee 20250116       # Fee report for Jan 16, 2025
 *     php Wrapper_Firm_Cost_Fee.php both 20250116      # Both reports for Jan 16, 2025
 * 
 * CHANGELOG:
 * ---------------------------------------------------------------------------
 * Version | Date       | Author              | Description
 * ---------------------------------------------------------------------------
 * 1.0     | 2026-01-26 | KEANT Technologies  | Initial version with proper
 *         |            |                     | documentation for Cost report
 * ---------------------------------------------------------------------------
 * 2.0     | 2026-01-26 | KEANT Technologies  | Enhanced to dual-execution
 *         |            |                     | wrapper - now processes both
 *         |            |                     | Cost and Fee reports for same
 *         |            |                     | client codes with comprehensive
 *         |            |                     | summary reporting
 * ---------------------------------------------------------------------------
 * 3.0     | 2026-01-26 | KEANT Technologies  | Added command-line parameter
 *         |            |                     | support for report type selection
 *         |            |                     | (cost/fee/both) and custom date
 *         |            |                     | override functionality
 * ---------------------------------------------------------------------------
 * 4.0     | 2026-05-08 | AH08052026          | AS400 pre-flight validation. Added
 *         |            |                     | connectAS400Wrapper(), preFlightCount(),
 *         |            |                     | resolveQueryDate(). Per-firm COUNT(*) check
 *         |            |                     | before report execution to prevent log
 *         |            |                     | flooding on high-frequency cron jobs.
 *         |            |                     | Added sendConsolidatedMail() function.
 *         |            |                     | Sends unified status email with summary
 *         |            |                     | of generated and skipped firms.
 * ---------------------------------------------------------------------------
 * 4.1     | 2026-06-09 | AH09062026          | Added getLatestDTPRLT() to fetch latest
 *         |            |                     | processing date from AS400. Detects
 *         |            |                     | new/fresh data between execution cycles.
 * ---------------------------------------------------------------------------
 * 4.2     | 2026-06-10 | AH10062026          | Added getRecordCount() and MySQL state
 *         |            |                     | tracking (mysqlGetProcessedInfo,
 *         |            |                     | mysqlMarkSuccess). Prevents duplicate
 *         |            |                     | reports by comparing AS400 data with
 *         |            |                     | stored history.
 * ---------------------------------------------------------------------------
 */

require_once('FirmCostCheckDetailToMyDownload_KNT_POC.php');
require_once('FirmFeeCheckDetailToMyDownload_KNT_POC.php');
require_once('/var/www/html/bi/dist/Report/generateReportManual.php');


// DATABASE CONFIGURATION
define('WRAPPER_LOG_ONLY', true);
if (!defined('AS400_DSN'))  { define('AS400_DSN',  'DB2'); }
if (!defined('AS400_HOST')) { define('AS400_HOST', '192.168.21.8'); }
if (!defined('AS400_USER')) { define('AS400_USER', 'TESTADMIN'); }
if (!defined('AS400_PASS')) { define('AS400_PASS', 'x1ns8@p5d7'); }

// AH08062026 - Configurable email recipients for consolidated report
$CONSOLIDATED_MAIL_RECIPIENTS = [
    'adnan.hajwani.tech@gmail.com'
];

// MySQL connection details for state tracking
define('MYSQL_HOST', '192.168.13.167');
define('MYSQL_PORT', '3306');
define('MYSQL_USER', 'pipewaydb');
define('MYSQL_PASS', 'A8EK5hKjq*CX&w');
define('MYSQL_DB',   'aaca_live');


// AS400 CONNECTION
function connectAS400Wrapper() {
    $conn = @odbc_connect(AS400_DSN, AS400_USER, AS400_PASS);
    if (!$conn) {
        $err = odbc_errormsg();
        writeLogfee('[WRAPPER] AS400 ODBC connection failed: ' . $err, 'ERROR');
        throw new Exception('[WRAPPER] AS400 ODBC connection failed: ' . $err);
    }
    return $conn;
}

// AH09062026: Get latest processing date from AS400
function getLatestDTPRLT($whereFilter) {
    try {
        $conn = connectAS400Wrapper();
        $sql = "SELECT MAX(DTPRLT) AS MAXDATE FROM SABINESH.RMAACABLF WHERE {$whereFilter}";
        $stmt = odbc_exec($conn, $sql);

        if (!$stmt) {
            writeLogfee('[DTPRLT] Query failed: ' . odbc_errormsg($conn), 'ERROR');
            odbc_close($conn);
            return false;
        }

        $dtprlt = false;
        if (odbc_fetch_row($stmt)) {
            $dtprlt = trim(odbc_result($stmt, 'MAXDATE'));
        }

        odbc_free_result($stmt);
        odbc_close($conn);
        return $dtprlt;

    } catch (Exception $e) {
        writeLogfee('[DTPRLT] Exception: ' . $e->getMessage(), 'ERROR');
        return false;
    }
}

// AH10062026: Count records for a given processing date
function getRecordCount($whereFilter, $dtprlt) {
    try {
        $conn = connectAS400Wrapper();
        $sql = "SELECT COUNT(*) AS CNT FROM SABINESH.RMAACABLF WHERE {$whereFilter} AND DTPRLT = '{$dtprlt}'";
        $stmt = odbc_exec($conn, $sql);
        
        if (!$stmt) {
            odbc_close($conn);
            return 0;
        }
        
        $count = 0;
        if (odbc_fetch_row($stmt)) {
            $count = (int)odbc_result($stmt, 'CNT');
        }
        
        odbc_free_result($stmt);
        odbc_close($conn);
        return $count;
        
    } catch (Exception $e) {
        return 0;
    }
}

// Retrieve last processed date and record count from MySQL
function mysqlGetProcessedInfo($reportType) {
    $conn = mysqli_connect(MYSQL_HOST, MYSQL_USER, MYSQL_PASS, MYSQL_DB, (int)MYSQL_PORT);
    if (!$conn) {
        return false;
    }    
    $reportType = mysqli_real_escape_string($conn, $reportType);
    $sql = "SELECT dtprlt, record_count FROM firm_report_history 
            WHERE report_type='{$reportType}' ORDER BY updated_at DESC, generated_at DESC LIMIT 1";
    $result = mysqli_query($conn, $sql);
    $data = false;    
    if ($result && mysqli_num_rows($result) > 0) {
        $row = mysqli_fetch_assoc($result);
        $data = ['dtprlt' => $row['dtprlt'], 'record_count' => (int)$row['record_count']];
    }    
    mysqli_close($conn);
    return $data;
}

// Record successful report generation in MySQL
function mysqlMarkSuccess($reportType, $dtprlt, $recordCount) {
    $conn = mysqli_connect(MYSQL_HOST, MYSQL_USER, MYSQL_PASS, MYSQL_DB, (int)MYSQL_PORT);
    if (!$conn) {
        return;
    }    
    $reportType = mysqli_real_escape_string($conn, $reportType);
    $dtprlt = mysqli_real_escape_string($conn, $dtprlt);
    $recordCount = (int)$recordCount;    
    $sql = "INSERT INTO firm_report_history (report_type,dtprlt,record_count,status,generated_at) 
            VALUES('{$reportType}','{$dtprlt}',{$recordCount},'SUCCESS',NOW())
            ON DUPLICATE KEY UPDATE record_count = VALUES(record_count), updated_at = NOW(), status='SUCCESS'";
    
    mysqli_query($conn, $sql);
    mysqli_close($conn);
}

// Check if firm has data before processing
function preFlightCount($reportLabel, $vendorNum, $queryDate, $whereFilter) {
    $sql = "SELECT COUNT(*) AS CNT FROM SABINESH.RMAACABLF WHERE " . $whereFilter . 
           " AND DTPRLT >= '" . $queryDate . "' AND VENDORNUM = '" . $vendorNum . "'";

    if (!defined('WRAPPER_LOG_ONLY') || WRAPPER_LOG_ONLY === false) {
        writeLogfee('[PRE-FLIGHT] ' . $reportLabel . ' | Firm: ' . $vendorNum . ' | Checking AS400...', 'INFO');
    }    
    try {
        $conn = connectAS400Wrapper();
        $stmt = @odbc_exec($conn, $sql);        
        if (!$stmt) {
            $err = odbc_errormsg($conn);
            odbc_close($conn);
            writeLogfee('[PRE-FLIGHT] ' . $reportLabel . ' | Firm: ' . $vendorNum . ' | Query failed: ' . $err, 'ERROR');
            return false;
        }
        $count = 0;
        if (odbc_fetch_row($stmt)) {
            $count = (int)odbc_result($stmt, 'CNT');
        }        
        odbc_free_result($stmt);
        odbc_close($conn);
        if ($count > 0) {
            if (!defined('WRAPPER_LOG_ONLY') || WRAPPER_LOG_ONLY === false) {
                writeLogfee('[PRE-FLIGHT] ' . $reportLabel . ' | Firm: ' . $vendorNum . ' | Data found: ' . $count . ' row(s)', 'INFO');
            }
            return true;
        } else {
            if (!defined('WRAPPER_LOG_ONLY') || WRAPPER_LOG_ONLY === false) {
                writeLogfee('[PRE-FLIGHT] ' . $reportLabel . ' | Firm: ' . $vendorNum . ' | No data (COUNT=0)', 'INFO');
            }
            return false;
        }
    } catch (Exception $e) {
        writeLogfee('[PRE-FLIGHT] ' . $reportLabel . ' | Firm: ' . $vendorNum . ' | Exception: ' . $e->getMessage(), 'ERROR');
        return false;
    }
}

// Use provided date or default to today
function resolveQueryDate($reportDate) {
    if ($reportDate !== null && $reportDate !== '') {
        return $reportDate;
    }
    return date('Y-m-d');
}
// Execute MySQL query and return first column value
function mysqlQuery($sql) {
    $conn = @mysqli_connect(MYSQL_HOST, MYSQL_USER, MYSQL_PASS, MYSQL_DB, (int)MYSQL_PORT);
    if (!$conn) {
        writeLogfee('[MYSQL] Connection failed: ' . mysqli_connect_error(), 'ERROR');
        return null;
    }    
    $result = mysqli_query($conn, $sql);
    $value = null;    
    if ($result && mysqli_num_rows($result) > 0) {
        $row = mysqli_fetch_row($result);
        $value = $row[0] ?? null;
    }    
    mysqli_close($conn);
    return $value;
}

// AH08062026: Send consolidated report via email with summary of generated and skipped firms
function sendConsolidatedMail($costGenerated, $feeGenerated, $costSkipped, $feeSkipped, $runDate) {
    $hasCost = !empty($costGenerated);
    $hasFee = !empty($feeGenerated);

    if (!$hasCost && !$hasFee) {
        writeLogfee('[MAIL] No new reports generated  mail not sent.', 'WRAPPER');
        return;
    }

    // Build email subject
    if ($hasCost && $hasFee) {
        $subject = 'Firm Cost & Fee Reports Generated  ' . $runDate;
    } elseif ($hasCost) {
        $subject = 'Firm Cost Reports Generated  ' . $runDate;
    } else {
        $subject = 'Firm Fee Reports Generated  ' . $runDate;
    }
    // Build email body
    $body = '<p>The cost and fee checks were run successfully.</p>';
    $body .= '<p><strong>Run Date:</strong> ' . $runDate . '</p>';
    $body .= '<hr>';
    if ($hasCost) {
        $costCount = count($costGenerated);
        sort($costGenerated);
        $body .= '<p><strong>' . $costCount . ' firm(s)</strong>  Cost reports generated:</p>';
        $body .= '<p>' . implode(', ', $costGenerated) . '</p>';
    } else {
        $body .= '<p><strong>Cost Reports:</strong> No new reports.</p>';
    }    
    $body .= '<hr>';
    if ($hasFee) {
        $feeCount = count($feeGenerated);
        sort($feeGenerated);
        $body .= '<p><strong>' . $feeCount . ' firm(s)</strong>  Fee reports generated:</p>';
        $body .= '<p>' . implode(', ', $feeGenerated) . '</p>';
    } else {
        $body .= '<p><strong>Fee Reports:</strong> No new reports.</p>';
    }

    // Add skipped firms section
    if (!empty($costSkipped) || !empty($feeSkipped)) {
        $body .= '<hr>';
        $body .= '<p><em>Skipped :</em></p>';
        if (!empty($costSkipped)) {
            sort($costSkipped);
            $body .= '<p>Cost (' . count($costSkipped) . '): ' . implode(', ', $costSkipped) . '</p>';
        }
        if (!empty($feeSkipped)) {
            sort($feeSkipped);
            $body .= '<p>Fee (' . count($feeSkipped) . '): ' . implode(', ', $feeSkipped) . '</p>';
        }
    }
    $body .= '<hr>';
    $body .= '<p>Kindly review and let us know if further action is required.</p>';

    // Send email using PHPMailer
    try {
        include_once('/var/www/html/bi/dist/mailsetup.php');
        global $CONSOLIDATED_MAIL_RECIPIENTS;
        
        if (!empty($CONSOLIDATED_MAIL_RECIPIENTS)) {
            foreach ($CONSOLIDATED_MAIL_RECIPIENTS as $recipient) {
                if (filter_var($recipient, FILTER_VALIDATE_EMAIL)) {
                    $mail->addAddress($recipient);
                }
            }
        } else {
            writeLogfee('[MAIL] No recipients configured', 'ERROR');
            return;
        }
        $mail->isHTML(true);
        $mail->Subject = $subject;
        $mail->Body = $body;
        $mail->send();
        writeLogfee('[MAIL] Sent successfully. Subject: ' . $subject, 'WRAPPER');
        
    } catch (Exception $e) {
        writeLogfee('[MAIL] Send failed: ' . $e->getMessage(), 'ERROR');
    }
}


// SCRIPT INITIALIZATION & VALIDATION
$reportType = isset($argv[1]) ? strtolower(trim($argv[1])) : 'both';
$reportDate = isset($argv[2]) ? trim($argv[2]) : null;
$executionTime = date('Y-m-d H:i:s');
$runId = date('Ymd_His');
writeLogfee("[WRAPPER] ==================================================");
writeLogfee("[WRAPPER] WRAPPER STARTED");
writeLogfee("[WRAPPER] RUN ID         : " . $runId);
writeLogfee("[WRAPPER] Execution Time : " . $executionTime);
writeLogfee("[WRAPPER] Report Type    : " . strtoupper($reportType));
writeLogfee("[WRAPPER] Report Date    : " . ($reportDate ?: "AUTO"));
writeLogfee("[WRAPPER] ==================================================");
// Validate report type
if (!in_array($reportType, ['cost', 'fee', 'both'])) {
    echo "\nERROR: Invalid report type '$reportType'\n";
    echo "Valid options: cost, fee, both\n";
    echo "Usage: php Wrapper_Test_Script.php [REPORT_TYPE] [DATE]\n\n";
    exit(1);
}

// Validate date format (YYYYMMDD) if provided
if ($reportDate !== null && $reportDate !== '') {
    if (!preg_match('/^\d{8}$/', $reportDate)) {
        echo "\nERROR: Invalid date format '$reportDate'. Expected YYYYMMDD.\n\n";
        exit(1);
    }
    
    $year = substr($reportDate, 0, 4);
    $month = substr($reportDate, 4, 2);
    $day = substr($reportDate, 6, 2);
    
    if (!checkdate($month, $day, $year)) {
        echo "\nERROR: Invalid date '$reportDate' (Year: $year, Month: $month, Day: $day)\n\n";
        exit(1);
    }
}


// REPORT CONFIGURATION


$path = "'downloadfile/DCON/AC','downloadfile/WENR/AC','downloadfile/MHBG/AC','downloadfile/MLST/AC','downloadfile/WEBO/AC','downloadfile/FNTN/AC','downloadfile/BUCK/AC','downloadfile/LPRD/AC','downloadfile/MKR/AC','downloadfile/MCPH/AC','downloadfile/EDAB/AC','downloadfile/EATN/AC','downloadfile/STIL/AC'," .
    "'downloadfile/GURT/AC','downloadfile/PRST/AC','downloadfile/APAL/AC','downloadfile/BRET/AC','downloadfile/BXTR/AC','downloadfile/DBLF/AC','downloadfile/BRQA/AC','downloadfile/GRDN/AC','downloadfile/LOVE/AC','downloadfile/EINS/AC','downloadfile/GRWB/AC','downloadfile/EICH/AC','downloadfile/MSNK/AC','downloadfile/MAYE/AC'," .
    "'downloadfile/LVYL/AC','downloadfile/NEKN/AC','downloadfile/STLG/AC','downloadfile/WRTH/AC','downloadfile/VNLO/AC','downloadfile/MRSG/AC','downloadfile/NAJJ/AC','downloadfile/DANG/AC','downloadfile/SLON/AC','downloadfile/CSBC/AC','downloadfile/HAMR/AC','downloadfile/FINK/AC','downloadfile/ROUT/AC','downloadfile/RYAN/AC','downloadfile/MNDL/AC'," .
    "'downloadfile/HDLH/AC','downloadfile/DLWL/AC','downloadfile/BLIT/AC','downloadfile/KIMM/AC','downloadfile/FARR/AC','downloadfile/NEDR/AC'," .
    "'downloadfile/ERWN/AC','downloadfile/LITO/AC','downloadfile/HLVL/AC','downloadfile/NEGO/AC','downloadfile/BREW/AC','downloadfile/BATZ/AC','downloadfile/SHIN/AC','downloadfile/SDEB/AC','downloadfile/PLRS/AC','downloadfile/BOMC/AC','downloadfile/LYON/AC','downloadfile/HKMC/AC','downloadfile/RBNS/AC','downloadfile/LEOP/AC','downloadfile/TSAR/AC'";

$id = 1;
$reportName_Cost = 'FirmCostCheckDetail';
$code_name_Cost = 'FIRMCOST';
$userReportName_Cost = 'Firm Cost Check Detail Report';
$outputName_Cost = 'FirmCostReport';
$reportDescription_Cost = 'Daily firm cost check detail report';

$reportName_Fee = 'FirmFeeCheckDetail';
$code_name_Fee = 'FIRMFEE';
$userReportName_Fee = 'Firm Fee Check Detail Report';
$outputName_Fee = 'FirmFeeReport';
$reportDescription_Fee = 'Daily firm fee check detail report';

$userType = 'admin_KNT';
$mailNotification = 0;
$sftpId = 0;
$mode = 'manual';
$run_by = 'admin_KNT';
$reportBasePath = '/var/www/html/bi/dist/Mako/';


// EXECUTION START
echo "==========================================================================\n";
echo "FIRM REPORT GENERATION - KEANT Technologies\n";
echo "==========================================================================\n";
echo "Report Type: " . strtoupper($reportType) . "\n";
echo "Report Date: " . ($reportDate ?: "Auto-calculated") . "\n";
echo "==========================================================================\n\n";

$result_cost = null;
$result_fee = null;


// FIRM COST CHECK DETAIL REPORT
if ($reportType === 'cost' || $reportType === 'both') {
    echo "--- EXECUTING FIRM COST CHECK DETAIL REPORT ---\n";
    echo "Transaction Type: Court Cost (RMSTRANCDE='1A')\n";
    echo "Processing " . (substr_count($path, ',') + 1) . " client codes...\n\n";

    $preflight_queryDate = resolveQueryDate($reportDate);
    $preflight_where_cost = "RMSTRANCDE='1A'";

    $allPaths = array_map(fn($p) => str_replace(["'", " "], "", $p), explode(",", $path));
    $pathsCostPass = [];
    $pathsCostSkip = [];
    writeLogfee('[PRE-FLIGHT] FirmCost | Query date: ' . $preflight_queryDate, 'INFO');

    // Check each firm for data availability
    foreach ($allPaths as $companyPath) {
        $parts = explode('/', $companyPath);
        $vendorNum = isset($parts[1]) ? $parts[1] : $companyPath;

        if (preFlightCount('FirmCost', $vendorNum, $preflight_queryDate, $preflight_where_cost)) {
            $pathsCostPass[] = "'" . $companyPath . "'";
        } else {
            $pathsCostSkip[] = $vendorNum;
        }
    }
    writeLogfee('[PRE-FLIGHT] FirmCost | Pass: ' . count($pathsCostPass) . ' | Skip: ' . count($pathsCostSkip), 'INFO');

    if (empty($pathsCostPass)) {
        echo "Pre-flight check: No data found. Report SKIPPED.\n\n";
        writeLogfee('[PRE-FLIGHT] FirmCost | All firms returned COUNT=0  report NOT called.', 'INFO');
        $result_cost = ['status' => 0, 'status_msg' => 'Skipped (no data)', 'new_status' => 3,
                        'generated_companies' => [], 'existing_companies' => [], 'failed_companies' => []];
    } else {
        $pathCost = implode(',', $pathsCostPass);
        echo "Pre-flight passed for " . count($pathsCostPass) . " firm(s). Proceeding...\n\n";

        // AH09062026: Get latest processing date and check if data is new
        $costDtprlt = getLatestDTPRLT("RMSTRANCDE='1A'");
        writeLogfee('[DTPRLT] FirmCost Latest=' . $costDtprlt, 'WRAPPER');

        // AH10062026: Compare AS400 data with stored history
        $costCount = getRecordCount("RMSTRANCDE='1A'", $costDtprlt);
        $storedCost = mysqlGetProcessedInfo('FirmCost');
        
        writeLogfee("[MYSQL] FirmCost | AS400 DTPRLT={$costDtprlt} | AS400 Count={$costCount}", 'WRAPPER');
        
        if ($storedCost) {
            writeLogfee("[MYSQL] FirmCost | Stored DTPRLT={$storedCost['dtprlt']} | Stored Count={$storedCost['record_count']}", 'WRAPPER');
        } else {
            writeLogfee("[MYSQL] FirmCost | No previous history found", 'WRAPPER');
        }

        // Skip if same DTPRLT and count already processed
        $skipCost = ($storedCost && $storedCost['dtprlt'] == $costDtprlt && $storedCost['record_count'] >= $costCount);
        
        if ($skipCost) {
            echo "DTPRLT already processed. Report SKIPPED.\n\n";
            writeLogfee("[MYSQL] FirmCost | Count unchanged ({$costCount}) -> SKIPPED", 'WRAPPER');
            writeLogfee('[DTPRLT] FirmCost already processed: ' . $costDtprlt . ' Count=' . $costCount, 'WRAPPER');
            
            $result_cost = [
                'status' => 1,
                'status_msg' => 'Skipped (already processed)',
                'new_status' => 3,
                'generated_companies' => [],
                'existing_companies' => [],
                'failed_companies' => []
            ];
        } else {
            $oldCostCount = $storedCost['record_count'] ?? 0;
            writeLogfee("[MYSQL] FirmCost | New/Fresh Data ({$oldCostCount} -> {$costCount}) -> PROCEEDING", 'WRAPPER');

            // Execute report generation
            $result_cost = firmCostCheckDetailToMyDownload(
                $pathCost,
                $id,
                $reportName_Cost,
                $code_name_Cost,
                $userType,
                $userReportName_Cost,
                $outputName_Cost,
                $reportDescription_Cost,
                $mailNotification,
                $sftpId,
                $mode,
                $run_by,
                $reportBasePath,
                $costDtprlt
            );

            if (!empty($result_cost['status'])) {
                mysqlMarkSuccess('FirmCost', $costDtprlt, $costCount);
                writeLogfee("[MYSQL] FirmCost | count_seen={$costCount} recorded for {$costDtprlt}", 'WRAPPER');
            }
        }
    }

    echo "Status: " . ($result_cost['status'] ? 'Success' : 'Failed') . "\n";
    echo "Message: " . $result_cost['status_msg'] . "\n\n";
} else {
    echo "--- SKIPPING FIRM COST CHECK DETAIL REPORT ---\n\n";
}


// FIRM FEE CHECK DETAIL REPORT
if ($reportType === 'fee' || $reportType === 'both') {
    echo "--- EXECUTING FIRM FEE CHECK DETAIL REPORT ---\n";
    echo "Transaction Type: Firm Fees (RMSTRANCDE='50' to '59')\n";
    echo "Processing " . (substr_count($path, ',') + 1) . " client codes...\n\n";

    if (!isset($preflight_queryDate)) {
        $preflight_queryDate = resolveQueryDate($reportDate);
    }
    $preflight_where_fee = "RMSTRANCDE>='50' AND RMSTRANCDE<='59'";
    if (!isset($allPaths)) {
        $allPaths = array_map(fn($p) => str_replace(["'", " "], "", $p), explode(",", $path));
    }    
    $pathsFeePass = [];
    $pathsFeeSkip = [];

    writeLogfee('[PRE-FLIGHT] FirmFee | Query date: ' . $preflight_queryDate, 'INFO');

    // Check each firm for data availability
    foreach ($allPaths as $companyPath) {
        $parts = explode('/', $companyPath);
        $vendorNum = isset($parts[1]) ? $parts[1] : $companyPath;

        if (preFlightCount('FirmFee', $vendorNum, $preflight_queryDate, $preflight_where_fee)) {
            $pathsFeePass[] = "'" . $companyPath . "'";
        } else {
            $pathsFeeSkip[] = $vendorNum;
        }
    }

    writeLogfee('[PRE-FLIGHT] FirmFee | Pass: ' . count($pathsFeePass) . ' | Skip: ' . count($pathsFeeSkip), 'INFO');

    if (empty($pathsFeePass)) {
        echo "Pre-flight check: No data found. Report SKIPPED.\n\n";
        writeLogfee('[PRE-FLIGHT] FirmFee | All firms returned COUNT=0  report NOT called.', 'INFO');
        $result_fee = ['status' => 0, 'status_msg' => 'Skipped (no data)', 'new_status' => 3,
                       'generated_companies' => [], 'existing_companies' => [], 'failed_companies' => []];
    } else {
        $pathFee = implode(',', $pathsFeePass);
        echo "Pre-flight passed for " . count($pathsFeePass) . " firm(s). Proceeding...\n\n";

        // Get latest processing date and check if data is new
        $feeDtprlt = getLatestDTPRLT("RMSTRANCDE BETWEEN '50' AND '59'");
        writeLogfee('[DTPRLT] FirmFee Latest=' . $feeDtprlt, 'WRAPPER');

        // Compare AS400 data with stored history
        $feeCount = getRecordCount("RMSTRANCDE BETWEEN '50' AND '59'", $feeDtprlt);
        $storedFee = mysqlGetProcessedInfo('FirmFee');

        writeLogfee("[MYSQL] FirmFee | AS400 DTPRLT={$feeDtprlt} | AS400 Count={$feeCount}", 'WRAPPER');
        
        if ($storedFee) {
            writeLogfee("[MYSQL] FirmFee | Stored DTPRLT={$storedFee['dtprlt']} | Stored Count={$storedFee['record_count']}", 'WRAPPER');
        } else {
            writeLogfee("[MYSQL] FirmFee | No previous history found", 'WRAPPER');
        }

        // Skip if same DTPRLT and count already processed
        $skipFee = ($storedFee && $storedFee['dtprlt'] == $feeDtprlt && $storedFee['record_count'] >= $feeCount);
        
        if ($skipFee) {
            echo "DTPRLT already processed. Report SKIPPED.\n\n";
            writeLogfee("[MYSQL] FirmFee | Count unchanged ({$feeCount}) -> SKIPPED", 'WRAPPER');
            writeLogfee('[DTPRLT] FirmFee already processed: ' . $feeDtprlt . ' Count=' . $feeCount, 'WRAPPER');
            
            $result_fee = [
                'status' => 1,
                'status_msg' => 'Skipped (already processed)',
                'new_status' => 3,
                'generated_companies' => [],
                'existing_companies' => [],
                'failed_companies' => []
            ];
        } else {
            $oldFeeCount = $storedFee['record_count'] ?? 0;
            writeLogfee("[MYSQL] FirmFee | New/Fresh Data ({$oldFeeCount} -> {$feeCount}) -> PROCEEDING", 'WRAPPER');

            // Execute report generation
            $result_fee = firmFeeCheckDetailToMyDownload(
                $pathFee,
                $id,
                $reportName_Fee,
                $code_name_Fee,
                $userType,
                $userReportName_Fee,
                $outputName_Fee,
                $reportDescription_Fee,
                $mailNotification,
                $sftpId,
                $mode,
                $run_by,
                $reportBasePath,
                $feeDtprlt
            );

            if (!empty($result_fee['status'])) {
                mysqlMarkSuccess('FirmFee', $feeDtprlt, $feeCount);
                writeLogfee("[MYSQL] FirmFee | count_seen={$feeCount} recorded for {$feeDtprlt}", 'WRAPPER');
            }
        }
    }

    echo "Status: " . ($result_fee['status'] ? 'Success' : 'Failed') . "\n";
    echo "Message: " . $result_fee['status_msg'] . "\n";
    echo "Generated Firms: " . count($result_fee['generated_companies'] ?? []) . "\n";
    echo "Skipped Firms: " . count($result_fee['existing_companies'] ?? []) . "\n\n";
}


// EXECUTION SUMMARY & REPORTING


echo "==========================================================================\n";
echo "OVERALL EXECUTION SUMMARY\n";
echo "==========================================================================\n";

$total_executed = 0;
$total_successful = 0;
$exit_code = 0;

if ($result_cost !== null) {
    echo "Firm Cost Check Detail: " . ($result_cost['status'] ? 'SUCCESS' : 'FAILED') . "\n";
    $total_executed++;
    if ($result_cost['status']) { $total_successful++; }
    else { $exit_code = 1; }
}

if ($result_fee !== null) {
    echo "Firm Fee Check Detail: " . ($result_fee['status'] ? 'SUCCESS' : 'FAILED') . "\n";
    $total_executed++;
    if ($result_fee['status']) { $total_successful++; }
    else { $exit_code = 1; }
}

echo "\nTotal Executed: $total_executed | Successful: $total_successful | Failed: " . ($total_executed - $total_successful) . "\n";
echo "==========================================================================\n\n";

$totalGenerated = count($result_cost['generated_companies'] ?? []) + count($result_fee['generated_companies'] ?? []);
$totalExisting = count($result_cost['existing_companies'] ?? []) + count($result_fee['existing_companies'] ?? []);
$totalFailed = count($result_cost['failed_companies'] ?? []) + count($result_fee['failed_companies'] ?? []);

writeLogfee("[WRAPPER] ==================================================", 'WRAPPER');
writeLogfee("[WRAPPER] Total Generated Reports : {$totalGenerated}", 'WRAPPER');
writeLogfee("[WRAPPER] Total Existing Reports : {$totalExisting}", 'WRAPPER');
writeLogfee("[WRAPPER] Total Failed Reports : {$totalFailed}", 'WRAPPER');

if (!empty($result_cost['generated_companies'])) {
    writeLogfee("[WRAPPER] Cost Generated Firms : " . implode(', ', $result_cost['generated_companies']), 'WRAPPER');
}

if (!empty($result_cost['existing_companies'])) {
    writeLogfee("[WRAPPER] Cost Existing Firms : " . implode(', ', $result_cost['existing_companies']), 'WRAPPER');
}

if (!empty($result_fee['generated_companies'])) {
    writeLogfee("[WRAPPER] Fee Generated Firms : " . implode(', ', $result_fee['generated_companies']), 'WRAPPER');
}

if (!empty($result_fee['existing_companies'])) {
    writeLogfee("[WRAPPER] Fee Existing Firms : " . implode(', ', $result_fee['existing_companies']), 'WRAPPER');
}

writeLogfee("[WRAPPER] ==================================================", 'WRAPPER');
writeLogfee("[WRAPPER] WRAPPER COMPLETED", 'WRAPPER');
writeLogfee("[WRAPPER] Cost Status : " . ($result_cost['status_msg'] ?? 'N/A'), 'WRAPPER');
writeLogfee("[WRAPPER] Fee Status  : " . ($result_fee['status_msg'] ?? 'N/A'), 'WRAPPER');
writeLogfee("[WRAPPER] Completion Time : " . date('Y-m-d H:i:s'), 'WRAPPER');
writeLogfee("[WRAPPER] ==================================================", 'WRAPPER');

// AH08062026: Send consolidated status email
$mail_cost_generated = $result_cost['generated_companies'] ?? [];
$mail_fee_generated = $result_fee['generated_companies'] ?? [];
$mail_cost_skipped = array_merge($pathsCostSkip ?? [], $result_cost['existing_companies'] ?? []);
$mail_fee_skipped = array_merge($pathsFeeSkip ?? [], $result_fee['existing_companies'] ?? []);
$mail_run_date = $preflight_queryDate ?? date('Y-m-d');

if (!empty($mail_cost_generated) || !empty($mail_fee_generated)) {
    sendConsolidatedMail($mail_cost_generated, $mail_fee_generated, $mail_cost_skipped, $mail_fee_skipped, $mail_run_date);
} else {
    writeLogfee('[MAIL] No new files generated  mail suppressed.', 'WRAPPER');
}

exit($exit_code);
?>