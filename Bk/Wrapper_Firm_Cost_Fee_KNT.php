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
 * @version     3.0
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
 * ----------------------------------------------------------------------------
 * Version | Date       | Author              | Description
 * ----------------------------------------------------------------------------
 * 1.0     | 2026-01-26 | KEANT Technologies  | Initial version with proper
 *         |            |                     | documentation for Cost report
 * ----------------------------------------------------------------------------
 * 2.0     | 2026-01-26 | KEANT Technologies  | Enhanced to dual-execution
 *         |            |                     | wrapper - now processes both
 *         |            |                     | Cost and Fee reports for same
 *         |            |                     | client codes with comprehensive
 *         |            |                     | summary reporting
 * ----------------------------------------------------------------------------
 * 3.0     | 2026-01-26 | KEANT Technologies  | Added command-line parameter
 *         |            |                     | support for report type selection
 *         |            |                     | (cost/fee/both) and custom date
 *         |            |                     | override functionality
 * ----------------------------------------------------------------------------
 */

// ============================================================================
// DEPENDENCY INCLUDES
// ============================================================================

// Core report generation functions
require_once('FirmCostCheckDetailToMyDownload_KNT.php');  // Court Cost report (RMSTRANCDE='1A')
require_once('FirmFeeCheckDetailToMyDownload_KNT.php');   // Firm Fee report (RMSTRANCDE='50'-'59')

// Database and helper functions
require_once('/var/www/html/bi/dist/Report/generateReportManual.php'); // For getResult(), etc.
// require_once('path/to/your/helper_functions.php');   // For other utilities

// ============================================================================
// COMMAND LINE ARGUMENTS PARSING
// ============================================================================

/**
 * Parse command line arguments
 * 
 * Argument 1: Report Type (optional, default: 'both')
 *   - 'cost' = Run only Firm Cost Check Detail Report
 *   - 'fee'  = Run only Firm Fee Check Detail Report
 *   - 'both' = Run both reports (default)
 * 
 * Argument 2: Report Date (optional, format: YYYYMMDD)
 *   - If provided, uses this date for report filtering
 *   - If omitted, auto-calculates based on current day (Mon: -3, Other: -1)
 */

$reportType = isset($argv[1]) ? strtolower(trim($argv[1])) : 'both';
$reportDate = isset($argv[2]) ? trim($argv[2]) : null;

// Validate report type
if (!in_array($reportType, ['cost', 'fee', 'both'])) {
    echo "\n========================================================================\n";
    echo "ERROR: Invalid report type '$reportType'\n";
    echo "========================================================================\n";
    echo "Valid options: cost, fee, both\n\n";
    echo "USAGE: php Wrapper_Firm_Cost_Fee.php [REPORT_TYPE] [DATE]\n\n";
    echo "Examples:\n";
    echo "  php Wrapper_Firm_Cost_Fee.php both\n";
    echo "  php Wrapper_Firm_Cost_Fee.php cost 20250116\n";
    echo "  php Wrapper_Firm_Cost_Fee.php fee 20250116\n";
    echo "  php Wrapper_Firm_Cost_Fee.php both 20250116\n";
    echo "========================================================================\n\n";
    exit(1);
}

// Validate date format if provided
if ($reportDate !== null && $reportDate !== '') {
    if (!preg_match('/^\d{8}$/', $reportDate)) {
        echo "\n========================================================================\n";
        echo "ERROR: Invalid date format '$reportDate'\n";
        echo "========================================================================\n";
        echo "Date must be in YYYYMMDD format (e.g., 20250116)\n";
        echo "========================================================================\n\n";
        exit(1);
    }
    
    // Additional date validation
    $year = substr($reportDate, 0, 4);
    $month = substr($reportDate, 4, 2);
    $day = substr($reportDate, 6, 2);
    
    if (!checkdate($month, $day, $year)) {
        echo "\n========================================================================\n";
        echo "ERROR: Invalid date '$reportDate'\n";
        echo "========================================================================\n";
        echo "Year: $year, Month: $month, Day: $day\n";
        echo "Please provide a valid date in YYYYMMDD format\n";
        echo "========================================================================\n\n";
        exit(1);
    }
}

// ============================================================================
// CONFIGURATION PARAMETERS
// ============================================================================

/**
 * Client Codes List
 * Comma-separated list of firm/client codes to process
 * Each code must be wrapped in single quotes with format: 'downloadfile/CLIENT_CODE/AC'
 */

$path = "'downloadfile/DCON/AC','downloadfile/WENR/AC','downloadfile/MHBG/AC','downloadfile/MLST/AC','downloadfile/WEBO/AC','downloadfile/FNTN/AC','downloadfile/BUCK/AC','downloadfile/LPRD/AC','downloadfile/MKR/AC','downloadfile/MCPH/AC','downloadfile/EDAB/AC','downloadfile/EATN/AC','downloadfile/STIL/AC'," .
        "'downloadfile/GURT/AC','downloadfile/PRST/AC','downloadfile/APAL/AC','downloadfile/BRET/AC','downloadfile/BXTR/AC','downloadfile/DBLF/AC','downloadfile/BRQA/AC','downloadfile/GRDN/AC','downloadfile/LOVE/AC','downloadfile/EINS/AC','downloadfile/GRWB/AC','downloadfile/EICH/AC','downloadfile/MSNK/AC','downloadfile/MAYE/AC'," .
        "'downloadfile/LVYL/AC','downloadfile/NEKN/AC','downloadfile/STLG/AC','downloadfile/WRTH/AC','downloadfile/VNLO/AC','downloadfile/MRSG/AC','downloadfile/NAJJ/AC','downloadfile/DANG/AC','downloadfile/SLON/AC','downloadfile/CSBC/AC','downloadfile/HAMR/AC','downloadfile/FINK/AC','downloadfile/ROUT/AC','downloadfile/RYAN/AC','downloadfile/MNDL/AC'," .
        "'downloadfile/HDLH/AC','downloadfile/DLWL/AC','downloadfile/BLIT/AC','downloadfile/KIMM/AC','downloadfile/FARR/AC','downloadfile/NEDR/AC'," .
        "'downloadfile/ERWN/AC','downloadfile/LITO/AC','downloadfile/HLVL/AC','downloadfile/NEGO/AC','downloadfile/BREW/AC','downloadfile/BATZ/AC','downloadfile/SHIN/AC','downloadfile/SDEB/AC','downloadfile/PLRS/AC','downloadfile/BOMC/AC','downloadfile/LYON/AC','downloadfile/HKMC/AC','downloadfile/RBNS/AC','downloadfile/LEOP/AC','downloadfile/TSAR/AC'";

// $path = "'downloadfile/BRET/AC'";

/**
 * Report Configuration
 */
$id                 = 1;                                  // Report ID

// Court Cost Report Configuration
$reportName_Cost         = 'FirmCostCheckDetail';              // Internal report name
$code_name_Cost          = 'FIRMCOST';                         // Report code identifier
$userReportName_Cost     = 'Firm Cost Check Detail Report';    // User-friendly report name
$outputName_Cost         = 'FirmCostReport';                   // Output file base name
$reportDescription_Cost  = 'Daily firm cost check detail report'; // Report description

// Firm Fee Report Configuration
$reportName_Fee          = 'FirmFeeCheckDetail';               // Internal report name
$code_name_Fee           = 'FIRMFEE';                          // Report code identifier
$userReportName_Fee      = 'Firm Fee Check Detail Report';     // User-friendly report name
$outputName_Fee          = 'FirmFeeReport';                    // Output file base name
$reportDescription_Fee   = 'Daily firm fee check detail report';  // Report description

// Common configuration for both reports
$userType           = 'admin_KNT';                            // User type executing report

/**
 * Notification & Delivery Settings
 */
$mailNotification   = 0;  // Email notification: 0 = disabled, 1 = enabled
$sftpId             = 0;  // SFTP delivery: 0 = disabled, >0 = SFTP configuration ID

/**
 * Execution Context
 */
$mode               = 'manual';                           // Execution mode: manual/automated
$run_by             = 'admin_KNT';                            // User/process executing the report
$reportBasePath     = '/var/www/html/bi/dist/Mako/'; // Base directory for report output

// ============================================================================
// REPORT GENERATION EXECUTION
// ============================================================================

echo "==========================================================================\n";
echo "        FIRM REPORT GENERATION PROCESS - KEANT Technologies\n";
echo "==========================================================================\n";
echo "Report Type: " . strtoupper($reportType) . "\n";
if ($reportDate !== null && $reportDate !== '') {
    echo "Report Date: $reportDate (Custom Override)\n";
} else {
    echo "Report Date: Auto-calculated (Monday: -3 days, Other: -1 day)\n";
}
echo "==========================================================================\n\n";

// Initialize result variables
$result_cost = null;
$result_fee = null;

// ----------------------------------------------------------------------------
// PART 1: FIRM COST CHECK DETAIL REPORT
// ----------------------------------------------------------------------------

if ($reportType === 'cost' || $reportType === 'both') {
    echo "--- EXECUTING FIRM COST CHECK DETAIL REPORT ---\n";
    echo "Transaction Type: Court Cost (RMSTRANCDE='1A')\n";
    echo "Processing " . substr_count($path, ',') + 1 . " client codes...\n";
    if ($reportDate !== null && $reportDate !== '') {
        echo "Using custom date: $reportDate\n\n";
    } else {
        echo "Using auto-calculated date\n\n";
    }

    /**
     * Execute the firmCostCheckDetailToMyDownload function
     * 
     * This function will:
     * 1. Parse the comma-separated client codes
     * 2. Loop through each client code
     * 3. Query database for court cost transactions (RMSTRANCDE='1A')
     * 4. Generate Excel report for each client with data
     * 5. Save reports to individual client directories
     * 6. Track success/failure status
     */
    $result_cost = firmCostCheckDetailToMyDownload(
        $path,                      // Comma-separated client codes
        $id,                        // Report ID
        $reportName_Cost,           // Report name
        $code_name_Cost,            // Report code
        $userType,                  // User type
        $userReportName_Cost,       // Display name
        $outputName_Cost,           // Output filename
        $reportDescription_Cost,    // Description
        $mailNotification,          // Email flag
        $sftpId,                    // SFTP ID
        $mode,                      // Execution mode
        $run_by,                    // Executor
        $reportBasePath,            // Output path
        $reportDate                 // Report date (can be null for auto-calc)
    );

    echo "\n--- FIRM COST CHECK DETAIL REPORT RESULTS ---\n";
    echo "Status: " . ($result_cost['status'] ? 'Success' : 'Failed') . "\n";
    echo "Message: " . $result_cost['status_msg'] . "\n";
    echo "New Status: " . $result_cost['new_status'] . "\n\n";
} else {
    echo "--- SKIPPING FIRM COST CHECK DETAIL REPORT ---\n";
    echo "(Not selected in report type parameter)\n\n";
}

// ----------------------------------------------------------------------------
// PART 2: FIRM FEE CHECK DETAIL REPORT
// ----------------------------------------------------------------------------

if ($reportType === 'fee' || $reportType === 'both') {
    echo "--- EXECUTING FIRM FEE CHECK DETAIL REPORT ---\n";
    echo "Transaction Type: Firm Fees (RMSTRANCDE='50' to '59')\n";
    echo "Processing " . substr_count($path, ',') + 1 . " client codes...\n";
    if ($reportDate !== null && $reportDate !== '') {
        echo "Using custom date: $reportDate\n\n";
    } else {
        echo "Using auto-calculated date\n\n";
    }

    /**
     * Execute the firmFeeCheckDetailToMyDownload function
     * 
     * This function will:
     * 1. Parse the comma-separated client codes
     * 2. Loop through each client code
     * 3. Query database for firm fee transactions (RMSTRANCDE >= '50' AND <= '59')
     * 4. Generate Excel report for each client with data
     * 5. Save reports to individual client directories
     * 6. Track success/failure status
     */
    $result_fee = firmFeeCheckDetailToMyDownload(
        $path,                      // Comma-separated client codes
        $id,                        // Report ID
        $reportName_Fee,            // Report name
        $code_name_Fee,             // Report code
        $userType,                  // User type
        $userReportName_Fee,        // Display name
        $outputName_Fee,            // Output filename
        $reportDescription_Fee,     // Description
        $mailNotification,          // Email flag
        $sftpId,                    // SFTP ID
        $mode,                      // Execution mode
        $run_by,                    // Executor
        $reportBasePath,            // Output path
        $reportDate                 // Report date (can be null for auto-calc)
    );

    // ========================================================================
    // OUTPUT RESULTS
    // ========================================================================

    echo "\n--- FIRM FEE CHECK DETAIL REPORT RESULTS ---\n";
    echo "Status: " . ($result_fee['status'] ? 'Success' : 'Failed') . "\n";
    echo "Message: " . $result_fee['status_msg'] . "\n";
    echo "New Status: " . $result_fee['new_status'] . "\n\n";
} else {
    echo "--- SKIPPING FIRM FEE CHECK DETAIL REPORT ---\n";
    echo "(Not selected in report type parameter)\n\n";
}

// ----------------------------------------------------------------------------
// SUMMARY OF EXECUTED REPORTS
// ----------------------------------------------------------------------------

echo "==========================================================================\n";
echo "                    OVERALL EXECUTION SUMMARY\n";
echo "==========================================================================\n";

$total_executed = 0;
$total_successful = 0;
$total_failed = 0;
$exit_code = 0;

if ($result_cost !== null) {
    echo "Firm Cost Check Detail Report: " . ($result_cost['status'] ? 'SUCCESS' : 'FAILED') . "\n";
    $total_executed++;
    if ($result_cost['status']) {
        $total_successful++;
    } else {
        $total_failed++;
        $exit_code = 1;
    }
}

if ($result_fee !== null) {
    echo "Firm Fee Check Detail Report:  " . ($result_fee['status'] ? 'SUCCESS' : 'FAILED') . "\n";
    $total_executed++;
    if ($result_fee['status']) {
        $total_successful++;
    } else {
        $total_failed++;
        $exit_code = 1;
    }
}

echo "\nTotal Reports Executed: $total_executed\n";
echo "Successful Reports: $total_successful\n";
echo "Failed Reports: $total_failed\n";
echo "==========================================================================\n";

// Exit with appropriate code (0 = all success, 1 = at least one failure)
exit($exit_code);

// ============================================================================
// END OF SCRIPT
// ============================================================================
?>