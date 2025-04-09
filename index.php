<?php
// Prevent PHP errors from breaking JSON responses
error_reporting(E_ALL);
ini_set('display_errors', 1); // Changed to 1 to see errors during setup

// Add this section for debugging
if (!isset($_POST['action']) && !isset($_GET['code'])) {
    echo "<h2>Debug Information:</h2>";
    echo "<p>Current URL: " . (isset($_SERVER['HTTPS']) && $_SERVER['HTTPS'] === 'on' ? "https" : "http") . "://$_SERVER[HTTP_HOST]$_SERVER[REQUEST_URI]</p>";

    if (file_exists('credentials.json')) {
        echo "<p style='color:green'>credentials.json exists!</p>";
    } else {
        echo "<p style='color:red'>credentials.json does not exist! Please upload it.</p>";
    }

    if (file_exists('token.json')) {
        echo "<p style='color:green'>token.json exists!</p>";
    } else {
        echo "<p style='color:red'>token.json does not exist!</p>";
    }
}

require __DIR__ . '/vendor/autoload.php';

// Configuration
$companySheets = [
    "Company A" => "1wOUaUs4cvu2H9eeGVfr67g0Vs73o5MCGG84bB8MEG8Y",
    "Company B" => ""
];

// Check if we need to manually get a token
if (isset($_GET['code'])) {
    try {
        $client = new Google_Client();
        $client->setApplicationName('Leave Management');
        $client->setScopes(Google_Service_Sheets::SPREADSHEETS);
        $client->setAuthConfig('credentials.json');

        // Get the current URL
        $redirectUri = (isset($_SERVER['HTTPS']) && $_SERVER['HTTPS'] === 'on' ? "https" : "http") . "://$_SERVER[HTTP_HOST]$_SERVER[REQUEST_URI]";
        // Remove the code parameter
        $redirectUri = preg_replace('/\?code=.*$/', '', $redirectUri);

        $client->setRedirectUri($redirectUri);
        $client->setAccessType('offline');

        $authCode = $_GET['code'];

        // Exchange authorization code for an access token.
        $accessToken = $client->fetchAccessTokenWithAuthCode($authCode);
        $client->setAccessToken($accessToken);

        // Store the token to disk.
        if (!file_exists(dirname('token.json'))) {
            mkdir(dirname('token.json'), 0700, true);
        }
        file_put_contents('token.json', json_encode($client->getAccessToken()));
        echo "<div style='background-color: #d4edda; color: #155724; padding: 15px; border-radius: 5px; margin-bottom: 20px;'>
                Token saved! <a href=''>Click here</a> to use the application.
              </div>";
        exit;
    } catch (Exception $e) {
        echo "<div style='background-color: #f8d7da; color: #721c24; padding: 15px; border-radius: 5px; margin-bottom: 20px;'>
                Error: " . htmlspecialchars($e->getMessage()) . "
              </div>";
        exit;
    }
}

// Add a route to initiate authentication if token doesn't exist
if (!file_exists('token.json') && !isset($_POST['action'])) {
    try {
        $client = new Google_Client();
        $client->setApplicationName('Leave Management');
        $client->setScopes(Google_Service_Sheets::SPREADSHEETS);
        $client->setAuthConfig('credentials.json');

        // Get the current URL as redirect URI
        $redirectUri = (isset($_SERVER['HTTPS']) && $_SERVER['HTTPS'] === 'on' ? "https" : "http") . "://$_SERVER[HTTP_HOST]$_SERVER[REQUEST_URI]";

        $client->setRedirectUri($redirectUri);
        $client->setAccessType('offline');
        $client->setPrompt('select_account consent');

        // Request authorization from the user
        $authUrl = $client->createAuthUrl();
        echo "<div style='text-align: center; margin-top: 100px;'>
                <h2>Google Authentication Required</h2>
                <p>To use this application, you need to authorize access to your Google Sheets.</p>
                <a href='$authUrl' style='background-color: #4285F4; color: white; padding: 10px 20px; text-decoration: none; border-radius: 5px;'>
                    Authorize with Google
                </a>
              </div>";
        exit;
    } catch (Exception $e) {
        echo "<div style='background-color: #f8d7da; color: #721c24; padding: 15px; border-radius: 5px; margin-bottom: 20px;'>
                Authentication Error: " . htmlspecialchars($e->getMessage()) . "
              </div>";
        exit;
    }
}

// Google API Setup
function getClient() {
    $client = new Google_Client();
    $client->setApplicationName('Leave Management');
    $client->setScopes(Google_Service_Sheets::SPREADSHEETS);
    $client->setAuthConfig('credentials.json');

    // Get the current URL as redirect URI
    $redirectUri = (isset($_SERVER['HTTPS']) && $_SERVER['HTTPS'] === 'on' ? "https" : "http") . "://$_SERVER[HTTP_HOST]";
    // Remove any query parameters
    $redirectUri = preg_replace('/\?.*$/', '', $redirectUri);

    $client->setRedirectUri($redirectUri);
    $client->setAccessType('offline');

    // Load previously authorized token, if it exists
    $tokenPath = 'token.json';
    if (file_exists($tokenPath)) {
        $accessToken = json_decode(file_get_contents($tokenPath), true);
        $client->setAccessToken($accessToken);
    }

    // If there is no previous token or it's expired
    if ($client->isAccessTokenExpired()) {
        // Refresh the token if possible
        if ($client->getRefreshToken()) {
            $client->fetchAccessTokenWithRefreshToken($client->getRefreshToken());
            file_put_contents($tokenPath, json_encode($client->getAccessToken()));
        } else {
            // Request authorization from the user
            $authUrl = $client->createAuthUrl();
            throw new Exception("Authorization required. Open this URL: " . $authUrl);
        }
    }
    return $client;
}

// The rest of your original code remains unchanged
/**
 * Function to get user ID by email from registered data
 */
function getUserIdByEmail($email, $registeredData) {
    // Skip header row
    for ($i = 1; $i < count($registeredData); $i++) {
        if ($registeredData[$i][4] === $email) { // Column E - Email
            return $registeredData[$i][0]; // Column A - User ID
        }
    }
    return null; // If not found
}

/**
 * Function to get company list
 */
function getCompanyList() {
    global $companySheets;
    return array_keys($companySheets);
}

/**
 * Function to get sheet data
 */
function getSheetData($companyName, $status = null) {
    global $companySheets;

    $sheetId = $companySheets[$companyName] ?? null;
    if (!$sheetId) {
        return [["Error", "Invalid company selected"]];
    }

    try {
        $client = getClient();
        $service = new Google_Service_Sheets($client);

        // Get sheet information
        $spreadsheet = $service->spreadsheets->get($sheetId);
        $sheets = $spreadsheet->getSheets();

        // Get the appropriate sheet
        $sheetName = $status ?: $sheets[0]->getProperties()->getTitle();
        $range = $sheetName;

        // Get data from the sheet
        $response = $service->spreadsheets_values->get($sheetId, $range);
        $data = $response->getValues();

        if (empty($data)) {
            return [["Error", "No data available"]];
        }

        // Format date values (simplified compared to Apps Script)
        foreach ($data as $i => $row) {
            foreach ($row as $j => $cell) {
                // Basic date format checking and formatting
                if (preg_match('/^\d{4}-\d{2}-\d{2}/', $cell)) {
                    $date = date_create($cell);
                    if ($date) {
                        $data[$i][$j] = date_format($date, "m/d/Y");
                    }
                }
            }
        }

        return $data;
    } catch (Exception $e) {
        return [["Error", "Could not load sheet: " . $e->getMessage()]];
    }
}

/**
 * Function to approve a row
 */
function approveRow($companyName, $rowIndex) {
    global $companySheets;

    $sheetId = $companySheets[$companyName] ?? null;
    if (!$sheetId) {
        return "Invalid company selected";
    }

    try {
        $client = getClient();
        $service = new Google_Service_Sheets($client);

        // Get spreadsheet information
        $spreadsheet = $service->spreadsheets->get($sheetId);
        $sheets = $spreadsheet->getSheets();
        $mainSheet = $sheets[0]->getProperties()->getTitle();

        // Check if Approved sheet exists
        $approvedSheetExists = false;
        $approvedSheetIndex = 0;
        foreach ($sheets as $index => $sheet) {
            if ($sheet->getProperties()->getTitle() === "Approved") {
                $approvedSheetExists = true;
                $approvedSheetIndex = $index;
                break;
            }
        }

        // Create Approved sheet if it doesn't exist
        if (!$approvedSheetExists) {
            $requests = [
                new Google_Service_Sheets_Request([
                    'addSheet' => [
                        'properties' => [
                            'title' => 'Approved'
                        ]
                    ]
                ])
            ];

            $batchUpdateRequest = new Google_Service_Sheets_BatchUpdateSpreadsheetRequest([
                'requests' => $requests
            ]);

            $service->spreadsheets->batchUpdate($sheetId, $batchUpdateRequest);
        }

        // Get main sheet data
        $range = "$mainSheet";
        $response = $service->spreadsheets_values->get($sheetId, $range);
        $values = $response->getValues();

        // Get the row to approve
        if (!isset($values[$rowIndex])) {
            return "Row does not exist";
        }
        $rowData = $values[$rowIndex];

        // Update status to Approved
        $rowData[count($rowData) - 1] = "Approved";

        // Append to Approved sheet
        $valueRange = new Google_Service_Sheets_ValueRange();
        $valueRange->setValues([$rowData]);
        $service->spreadsheets_values->append(
            $sheetId,
            'Approved',
            $valueRange,
            ['valueInputOption' => 'USER_ENTERED']
        );

        // Delete row from main sheet
        $requests = [
            new Google_Service_Sheets_Request([
                'deleteDimension' => [
                    'range' => [
                        'sheetId' => $sheets[0]->getProperties()->getSheetId(),
                        'dimension' => 'ROWS',
                        'startIndex' => $rowIndex,
                        'endIndex' => $rowIndex + 1
                    ]
                ]
            ])
        ];

        $batchUpdateRequest = new Google_Service_Sheets_BatchUpdateSpreadsheetRequest([
            'requests' => $requests
        ]);

        $service->spreadsheets->batchUpdate($sheetId, $batchUpdateRequest);

        return "Row approved and moved successfully.";
    } catch (Exception $e) {
        return "Error: " . $e->getMessage();
    }
}

/**
 * Function to decline a row
 */
function declineRow($companyName, $rowIndex) {
    return processRow($companyName, $rowIndex, "Declined");
}

/**
 * Function to process a row
 */
function processRow($companyName, $rowIndex, $status) {
    global $companySheets;

    $sheetId = $companySheets[$companyName] ?? null;
    if (!$sheetId) {
        return "Invalid company selected";
    }

    try {
        $client = getClient();
        $service = new Google_Service_Sheets($client);

        // Get spreadsheet information
        $spreadsheet = $service->spreadsheets->get($sheetId);
        $sheets = $spreadsheet->getSheets();
        $mainSheet = $sheets[0]->getProperties()->getTitle();

        // Check if target sheet exists
        $targetSheetExists = false;
        $targetSheetIndex = 0;
        foreach ($sheets as $index => $sheet) {
            if ($sheet->getProperties()->getTitle() === $status) {
                $targetSheetExists = true;
                $targetSheetIndex = $index;
                break;
            }
        }

        // Create target sheet if it doesn't exist
        if (!$targetSheetExists) {
            $requests = [
                new Google_Service_Sheets_Request([
                    'addSheet' => [
                        'properties' => [
                            'title' => $status
                        ]
                    ]
                ])
            ];

            $batchUpdateRequest = new Google_Service_Sheets_BatchUpdateSpreadsheetRequest([
                'requests' => $requests
            ]);

            $service->spreadsheets->batchUpdate($sheetId, $batchUpdateRequest);
        }

        // Get main sheet data
        $range = "$mainSheet";
        $response = $service->spreadsheets_values->get($sheetId, $range);
        $values = $response->getValues();

        // Get the row to process
        if (!isset($values[$rowIndex])) {
            return "Row does not exist";
        }
        $rowData = $values[$rowIndex];

        // Update status
        $rowData[count($rowData) - 1] = $status;

        // Append to target sheet
        $valueRange = new Google_Service_Sheets_ValueRange();
        $valueRange->setValues([$rowData]);
        $service->spreadsheets_values->append(
            $sheetId,
            $status,
            $valueRange,
            ['valueInputOption' => 'USER_ENTERED']
        );

        // Delete row from main sheet
        $requests = [
            new Google_Service_Sheets_Request([
                'deleteDimension' => [
                    'range' => [
                        'sheetId' => $sheets[0]->getProperties()->getSheetId(),
                        'dimension' => 'ROWS',
                        'startIndex' => $rowIndex,
                        'endIndex' => $rowIndex + 1
                    ]
                ]
            ])
        ];

        $batchUpdateRequest = new Google_Service_Sheets_BatchUpdateSpreadsheetRequest([
            'requests' => $requests
        ]);

        $service->spreadsheets->batchUpdate($sheetId, $batchUpdateRequest);

        return "Row {$status} successfully";
    } catch (Exception $e) {
        return "Error: " . $e->getMessage();
    }
}

/**
 * Function to delete a row
 */
function deleteRow($companyName, $rowIndex, $status) {
    global $companySheets;

    $sheetId = $companySheets[$companyName] ?? null;
    if (!$sheetId) {
        return "Invalid company selected";
    }

    try {
        $client = getClient();
        $service = new Google_Service_Sheets($client);

        // Get spreadsheet information
        $spreadsheet = $service->spreadsheets->get($sheetId);
        $sheets = $spreadsheet->getSheets();

        // Get sheet ID for status
        $statusSheetId = null;
        foreach ($sheets as $sheet) {
            if ($sheet->getProperties()->getTitle() === $status) {
                $statusSheetId = $sheet->getProperties()->getSheetId();
                break;
            }
        }

        if ($statusSheetId === null) {
            return "Sheet '$status' not found";
        }

        // Check if Deleted sheet exists
        $deletedSheetExists = false;
        $deletedSheetId = null;
        foreach ($sheets as $sheet) {
            if ($sheet->getProperties()->getTitle() === "Deleted") {
                $deletedSheetExists = true;
                $deletedSheetId = $sheet->getProperties()->getSheetId();
                break;
            }
        }

        // Create Deleted sheet if it doesn't exist
        if (!$deletedSheetExists) {
            $requests = [
                new Google_Service_Sheets_Request([
                    'addSheet' => [
                        'properties' => [
                            'title' => 'Deleted'
                        ]
                    ]
                ])
            ];

            $batchUpdateRequest = new Google_Service_Sheets_BatchUpdateSpreadsheetRequest([
                'requests' => $requests
            ]);

            $response = $service->spreadsheets->batchUpdate($sheetId, $batchUpdateRequest);
            $deletedSheetId = $response->getReplies()[0]->getAddSheet()->getProperties()->getSheetId();
        }

        // Get row data
        $range = "$status";
        $response = $service->spreadsheets_values->get($sheetId, $range);
        $values = $response->getValues();

        if (!isset($values[$rowIndex])) {
            return "Row does not exist";
        }
        $rowData = $values[$rowIndex];

        // Append to Deleted sheet
        $valueRange = new Google_Service_Sheets_ValueRange();
        $valueRange->setValues([$rowData]);
        $service->spreadsheets_values->append(
            $sheetId,
            'Deleted',
            $valueRange,
            ['valueInputOption' => 'USER_ENTERED']
        );

        // Delete row from status sheet
        $requests = [
            new Google_Service_Sheets_Request([
                'deleteDimension' => [
                    'range' => [
                        'sheetId' => $statusSheetId,
                        'dimension' => 'ROWS',
                        'startIndex' => $rowIndex,
                        'endIndex' => $rowIndex + 1
                    ]
                ]
            ])
        ];

        $batchUpdateRequest = new Google_Service_Sheets_BatchUpdateSpreadsheetRequest([
            'requests' => $requests
        ]);

        $service->spreadsheets->batchUpdate($sheetId, $batchUpdateRequest);

        return "Row moved to Deleted";
    } catch (Exception $e) {
        return "Error: " . $e->getMessage();
    }
}

/**
 * Function to save a comment
 */
function saveComment($companyName, $rowIndex, $comment) {
    global $companySheets;

    $sheetId = $companySheets[$companyName] ?? null;
    if (!$sheetId) {
        return "Invalid company selected";
    }

    try {
        $client = getClient();
        $service = new Google_Service_Sheets($client);

        // Get spreadsheet information
        $spreadsheet = $service->spreadsheets->get($sheetId);
        $sheets = $spreadsheet->getSheets();
        $mainSheet = $sheets[0]->getProperties()->getTitle();

        // Check if Comment sheet exists
        $commentSheetExists = false;
        foreach ($sheets as $sheet) {
            if ($sheet->getProperties()->getTitle() === "Comment") {
                $commentSheetExists = true;
                break;
            }
        }

        // Create Comment sheet if it doesn't exist
        if (!$commentSheetExists) {
            $requests = [
                new Google_Service_Sheets_Request([
                    'addSheet' => [
                        'properties' => [
                            'title' => 'Comment'
                        ]
                    ]
                ])
            ];

            $batchUpdateRequest = new Google_Service_Sheets_BatchUpdateSpreadsheetRequest([
                'requests' => $requests
            ]);

            $service->spreadsheets->batchUpdate($sheetId, $batchUpdateRequest);
        }

        // Get row data
        $range = "$mainSheet";
        $response = $service->spreadsheets_values->get($sheetId, $range);
        $values = $response->getValues();

        if (!isset($values[$rowIndex])) {
            return "Row does not exist";
        }
        $rowData = $values[$rowIndex];

        // Check if we have enough columns
        $name = isset($rowData[4]) ? $rowData[4] : "Unknown"; // Column E - Name
        $email = isset($rowData[1]) ? $rowData[1] : "Unknown"; // Column B - Email

        // Append to Comment sheet
        $valueRange = new Google_Service_Sheets_ValueRange();
        $valueRange->setValues([[$name, $email, $comment]]);
        $service->spreadsheets_values->append(
            $sheetId,
            'Comment',
            $valueRange,
            ['valueInputOption' => 'USER_ENTERED']
        );

        return "Comment saved successfully!";
    } catch (Exception $e) {
        return "Error: " . $e->getMessage();
    }
}

// Check if this is an API request
if (isset($_POST['action'])) {
    try {
        // Add headers at the beginning for AJAX requests
        header('Content-Type: application/json');

        // Process API requests
        $response = [];

        switch ($_POST['action']) {
            case 'getCompanyList':
                $response = getCompanyList();
                break;

            case 'getSheetData':
                $companyName = $_POST['company'] ?? "Company A";
                $status = $_POST['status'] ?? null;
                $response = getSheetData($companyName, $status);
                break;

            case 'approveRow':
                $companyName = $_POST['company'] ?? "Company A";
                $rowIndex = isset($_POST['rowIndex']) ? (int)$_POST['rowIndex'] : 0;
                $response = ['message' => approveRow($companyName, $rowIndex)];
                break;

            case 'declineRow':
                $companyName = $_POST['company'] ?? "Company A";
                $rowIndex = isset($_POST['rowIndex']) ? (int)$_POST['rowIndex'] : 0;
                $response = ['message' => declineRow($companyName, $rowIndex)];
                break;

            case 'deleteRow':
                $companyName = $_POST['company'] ?? "Company A";
                $rowIndex = isset($_POST['rowIndex']) ? (int)$_POST['rowIndex'] : 0;
                $status = $_POST['status'] ?? "";
                $response = ['message' => deleteRow($companyName, $rowIndex, $status)];
                break;

            case 'saveComment':
                $companyName = $_POST['company'] ?? "Company A";
                $rowIndex = isset($_POST['rowIndex']) ? (int)$_POST['rowIndex'] : 0;
                $comment = $_POST['comment'] ?? "";
                $response = ['message' => saveComment($companyName, $rowIndex, $comment)];
                break;

            default:
                $response = ['error' => 'Invalid action'];
        }

        echo json_encode($response);
        exit;
    } catch (Exception $e) {
        // Make sure any errors are also returned as JSON
        echo json_encode(['error' => $e->getMessage()]);
        exit;
    }
}
?>
<!DOCTYPE html>
<html>
<head>
    <meta name="viewport" content="width=device-width, initial-scale=1" />
    <link rel="stylesheet" href="https://cdnjs.cloudflare.com/ajax/libs/animate.css/4.1.1/animate.min.css"/>
    <title>Leave Management</title>
    <style>
      .approve-btn {
        background-color: #4CAF50;
        color: white;
        border: none;
      }

      .decline-btn {
        background-color: #f44336;
        color: white;
        border: none;
      }

      .delete-btn {
        background-color: #808080;
        color: white;
        border: none;
      }

      /* Heading box */
      #leave-heading {
        width: auto;  
        height: auto; 
        background-color: white; 
        padding: 10px; 
        border: solid 1px #ccc; 
        text-align: center; 
        display: flex;
        align-items: center;
        justify-content: center;
        border-radius: 10px; 
      }

      /* Container box */
      #leave-sheetData {
        width: 100%;
        margin-top: 20px;
        overflow-x: auto; 
      }

      /* Controls (Buttons and Search) */
      #leave-controls {
        display: flex;
        justify-content: center;
        align-items: center;
        gap: 10px;
        margin-bottom: 10px;
        flex-wrap: wrap;
      }

      /* Search bar container */
      #leave-search-container {
        width: 100%;
        display: flex;
        justify-content: center;
        margin-bottom: 15px;
      }

      /* Search bar styling */
      #leave-search {
        width: 60%;
        padding: 8px 15px;
        border: 1px solid #ccc;
        border-radius: 5px;
        font-size: 14px;
      }

      /* Button styling */
      #leave-controls button {
        width: 120px;
        height: 40px;
        padding: 8px;
        border: none;
        background-color: #2dd413;
        color: white;
        cursor: pointer;
        border-radius: 5px;
        font-size: 14px;
      }

      #leave-controls button:hover {
        background-color: #0056b3;
      }

      /* Table styling */
      #leave-table {
        width: 100%;
        border-collapse: collapse;
        margin-top: 10px;
      }

      #leave-table th, #leave-table td {
        border: 1px solid black;
        padding: 8px;
        text-align: center;
      }

      #leave-table th {
        background-color: #2dd413;
        color: white;
      }

      /* Table container */
      #leave-tableContainer {
        width: 100%;
        overflow-x: auto;
      }

      /* Table action buttons */
      .leave-actionButton {
        width: auto;
        height: 30px;
        padding: 5px 10px;
        border: none;
        cursor: pointer;
        border-radius: 5px;
        font-size: 13px;
      }

      /* Approve Button */
      .leave-approve {
        background-color: #2dd413;
        color: white;
        width: 120px;
        height: 40px;
      }

      .leave-approve:hover {
        background-color: #1aa00e;
      }

      /* Decline Button */
      .leave-decline {
        background-color: #2dd413;
        color: white;
        width: 120px;
        height: 40px;
        margin-top: 1px;
        margin-bottom: 1px;
      }

      .leave-decline:hover {
        background-color: #c9302c;
      }

      /* Comment Button */
      .leave-comment {
        background-color: #2dd413;
        color: white;
        width: 120px;
        height: 40px;
      }

      .leave-comment:hover {
        background-color: #286090;
      }

      /* Delete button */
      .leave-deletebutton {
        background-color: red;
        color: white;
        border: none;
        padding: 5px 10px;
        cursor: pointer;
        border-radius: 5px;
      }

      .leave-deletebutton:hover {
        background-color: darkred;
      }

      /* Error message styling */
      .error-message {
        color: red;
        text-align: center;
        padding: 10px;
        background-color: #ffeeee;
        border: 1px solid #ffcccc;
        border-radius: 5px;
        margin: 10px 0;
      }
    </style>
</head>
<body>
    <div class="main-content">
        <div class="flex justify-between items-center gap-10 header">
          <div class="flex flex-col gap-5 greetings">
            <h1 class="">Leave Management</h1>
          </div>
        </div>

        <!-- Main Container -->
        <div class="rounded-shadow-box">
          <div id="leave-sheetData">
            <!-- Search Bar -->
            <div id="leave-search-container">
              <input type="text" id="leave-search" placeholder="Search..." onkeyup="searchTable()">
            </div>

            <div id="leave-controls">
              <button onclick="fetchSheetData()">View All</button>
              <button onclick="fetchSheetData('Approved')">Approved</button>
              <button onclick="fetchSheetData('Declined')">Declined</button>
              <button onclick="viewForms()">VIEW FORMS</button>
              <button onclick="viewSheet()">VIEW SHEET</button>
            </div>

            <!-- Error display div -->
            <div id="error-container"></div>

            <!-- Table Container -->
            <div id="leave-tableContainer"></div>
          </div>
        </div>
    </div>

    <script src="https://code.jquery.com/jquery-3.6.0.min.js"></script>
    <script>
      let currentStatus = null;
      const defaultCompany = "Company A"; // Set the default company here
      let currentData = []; // Store the current data for searching

      function fetchSheetData(status = null) {
        currentStatus = status;

        $.ajax({
          url: window.location.href, // Send to the same file
          type: 'POST',
          data: {
            action: 'getSheetData',
            company: defaultCompany,
            status: status
          },
          dataType: 'json',
          success: function(response) {
            currentData = response;
            renderTable(response);
          },
          error: function(xhr, status, error) {
            alert("Error: " + error);
          }
        });
      }

      function renderTable(data) {
        let tableContainer = document.getElementById("leave-tableContainer");

        if (!data || data.length === 0) {
          tableContainer.innerHTML = "<p>No data available.</p>";
          return;
        }

        let table = `<table id="leave-table"><tr>`;
        data[0].forEach(cell => table += `<th>${cell}</th>`);
        table += "<th>Actions</th></tr>";

        for (let i = 1; i < data.length; i++) {
          table += "<tr>";
          data[i].forEach(cell => table += `<td>${cell}</td>`);

          if (currentStatus === "Approved" || currentStatus === "Declined") {
            table += `<td><button class="leave-deletebutton" onclick="deleteRow(${i}, '${currentStatus}')">Delete</button></td>`;
          } else {
            table += `<td>
                        <button class="leave-actionButton leave-approve" onclick="approveRow(${i})">Approve</button>
                        <button class="leave-actionButton leave-decline" onclick="declineRow(${i})">Decline</button>
                        <button class="leave-actionButton leave-comment" onclick="addComment(${i})">Comment</button>
                      </td>`;
          }
          table += "</tr>";
        }

        table += "</table>";

        tableContainer.innerHTML = table;
      }

      function searchTable() {
        const searchTerm = document.getElementById("leave-search").value.toLowerCase();

        if (!currentData || currentData.length === 0) {
          return;
        }

        if (searchTerm === "") {
          renderTable(currentData); // Show all data when search is empty
          return;
        }

        // Filter data based on search term
        const headers = currentData[0];
        const filteredData = [headers];

        for (let i = 1; i < currentData.length; i++) {
          const row = currentData[i];
          const rowText = row.join(" ").toLowerCase();

          if (rowText.includes(searchTerm)) {
            filteredData.push(row);
          }
        }

        renderTable(filteredData);
      }

      function approveRow(rowIndex) {
        $.ajax({
          url: window.location.href, // Send to the same file
          type: 'POST',
          data: {
            action: 'approveRow',
            company: defaultCompany,
            rowIndex: rowIndex
          },
          dataType: 'json',
          success: function(response) {
            alert(response.message);
            fetchSheetData();
          },
          error: function(xhr, status, error) {
            alert("Error: " + error);
          }
        });
      }

      function declineRow(rowIndex) {
        $.ajax({
          url: window.location.href, // Send to the same file
          type: 'POST',
          data: {
            action: 'declineRow',
            company: defaultCompany,
            rowIndex: rowIndex
          },
          dataType: 'json',
          success: function(response) {
            alert(response.message);
            fetchSheetData();
          },
          error: function(xhr, status, error) {
            alert("Error: " + error);
          }
        });
      }

      function addComment(rowIndex) {
        let comment = prompt("Enter your comment:");
        if (comment) {
          $.ajax({
            url: window.location.href, // Send to the same file
            type: 'POST',
            data: {
              action: 'saveComment',
              company: defaultCompany,
              rowIndex: rowIndex,
              comment: comment
            },
            dataType: 'json',
            success: function(response) {
              alert(response.message);
            },
            error: function(xhr, status, error) {
              alert("Error: " + error);
            }
          });
        }
      }

      function deleteRow(rowIndex, status) {
        $.ajax({
          url: window.location.href, // Send to the same file
          type: 'POST',
          data: {
            action: 'deleteRow',
            company: defaultCompany,
            rowIndex: rowIndex,
            status: status
          },
          dataType: 'json',
          success: function(response) {
            alert(response.message);
            fetchSheetData(status);
          },
          error: function(xhr, status, error) {
            alert("Error: " + error);
          }
        });
      }

      function viewForms() {
        window.open("https://docs.google.com/forms/d/1n6hObRlvA8nrFC6bhc1-1eo_R3ZY_Mlp1KM_C_uobY8/edit", "_blank");
      }

      function viewSheet() {
        window.open("https://docs.google.com/spreadsheets/d/1wOUaUs4cvu2H9eeGVfr67g0Vs73o5MCGG84bB8MEG8Y/edit?gid=1080764535#gid=1080764535", "_blank");
      }

      // Load data automatically when page loads
      $(document).ready(function() {
        fetchSheetData();
      });
    </script>
</body>
</html>