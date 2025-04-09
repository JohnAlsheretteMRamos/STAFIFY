<?php
// Configuration
require_once 'vendor/autoload.php'; // Include the Google API PHP Client Library

// Set up the Google Sheets API with OAuth2
$client = new Google_Client();
$client->setApplicationName('Leave Management System');
$client->setScopes([
    Google_Service_Sheets::SPREADSHEETS, // Full access is needed for appending/editing
]);

// Add token-based authentication
$tokenPath = 'token.json';

// First, try to use service account if available
$keyFilePath = 'credentials.json'; // Download this from Google Cloud Console
if (file_exists($keyFilePath)) {
    $client->setAuthConfig($keyFilePath);
    error_log("Service account credentials loaded successfully");
} else {
    error_log("Service account credentials file not found at path: " . $keyFilePath);
    
    // If service account not found, try OAuth token
    if (file_exists($tokenPath)) {
        $accessToken = json_decode(file_get_contents($tokenPath), true);
        $client->setAccessToken($accessToken);
        error_log("OAuth token loaded successfully");
    }
    
    // If token exists but is expired
    if ($client->isAccessTokenExpired()) {
        // If we have a refresh token, use it to get a new access token
        if ($client->getRefreshToken()) {
            error_log("Refreshing expired token");
            $client->fetchAccessTokenWithRefreshToken($client->getRefreshToken());
            file_put_contents($tokenPath, json_encode($client->getAccessToken()));
        } else {
            // If no valid authentication method is available, provide instructions
            error_log("No valid authentication method found");
            if (!isset($_GET['code'])) {
                // Redirect for OAuth authorization if this is not a callback
                if (!isset($_GET['action'])) { // Only redirect if this is a page load, not an AJAX call
                    // Set up OAuth
                    $client->setRedirectUri('http://' . $_SERVER['HTTP_HOST'] . $_SERVER['PHP_SELF']);
                    $client->setAccessType('offline');
                    $client->setPrompt('consent'); // Force to choose account and generate refresh token
                    
                    // Generate authorization URL
                    $authUrl = $client->createAuthUrl();
                    header('Location: ' . filter_var($authUrl, FILTER_SANITIZE_URL));
                    exit;
                }
            } else {
                // Handle OAuth callback
                $authCode = $_GET['code'];
                $accessToken = $client->fetchAccessTokenWithAuthCode($authCode);
                $client->setAccessToken($accessToken);
                
                // Save the token to file
                if (!isset($accessToken['error'])) {
                    file_put_contents($tokenPath, json_encode($accessToken));
                    header('Location: ' . $_SERVER['PHP_SELF']);
                    exit;
                }
            }
        }
    }
}

// Set access type to offline to get refresh token
$client->setAccessType('offline');

// Initialize services
$service = new Google_Service_Sheets($client);

// Rest of your code remains the same...

// Define company spreadsheet IDs
$companySheets = [
    "Company A" => "1wOUaUs4cvu2H9eeGVfr67g0Vs73o5MCGG84bB8MEG8Y",
    "Company B" => ""
];

// Default company
$defaultCompany = "Company A";

// Initialize services
$service = new Google_Service_Sheets($client);

/**
 * Get user ID by email from registered data
 */
function getUserIdByEmail($email, $registeredData) {
    foreach ($registeredData as $index => $row) {
        if ($index === 0) continue; // Skip header row
        if ($row[4] === $email) { // Column E - Email
            return $row[0]; // Column A - User ID
        }
    }
    return null; // If not found
}

/**
 * Get list of available companies
 */
function getCompanyList() {
    global $companySheets;
    return array_keys($companySheets);
}

/**
 * Get sheet data from a company's spreadsheet
 */
function getSheetData($companyName, $status = null) {
    global $companySheets, $service;
    
    $sheetId = $companySheets[$companyName] ?? null;
    if (!$sheetId) return [["Error", "Invalid company selected"]];
    
    try {
        if ($status) {
            // Get data from a specific sheet (Approved, Declined, etc.)
            $response = $service->spreadsheets_values->get($sheetId, $status);
        } else {
            // Get data from the first sheet
            $spreadsheet = $service->spreadsheets->get($sheetId);
            $firstSheetTitle = $spreadsheet->getSheets()[0]->properties->title;
            $response = $service->spreadsheets_values->get($sheetId, $firstSheetTitle);
        }
        
        $data = $response->getValues();
        if (!$data || count($data) === 0) {
            return [["Error", "No data available"]];
        }
        
        // Format date values
        foreach ($data as &$row) {
            foreach ($row as &$cell) {
                // Check if cell contains a date string and format it (simplified for PHP)
                if (preg_match('/^\d{4}-\d{2}-\d{2}/', $cell)) {
                    $date = new DateTime($cell);
                    $cell = $date->format('m/d/Y');
                }
            }
        }
        
        return $data;
    } catch (Exception $e) {
        return [["Error", "Could not load sheet: " . $e->getMessage()]];
    }
}

/**
 * Approve a leave request and move it to the Approved sheet
 */
function approveRow($companyName, $rowIndex) {
    global $companySheets, $service;
    
    $sheetId = $companySheets[$companyName] ?? null;
    if (!$sheetId) return "Invalid company selected";
    
    try {
        // Get the spreadsheet structure
        $spreadsheet = $service->spreadsheets->get($sheetId);
        $sheets = $spreadsheet->getSheets();
        $mainSheetTitle = $sheets[0]->properties->title;
        
        // Check if Approved sheet exists, create it if not
        $approvedSheetExists = false;
        foreach ($sheets as $sheet) {
            if ($sheet->properties->title === "Approved") {
                $approvedSheetExists = true;
                break;
            }
        }
        
        if (!$approvedSheetExists) {
            // Create Approved sheet
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
        
        // Get the row data from main sheet
        $range = sprintf("%s!%d:%d", $mainSheetTitle, $rowIndex + 1, $rowIndex + 1);
        $response = $service->spreadsheets_values->get($sheetId, $range);
        $rowData = $response->getValues()[0];
        
        // Update status to "Approved"
        $rowData[count($rowData) - 1] = "Approved";
        
        // Append to Approved sheet
        $valueRange = new Google_Service_Sheets_ValueRange();
        $valueRange->setValues([$rowData]);
        $service->spreadsheets_values->append(
            $sheetId,
            'Approved',
            $valueRange,
            ['valueInputOption' => 'RAW']
        );
        
        // Delete from main sheet
        $deleteRequest = [
            new Google_Service_Sheets_Request([
                'deleteDimension' => [
                    'range' => [
                        'sheetId' => $sheets[0]->properties->sheetId,
                        'dimension' => 'ROWS',
                        'startIndex' => $rowIndex,
                        'endIndex' => $rowIndex + 1
                    ]
                ]
            ])
        ];
        
        $batchUpdateRequest = new Google_Service_Sheets_BatchUpdateSpreadsheetRequest([
            'requests' => $deleteRequest
        ]);
        
        $service->spreadsheets->batchUpdate($sheetId, $batchUpdateRequest);
        
        return "Row approved and moved successfully";
    } catch (Exception $e) {
        return "Error: " . $e->getMessage();
    }
}

/**
 * Decline a leave request and move it to the Declined sheet
 */
function declineRow($companyName, $rowIndex) {
    return processRow($companyName, $rowIndex, "Declined");
}

/**
 * Process a row by moving it to the target sheet
 */
function processRow($companyName, $rowIndex, $status) {
    global $companySheets, $service;
    
    $sheetId = $companySheets[$companyName] ?? null;
    if (!$sheetId) return "Invalid company selected";
    
    try {
        // Get the spreadsheet structure
        $spreadsheet = $service->spreadsheets->get($sheetId);
        $sheets = $spreadsheet->getSheets();
        $mainSheetTitle = $sheets[0]->properties->title;
        
        // Check if target sheet exists, create it if not
        $targetSheetExists = false;
        $targetSheetId = null;
        
        foreach ($sheets as $sheet) {
            if ($sheet->properties->title === $status) {
                $targetSheetExists = true;
                $targetSheetId = $sheet->properties->sheetId;
                break;
            }
        }
        
        if (!$targetSheetExists) {
            // Create target sheet
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
            
            $response = $service->spreadsheets->batchUpdate($sheetId, $batchUpdateRequest);
            $targetSheetId = $response->getReplies()[0]->getAddSheet()->getProperties()->sheetId;
        }
        
        // Get the row data
        $range = sprintf("%s!%d:%d", $mainSheetTitle, $rowIndex + 1, $rowIndex + 1);
        $response = $service->spreadsheets_values->get($sheetId, $range);
        $rowData = $response->getValues()[0];
        
        // Update status
        $rowData[count($rowData) - 1] = $status;
        
        // Append to target sheet
        $valueRange = new Google_Service_Sheets_ValueRange();
        $valueRange->setValues([$rowData]);
        $service->spreadsheets_values->append(
            $sheetId,
            $status,
            $valueRange,
            ['valueInputOption' => 'RAW']
        );
        
        // Delete from main sheet
        $deleteRequest = [
            new Google_Service_Sheets_Request([
                'deleteDimension' => [
                    'range' => [
                        'sheetId' => $sheets[0]->properties->sheetId,
                        'dimension' => 'ROWS',
                        'startIndex' => $rowIndex,
                        'endIndex' => $rowIndex + 1
                    ]
                ]
            ])
        ];
        
        $batchUpdateRequest = new Google_Service_Sheets_BatchUpdateSpreadsheetRequest([
            'requests' => $deleteRequest
        ]);
        
        $service->spreadsheets->batchUpdate($sheetId, $batchUpdateRequest);
        
        return "Row $status successfully";
    } catch (Exception $e) {
        return "Error: " . $e->getMessage();
    }
}

/**
 * Delete a row and move it to the Deleted sheet
 */
function deleteRow($companyName, $rowIndex, $status) {
    global $companySheets, $service;
    
    $sheetId = $companySheets[$companyName] ?? null;
    if (!$sheetId) return "Invalid company selected";
    
    try {
        // Get the spreadsheet structure
        $spreadsheet = $service->spreadsheets->get($sheetId);
        $sheets = $spreadsheet->getSheets();
        
        // Find the sheet ID for the status sheet
        $sourceSheetId = null;
        $sourceSheetTitle = null;
        
        foreach ($sheets as $sheet) {
            if ($sheet->properties->title === $status) {
                $sourceSheetId = $sheet->properties->sheetId;
                $sourceSheetTitle = $sheet->properties->title;
                break;
            }
        }
        
        if (!$sourceSheetId) {
            return "Sheet '$status' not found";
        }
        
        // Check if Deleted sheet exists, create it if not
        $deletedSheetExists = false;
        foreach ($sheets as $sheet) {
            if ($sheet->properties->title === "Deleted") {
                $deletedSheetExists = true;
                break;
            }
        }
        
        if (!$deletedSheetExists) {
            // Create Deleted sheet
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
            
            $service->spreadsheets->batchUpdate($sheetId, $batchUpdateRequest);
        }
        
        // Get the row data
        $range = sprintf("%s!%d:%d", $sourceSheetTitle, $rowIndex + 1, $rowIndex + 1);
        $response = $service->spreadsheets_values->get($sheetId, $range);
        $rowData = $response->getValues()[0];
        
        // Append to Deleted sheet
        $valueRange = new Google_Service_Sheets_ValueRange();
        $valueRange->setValues([$rowData]);
        $service->spreadsheets_values->append(
            $sheetId,
            'Deleted',
            $valueRange,
            ['valueInputOption' => 'RAW']
        );
        
        // Delete from source sheet
        $deleteRequest = [
            new Google_Service_Sheets_Request([
                'deleteDimension' => [
                    'range' => [
                        'sheetId' => $sourceSheetId,
                        'dimension' => 'ROWS',
                        'startIndex' => $rowIndex,
                        'endIndex' => $rowIndex + 1
                    ]
                ]
            ])
        ];
        
        $batchUpdateRequest = new Google_Service_Sheets_BatchUpdateSpreadsheetRequest([
            'requests' => $deleteRequest
        ]);
        
        $service->spreadsheets->batchUpdate($sheetId, $batchUpdateRequest);
        
        return "Row moved to Deleted";
    } catch (Exception $e) {
        return "Error: " . $e->getMessage();
    }
}

/**
 * Save a comment in the Comment sheet
 */
function saveComment($companyName, $rowIndex, $comment) {
    global $companySheets, $service;
    
    $sheetId = $companySheets[$companyName] ?? null;
    if (!$sheetId) return "Invalid company selected";
    
    try {
        // Get the spreadsheet structure
        $spreadsheet = $service->spreadsheets->get($sheetId);
        $sheets = $spreadsheet->getSheets();
        $mainSheetTitle = $sheets[0]->properties->title;
        
        // Check if Comment sheet exists, create it if not
        $commentSheetExists = false;
        foreach ($sheets as $sheet) {
            if ($sheet->properties->title === "Comment") {
                $commentSheetExists = true;
                break;
            }
        }
        
        if (!$commentSheetExists) {
            // Create Comment sheet
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
        
        // Get the row data
        $range = sprintf("%s!%d:%d", $mainSheetTitle, $rowIndex + 1, $rowIndex + 1);
        $response = $service->spreadsheets_values->get($sheetId, $range);
        $rowData = $response->getValues()[0];
        
        $name = $rowData[4]; // Column E - Name
        $email = $rowData[1]; // Column B - Email
        
        // Append comment to Comment sheet
        $valueRange = new Google_Service_Sheets_ValueRange();
        $valueRange->setValues([[$name, $email, $comment]]);
        $service->spreadsheets_values->append(
            $sheetId,
            'Comment',
            $valueRange,
            ['valueInputOption' => 'RAW']
        );
        
        return "Comment saved successfully!";
    } catch (Exception $e) {
        return "Error: " . $e->getMessage();
    }
}

// Handle AJAX requests
if (isset($_GET['action'])) {
    $action = $_GET['action'];
    $response = [];
    
    switch ($action) {
        case 'getCompanyList':
            $response = getCompanyList();
            break;
            
        case 'getSheetData':
            $company = $_GET['company'] ?? $defaultCompany;
            $status = $_GET['status'] ?? null;
            $response = getSheetData($company, $status);
            break;
            
        case 'approveRow':
            $company = $_GET['company'] ?? $defaultCompany;
            $rowIndex = (int)$_GET['rowIndex'];
            $response = approveRow($company, $rowIndex);
            break;
            
        case 'declineRow':
            $company = $_GET['company'] ?? $defaultCompany;
            $rowIndex = (int)$_GET['rowIndex'];
            $response = declineRow($company, $rowIndex);
            break;
            
        case 'deleteRow':
            $company = $_GET['company'] ?? $defaultCompany;
            $rowIndex = (int)$_GET['rowIndex'];
            $status = $_GET['status'];
            $response = deleteRow($company, $rowIndex, $status);
            break;
            
        case 'saveComment':
            $company = $_GET['company'] ?? $defaultCompany;
            $rowIndex = (int)$_GET['rowIndex'];
            $comment = $_GET['comment'];
            $response = saveComment($company, $rowIndex, $comment);
            break;
    }
    
    header('Content-Type: application/json');
    echo json_encode($response);
    exit;
}
?>

<!DOCTYPE html>
<html>
<head>
    <meta name="viewport" content="width=device-width, initial-scale=1" />
    <title>Leave Management</title>
    <link rel="stylesheet" href="https://cdnjs.cloudflare.com/ajax/libs/animate.css/4.1.1/animate.min.css"/>
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

                <!-- Table Container -->
                <div id="leave-tableContainer"></div>
            </div>
        </div>
    </div>

    <script>
        let currentStatus = null;
        const defaultCompany = "Company A"; // Set default company here
        let currentData = []; // Store current data for searching

        function fetchSheetData(status = null) {
            currentStatus = status;
            const url = `?action=getSheetData&company=${defaultCompany}${status ? `&status=${status}` : ''}`;
            
            fetch(url)
                .then(response => response.json())
                .then(data => {
                    currentData = data;
                    renderTable(data);
                })
                .catch(error => {
                    console.error('Error:', error);
                    document.getElementById("leave-tableContainer").innerHTML = "<p>Error loading data.</p>";
                });
        }

        function renderTable(data) {
            let tableContainer = document.getElementById("leave-tableContainer");

            if (!data || data.length === 0 || (data.length === 1 && data[0][0] === "Error")) {
                tableContainer.innerHTML = "<p>No data available or " + (data[0][1] || "an error occurred") + ".</p>";
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
            fetch(`?action=approveRow&company=${defaultCompany}&rowIndex=${rowIndex}`)
                .then(response => response.json())
                .then(data => {
                    alert(data);
                    fetchSheetData();
                })
                .catch(error => {
                    console.error('Error:', error);
                    alert("An error occurred. Please try again.");
                });
        }

        function declineRow(rowIndex) {
            fetch(`?action=declineRow&company=${defaultCompany}&rowIndex=${rowIndex}`)
                .then(response => response.json())
                .then(data => {
                    alert(data);
                    fetchSheetData();
                })
                .catch(error => {
                    console.error('Error:', error);
                    alert("An error occurred. Please try again.");
                });
        }

        function addComment(rowIndex) {
            let comment = prompt("Enter your comment:");
            if (comment) {
                fetch(`?action=saveComment&company=${defaultCompany}&rowIndex=${rowIndex}&comment=${encodeURIComponent(comment)}`)
                    .then(response => response.json())
                    .then(data => {
                        alert(data);
                    })
                    .catch(error => {
                        console.error('Error:', error);
                        alert("An error occurred. Please try again.");
                    });
            }
        }

        function deleteRow(rowIndex, status) {
            fetch(`?action=deleteRow&company=${defaultCompany}&rowIndex=${rowIndex}&status=${status}`)
                .then(response => response.json())
                .then(data => {
                    alert(data);
                    fetchSheetData(status);
                })
                .catch(error => {
                    console.error('Error:', error);
                    alert("An error occurred. Please try again.");
                });
        }

        function viewForms() {
            window.open("https://docs.google.com/forms/d/1n6hObRlvA8nrFC6bhc1-1eo_R3ZY_Mlp1KM_C_uobY8/edit", "_blank");
        }

        function viewSheet() {
            window.open("https://docs.google.com/spreadsheets/d/1wOUaUs4cvu2H9eeGVfr67g0Vs73o5MCGG84bB8MEG8Y/edit?gid=1080764535#gid=1080764535", "_blank");
        }

        // Load data automatically when page loads
        window.onload = function() {
            fetchSheetData();
        };
    </script>
</body>
</html>