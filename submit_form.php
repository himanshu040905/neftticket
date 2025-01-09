<?php
require 'vendor/autoload.php';  // Include PhpSpreadsheet library

use PhpOffice\PhpSpreadsheet\Spreadsheet;
use PhpOffice\PhpSpreadsheet\Writer\Xlsx;

if ($_SERVER['REQUEST_METHOD'] == 'POST') {
    // Collect form data
    $name = $_POST['name'];
    $email = $_POST['email'];
    $message = $_POST['message'];

    // Create a new Spreadsheet object
    $spreadsheet = new Spreadsheet();
    $sheet = $spreadsheet->getActiveSheet();

    // Check if the file already exists
    $fileName = 'contact_form_data.numbers';
    if (file_exists($fileName)) {
        // Load the existing file
        $spreadsheet = \PhpOffice\PhpSpreadsheet\IOFactory::load($fileName);
        $sheet = $spreadsheet->getActiveSheet();
        $lastRow = $sheet->getHighestRow() + 1;  // Find the next available row
    } else {
        // Create headers if the file doesn't exist
        $sheet->setCellValue('A1', 'Name');
        $sheet->setCellValue('B1', 'Email');
        $sheet->setCellValue('C1', 'Message');
        $lastRow = 2;
    }

    // Insert form data into the next row
    $sheet->setCellValue('A' . $lastRow, $name);
    $sheet->setCellValue('B' . $lastRow, $email);
    $sheet->setCellValue('C' . $lastRow, $message);

    // Save the spreadsheet
    $writer = new Xlsx($spreadsheet);
    $writer->save($fileName);

    // Redirect or give success message
    echo "Thank you for your message!";
    // Optionally, you could redirect back to the contact page with a success message
}
?>
