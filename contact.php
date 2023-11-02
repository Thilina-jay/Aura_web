require('PHPExcel.php');

if ($_SERVER["REQUEST_METHOD"] == "POST") {
    $name = $_POST["name"];
    $email = $_POST["email"];
    $message = $_POST["message"];
    
    // Create a new PHPExcel object
    $objPHPExcel = new PHPExcel();

    // Open the Excel file
    $inputFileName = 'D:\stuff\chamo\Counter\New folder\weba.xlsx'; // Replace with your file path
    $objPHPExcel = PHPExcel_IOFactory::load($inputFileName);

    // Get the active sheet
    $worksheet = $objPHPExcel->getActiveSheet();

    // Find the next empty row in the sheet
    $highestRow = $worksheet->getHighestRow();
    $row = $highestRow + 1;

    // Write data to the Excel sheet
    $worksheet->setCellValueByColumnAndRow(0, $row, $name);
    $worksheet->setCellValueByColumnAndRow(1, $row, $email);
    $worksheet->setCellValueByColumnAndRow(2, $row, $message);

    // Save the Excel file
    $writer = PHPExcel_IOFactory::createWriter($objPHPExcel, 'Excel2007');
    $writer->save($inputFileName);

    // ...
}
