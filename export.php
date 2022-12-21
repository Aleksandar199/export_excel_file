<?php 
session_name("budget_tracker");
session_start();
error_reporting(0);
include 'connect.php';
include_once "../vendor/autoload.php";




    $sql1 = "SELECT * FROM budget_db.sp ORDER BY po_date DESC";      

    $query1 = mysqli_query($conn, $sql1);



use PhpOffice\PhpSpreadsheet\Spreadsheet;
use PhpOffice\PhpSpreadsheet\Writer\Xlsx;

$spreadsheet = new Spreadsheet();
$spreadsheet->getActiveSheet()->getDefaultColumnDimension()->setWidth(22);
$styleArray = [
    'font' => [
        'bold' => true,
    ],
    'alignment' => [
        'horizontal' => \PhpOffice\PhpSpreadsheet\Style\Alignment::HORIZONTAL_CENTER,
    ],
    'borders' => [
        'outline' => [
            'borderStyle' => \PhpOffice\PhpSpreadsheet\Style\Border::BORDER_THIN,
        ],
    ],
    'fill' => [
        'fillType' => \PhpOffice\PhpSpreadsheet\Style\Fill::FILL_GRADIENT_LINEAR,
        'rotation' => 90,
        'startColor' => [
            'argb' => 'FF87CEEB',
        ],
        'endColor' => [
            'argb' => 'F5FFFA',   
        ],
    ],
];
$styleArray2 = [
    'font' => [
        'bold' => true,
    ],
    'alignment' => [
        'horizontal' => \PhpOffice\PhpSpreadsheet\Style\Alignment::HORIZONTAL_CENTER,
    ],
];
$spreadsheet->getActiveSheet()->getStyle('A1:L1')->applyFromArray($styleArray);
$sheet = $spreadsheet->getActiveSheet();
$sheet->setCellValue('A1', 'Vendor');
$sheet->setCellValue('B1', 'Description');
$sheet->setCellValue('C1', 'Comment');
$sheet->setCellValue('D1', 'PO');
$sheet->setCellValue('E1', 'Accrual');
$sheet->setCellValue('F1', 'GR');
$sheet->setCellValue('G1', 'RSD');
$sheet->setCellValue('H1', 'EUR');
$sheet->setCellValue('I1', 'PO date');
$sheet->setCellValue('J1', 'GR date');
$sheet->setCellValue('K1', 'Plant');
$sheet->setCellValue('L1', 'Budget month');
$i = 2;


while($res = mysqli_fetch_assoc($query1)){


    $vendor = $res['vendor'];
    $description = $res['description'];
    $comment = $res['comment'];
    $po = $res['po'];
    // $accrual = $res['accrual'];
    $gr = $res['gr'];
    $rsd = $res['rsd'];
    $eur = $res['eur'];
    $po_date = $res['po_date'];
    // $gr_date = $res['gr_date'];
    // $days = $res['days'];
    $plant = $res['plant'];
    $budget_month = $res['budget_month'];
    // $budget = $res['budget'];


    $spreadsheet->getActiveSheet()->getStyle('A'.$i.':L'.$i)->applyFromArray($styleArray2);
    $sheet = $spreadsheet->getActiveSheet();
    $sheet->setCellValue('A'.$i, $vendor);
    $sheet->setCellValue('B'.$i, $description);
    $sheet->setCellValue('C'.$i, $comment);
    $sheet->setCellValue('D'.$i, $po);
    $sheet->setCellValue('E'.$i, $accrual);
    $sheet->setCellValue('F'.$i, $gr);
    $sheet->setCellValue('G'.$i, $rsd);
    $sheet->setCellValue('H'.$i, $eur);
    $sheet->setCellValue('I'.$i, $po_date);
    $sheet->setCellValue('J'.$i, $gr_date);
    $sheet->setCellValue('K'.$i, $plant);
    $sheet->setCellValue('L'.$i, $budget_month);
    $i++;

}

$date = date('d-m-Y');
$file_name = "Spare_parts_".$date.".xlsx";
//SAVE FILE TO EXCEL
$writer = new Xlsx($spreadsheet);
$writer->save($file_name);


//DOWNLOAD FILE
       $f = $file_name;   

       $file = ("$f");

       $filetype=filetype($file);

       $filename=basename($file);

       header ("Content-Type: ".$filetype);

       header ("Content-Length: ".filesize($file));

       header ("Content-Disposition: attachment; filename=".$filename);

       readfile($file);

       //DELETE FILE FROM SERVER AFTER FILE HAS BEEN DOWNLOADED FROM USER
       unlink($file_name);      


       //REDIRECT BACK FROM SCRIPT 
       header("Location: database_sp.php");
?>
