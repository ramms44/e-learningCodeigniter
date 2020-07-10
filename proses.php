<?php
include('koneksi.php');
require 'vendor/autoload.php';
 
use PhpOffice\PhpSpreadsheet\Spreadsheet;
use PhpOffice\PhpSpreadsheet\Reader\Csv;
use PhpOffice\PhpSpreadsheet\Reader\Xlsx;
 
$file_mimes = array('application/octet-stream', 'application/vnd.ms-excel', 'application/x-csv', 'text/x-csv', 'text/csv', 'application/csv', 'application/excel', 'application/vnd.msexcel', 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');

if(isset($_FILES['berkas_excel']['name']) && in_array($_FILES['berkas_excel']['type'], $file_mimes)) {
 
    $arr_file = explode('.', $_FILES['berkas_excel']['name']);
    $extension = end($arr_file);
 
    if('csv' == $extension) {
        $reader = new \PhpOffice\PhpSpreadsheet\Reader\Csv();
    } else {
        $reader = new \PhpOffice\PhpSpreadsheet\Reader\Xlsx();
    }
 
    $spreadsheet = $reader->load($_FILES['berkas_excel']['tmp_name']);
     
    $sheetData = $spreadsheet->getActiveSheet()->toArray();
	for($i = 1;$i < count($sheetData);$i++)
	{
        $nama = $sheetData[$i]['1'];
        $kelas = $sheetData[$i]['2'];
        $alamat = $sheetData[$i]['3'];
        $alamat = $sheetData[$i]['4'];
        $alamat = $sheetData[$i]['5'];
        $alamat = $sheetData[$i]['6'];
        $alamat = $sheetData[$i]['7'];
        $alamat = $sheetData[$i]['8'];
        $alamat = $sheetData[$i]['9'];
        $alamat = $sheetData[$i]['10'];
        $alamat = $sheetData[$i]['11'];
        mysqli_query($koneksi,"insert into siswa (id, nis, nama, jenis_kelamin, tempat_lahir, tgl_lahir, agama, alamat, tahun_masuk, foto, status_id) values ('','$nama','$kelas','$alamat','$alamat','$alamat','$alamat','$alamat','$alamat','$alamat','$alamat')");
    }
    header("Location: form_upload.html"); 
}
?>