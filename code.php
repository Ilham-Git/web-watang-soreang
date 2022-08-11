<?php
session_start();
include 'inc/koneksi.php';
require 'vendor/autoload.php';

use PhpOffice\PhpSpreadsheet\Spreadsheet;
use PhpOffice\PhpSpreadsheet\Writer\Xlsx;

if (isset($_POST['save_data'])) {
    $filename = $_FILES['import_file']['name'];
    $file_ext = pathinfo($filename, PATHINFO_EXTENSION);

    $allowed_ext = ['xls', 'csv', 'xlsx'];

    if (in_array($file_ext, $allowed_ext)) {
        $inputFileName = $_FILES['import_file']['tmp_name'];
        /** Load $inputFileName to a Spreadsheet object **/
        $spreadsheet = \PhpOffice\PhpSpreadsheet\IOFactory::load($inputFileName);
        $data = $spreadsheet->getActiveSheet()->toArray();

        $count = '0';

        foreach ($data as $row) {
            if ($count > 0) {
                $nik = '-';
                $nama = $row['1'];
                $tempat_lh = $row['2'];
                $tgl_lh = $row['3'];
                $jekel = $row['4'];
                $agama = $row['6'];
                $pekerjaan = $row['8'];
                $kawin = $row['9'];
                $alamat = $row['11'];
                $rt = $row['12'];
                $rw = $row['13'];
                $status = "Ada";
                $stunting = "Tidak";

                $uploadQuery = "INSERT INTO tb_pdd (nik,nama,tempat_lh,tgl_lh,jekel,desa,rt,rw,agama,kawin,pekerjaan,status,stunting) VALUES (
                '$nik','$nama','$tempat_lh','$tgl_lh','$jekel','$alamat','$rt','$rw','$agama','$kawin','$pekerjaan','$status','$stunting')";
                $result = mysqli_query($koneksi, $uploadQuery);
                $msg = true;
            } else {
                $count = '1';
            }
        }

        if (isset($msg)) {
            $_SESSION['message'] = "Upload Data Sukses";
            header('Location: index.php?page=upload-data');
            exit(0);
        } else {
            $_SESSION['message'] = "Upload Data Gagal";
            header('Location: index.php?page=upload-data');
            exit(0);
        }
    } else {
        $_SESSION['message'] = "Invalid File";
        header('Location: index.php?page=upload-data');
        exit(0);
    }
}
