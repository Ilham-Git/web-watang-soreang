<?php
include_once('excel_reader2.php');
include_once('inc/koneksi.php');

$target = basename($_FILES['data_karyawan']['name']);
move_uploaded_file($_FILES['data_karyawan']['tmp_name'], $target);

// permision agar file bisa terbaca
chmod($_FILES['data_karyawan']['name'], 0777);

// mengambil isi file xls
$data = new Spreadsheet_Excel_Reader($_FILES['data_karyawan']['name'], false);

// hitung jumlah baris
$jumlah_baris = $data->rowcount($sheet_index = 0);
$success = 0;
for ($i = 2; $i <= $jumlah_baris; $i++) {
    $nik = $data->val($i, 1);
    $nama = $data->val($i, 1);
    $password = $data->val($i, 1);
    $dept = $data->val($i, 1);
    $jabatan = $data->val($i, 1);
    $akses = $data->val($i, 1);

    if ($nik != "" && $nama != "") {
        $enc_password = md5($password);
        mysqli_query($connect, "INSERT INTO tb_karyawan VALUES('',
        '$nik','$nama','$enc_password','$dept','$jabatan',
        '$akses')");
        $success++;
    }
    unlink($_FILES['data_karyawan']['name']);

    if ($success > 0) {
        header("location:index.php?upload=success");
    } else {
        header("location:index.php?upload=gagal");
    }
}
