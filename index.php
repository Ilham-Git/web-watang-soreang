<?php
//Mulai Sesion
session_start();
if (isset($_SESSION["ses_username"]) == "") {
	header("location: login.php");
} else {
	$data_id = $_SESSION["ses_id"];
	$data_nama = $_SESSION["ses_nama"];
	$data_user = $_SESSION["ses_username"];
	$data_level = $_SESSION["ses_level"];
}

//KONEKSI DB
include "inc/koneksi.php";
?>

<!DOCTYPE html>
<html>

<head>
	<meta charset="utf-8">
	<meta http-equiv="X-UA-Compatible" content="IE=edge">
	<title>SINTING</title>
	<link rel="icon" href="dist/img/LogoWeb.png">
	<!-- Tell the browser to be responsive to screen width -->
	<meta name="viewport" content="width=device-width, initial-scale=1">

	<!-- Font Awesome -->
	<link rel="stylesheet" href="plugins/fontawesome-free/css/all.min.css">
	<!-- Ionicons -->
	<link rel="stylesheet" href="https://code.ionicframework.com/ionicons/2.0.1/css/ionicons.min.css">
	<!-- DataTables -->
	<link rel="stylesheet" href="plugins/datatables-bs4/css/dataTables.bootstrap4.css">
	<!-- overlayScrollbars -->
	<link rel="stylesheet" href="dist/css/adminlte.min.css">
	<!-- Select2 -->
	<link rel="stylesheet" href="plugins/select2/css/select2.min.css">
	<link rel="stylesheet" href="plugins/select2-bootstrap4-theme/select2-bootstrap4.min.css">
	<!-- Google Font: Source Sans Pro -->
	<link href="https://fonts.googleapis.com/css?family=Source+Sans+Pro:300,400,400i,700" rel="stylesheet">
	<!-- Alert -->
	<script src="plugins/alert.js"></script>
</head>

<body class="hold-transition sidebar-mini">
	<!-- Site wrapper -->
	<div class="wrapper">
		<!-- Navbar -->
		<nav class="main-header navbar navbar-expand navbar-green navbar-light">
			<!-- Left navbar links -->
			<ul class="navbar-nav">
				<li class="nav-item">
					<a class="nav-link" data-widget="pushmenu" href="#">
						<i class="fas fa-bars text-white"></i>
					</a>
				</li>

			</ul>

			<!-- SEARCH FORM -->
			<ul class="navbar-nav ml-auto">

				<li class="nav-item d-none d-sm-inline-block">
					<a href="index.php" class="nav-link">
						<font color="white">
							<b>SISTEM INFORMASI STUNTING WATANG SOREANG</b>
						</font>
					</a>
				</li>
			</ul>

		</nav>
		<!-- /.navbar -->

		<!-- Main Sidebar Container -->
		<aside class="main-sidebar sidebar-dark-primary elevation-4">
			<!-- Brand Logo -->
			<a href="index.php" class="brand-link">
				<img src="dist/img/LogoWeb.png" alt="AdminLTE Logo" class="brand-image" style="opacity: .8">
				<span class="brand-text"> SINTING</span>
			</a>

			<!-- Sidebar -->
			<div class="sidebar">
				<!-- Sidebar user (optional) -->
				<div class="user-panel mt-2 pb-2 mb-2 d-flex">
					<div class="image">
						<img src="dist/img/admin.ico">
					</div>

					<div class="info">
						<a href="index.php" class="d-block">
							<?php echo $data_nama; ?>
						</a>
						<span class="badge badge-success">
							<?php echo $data_level; ?>
						</span>
					</div>
				</div>

				<!-- Sidebar Menu -->
				<nav class="mt-2">
					<ul class="nav nav-pills nav-sidebar flex-column" data-widget="treeview" role="menu" data-accordion="false">

						<!-- Level  -->
						<?php
						if ($data_level == "Administrator") {
						?>
							<li class="nav-item">
								<a href="index.php" class="nav-link">
									<i class="nav-icon fas fa-tachometer-alt"></i>
									<p>
										Dashboard
									</p>
								</a>
							</li>

							<li class="nav-item has-treeview">
								<a href="#" class="nav-link">
									<i class="nav-icon fas fa-table"></i>
									<p>
										Kelola Data
										<i class="fas fa-angle-left right"></i>
									</p>
								</a>
								<ul class="nav nav-treeview">
									<li class="nav-item">
										<a href="?page=data-pend" class="nav-link">
											<i class="nav-icon far fa-circle text-warning"></i>
											<p>Data Penduduk</p>
										</a>
									</li>
									<li class="nav-item">
										<a href="?page=data-kartu" class="nav-link">
											<i class="nav-icon far fa-circle text-warning"></i>
											<p>Data Kartu Keluarga</p>
										</a>
									</li>
								</ul>
							</li>

							<li class="nav-item has-treeview">
								<a href="#" class="nav-link">
									<i class="nav-icon fas fa-stethoscope"></i>
									<p>
										Data Stunting
										<i class="fas fa-angle-left right"></i>
									</p>
								</a>
								<ul class="nav nav-treeview">
									<li class="nav-item">
										<a href="?page=data-stunting" class="nav-link">
											<i class="nav-icon far fa-circle text-warning"></i>
											<p>Data Stunting</p>
										</a>
									</li>
									<li class="nav-item">
										<a href="?page=data-penyintas" class="nav-link">
											<i class="nav-icon far fa-circle text-warning"></i>
											<p>Data Penyintas</p>
										</a>
									</li>
									<li class="nav-item">
										<a href="?page=data-berisiko" class="nav-link">
											<i class="nav-icon far fa-circle text-warning"></i>
											<p>Data Beresiko</p>
										</a>
									</li>
								</ul>
							</li>

							<li class="nav-item has-treeview">
								<a href="#" class="nav-link">
									<i class="nav-icon fas fa-cogs"></i>
									<p>
										Sirkulasi Penduduk
										<i class="fas fa-angle-left right"></i>
									</p>
								</a>
								<ul class="nav nav-treeview">
									<li class="nav-item">
										<a href="?page=data-lahir" class="nav-link">
											<i class="nav-icon far fa-circle text-warning"></i>
											<p>Data Lahir</p>
										</a>
									</li>
									<li class="nav-item">
										<a href="?page=data-mendu" class="nav-link">
											<i class="nav-icon far fa-circle text-warning"></i>
											<p>Data Meninggal</p>
										</a>
									</li>

									<li class="nav-item">
										<a href="?page=data-datang" class="nav-link">
											<i class="nav-icon far fa-circle text-warning"></i>
											<p>Data Pendatang</p>
										</a>
									</li>
									<li class="nav-item">
										<a href="?page=data-pindah" class="nav-link">
											<i class="nav-icon far fa-circle text-warning"></i>
											<p>Data Pindah</p>
										</a>
									</li>
								</ul>
							</li>

							<li class="nav-item has-treeview">
								<a href="#" class="nav-link">
									<i class="nav-icon fas fa-file"></i>
									<p>
										Kelola Surat
										<i class="fas fa-angle-left right"></i>
									</p>
								</a>
								<ul class="nav nav-treeview">

									<li class="nav-item">
										<a href="?page=suket-domisili" class="nav-link">
											<i class="nav-icon far fa-circle text-warning"></i>
											<p>Su-Ket Domisili</p>
										</a>
									</li>
									<li class="nav-item">
										<a href="?page=suket-lahir" class="nav-link">
											<i class="nav-icon far fa-circle text-warning"></i>
											<p>Su-Ket Kelahiran</p>
										</a>
									</li>
									<li class="nav-item">
										<a href="?page=suket-mati" class="nav-link">
											<i class="nav-icon far fa-circle text-warning"></i>
											<p>Su-Ket Kematian</p>
										</a>
									</li>
									<li class="nav-item">
										<a href="?page=suket-datang" class="nav-link">
											<i class="nav-icon far fa-circle text-warning"></i>
											<p>Su-Ket Pendatang</p>
										</a>
									</li>
									<li class="nav-item">
										<a href="?page=suket-pindah" class="nav-link">
											<i class="nav-icon far fa-circle text-warning"></i>
											<p>Su-Ket Pindah</p>
										</a>
									</li>
								</ul>
							</li>


							<!-- <li class="nav-item has-treeview">
								<a href="#" class="nav-link">
									<i class="nav-icon fas fa-print"></i>
									<p>
										Kelola Laporan
										<i class="fas fa-angle-left right"></i>
									</p>
								</a>
								<ul class="nav nav-treeview">

									<li class="nav-item">
										<a href="?page=laporan-domisili" class="nav-link">
											<i class="nav-icon far fa-circle text-warning"></i>
											<p>Laporan Penduduk</p>
										</a>
									</li>
									<li class="nav-item">
										<a href="?page=laporan-lahir" class="nav-link">
											<i class="nav-icon far fa-circle text-warning"></i>
											<p>Laporan Lahir</p>
										</a>
									</li>
									<li class="nav-item">
										<a href="?page=laporan-mati" class="nav-link">
											<i class="nav-icon far fa-circle text-warning"></i>
											<p>Laporan Meninggal</p>
										</a>
									</li>
									<li class="nav-item">
										<a href="?page=laporan-datang" class="nav-link">
											<i class="nav-icon far fa-circle text-warning"></i>
											<p>Laporan Pendatang</p>
										</a>
									</li>
									<li class="nav-item">
										<a href="?page=laporan-pindah" class="nav-link">
											<i class="nav-icon far fa-circle text-warning"></i>
											<p>Laporan Pindah</p>
										</a>
									</li>
								</ul>
							</li> -->

							<li class="nav-header">Setting</li>

							<li class="nav-item">
								<a href="?page=data-pengguna" class="nav-link">
									<i class="nav-icon fas fa-user"></i>
									<p>
										Pengguna Sistem
									</p>
								</a>
							</li>
							<li class="nav-item">
								<a href="?page=import-data" class="nav-link">
									<i class="nav-icon fas fa-download"></i>
									<p>
										Import Data
									</p>
								</a>
							</li>
							<li class="nav-item">
								<a href="?page=export-data" class="nav-link">
									<i class="nav-icon fas fa-upload"></i>
									<p>
										Export Data
									</p>
								</a>
							</li>

						<?php
						} elseif ($data_level == "Kaur Pemerintah") {
						?>

							<li class="nav-item">
								<a href="index.php" class="nav-link">
									<i class="nav-icon fas fa-tachometer-alt"></i>
									<p>
										Dashboard
									</p>
								</a>
							</li>

							<li class="nav-item has-treeview">
								<a href="#" class="nav-link">
									<i class="nav-icon fas fa-table"></i>
									<p>
										Kelola Data
										<i class="fas fa-angle-left right"></i>
									</p>
								</a>
								<ul class="nav nav-treeview">
									<li class="nav-item">
										<a href="?page=data-pend" class="nav-link">
											<i class="nav-icon far fa-circle text-warning"></i>
											<p>Data Penduduk</p>
										</a>
									</li>
									<li class="nav-item">
										<a href="?page=data-kartu" class="nav-link">
											<i class="nav-icon far fa-circle text-warning"></i>
											<p>Data Kartu Keluarga</p>
										</a>
									</li>
								</ul>
							</li>

							<li class="nav-item has-treeview">
								<a href="#" class="nav-link">
									<i class="nav-icon fas fa-stethoscope"></i>
									<p>
										Data Stunting
										<i class="fas fa-angle-left right"></i>
									</p>
								</a>
								<ul class="nav nav-treeview">
									<li class="nav-item">
										<a href="?page=data-stunting" class="nav-link">
											<i class="nav-icon far fa-circle text-warning"></i>
											<p>Data Stunting</p>
										</a>
									</li>
									<li class="nav-item">
										<a href="?page=data-penyintas" class="nav-link">
											<i class="nav-icon far fa-circle text-warning"></i>
											<p>Data Penyintas</p>
										</a>
									</li>
									<li class="nav-item">
										<a href="?page=data-berisiko" class="nav-link">
											<i class="nav-icon far fa-circle text-warning"></i>
											<p>Data Beresiko</p>
										</a>
									</li>
								</ul>
							</li>

							<li class="nav-item has-treeview">
								<a href="#" class="nav-link">
									<i class="nav-icon fas fa-cogs"></i>
									<p>
										Sirkulasi Penduduk
										<i class="fas fa-angle-left right"></i>
									</p>
								</a>
								<ul class="nav nav-treeview">
									<li class="nav-item">
										<a href="?page=data-lahir" class="nav-link">
											<i class="nav-icon far fa-circle text-warning"></i>
											<p>Data Lahir</p>
										</a>
									</li>
									<li class="nav-item">
										<a href="?page=data-mendu" class="nav-link">
											<i class="nav-icon far fa-circle text-warning"></i>
											<p>Data Meninggal</p>
										</a>
									</li>

									<li class="nav-item">
										<a href="?page=data-datang" class="nav-link">
											<i class="nav-icon far fa-circle text-warning"></i>
											<p>Data Pendatang</p>
										</a>
									</li>
									<li class="nav-item">
										<a href="?page=data-pindah" class="nav-link">
											<i class="nav-icon far fa-circle text-warning"></i>
											<p>Data Pindah</p>
										</a>
									</li>
								</ul>
							</li>

							<li class="nav-item has-treeview">
								<a href="#" class="nav-link">
									<i class="nav-icon fas fa-file"></i>
									<p>
										Kelola Surat
										<i class="fas fa-angle-left right"></i>
									</p>
								</a>
								<ul class="nav nav-treeview">

									<li class="nav-item">
										<a href="?page=suket-domisili" class="nav-link">
											<i class="nav-icon far fa-circle text-warning"></i>
											<p>Su-Ket Domisili</p>
										</a>
									</li>
									<li class="nav-item">
										<a href="?page=suket-lahir" class="nav-link">
											<i class="nav-icon far fa-circle text-warning"></i>
											<p>Su-Ket Kelahiran</p>
										</a>
									</li>
									<li class="nav-item">
										<a href="?page=suket-mati" class="nav-link">
											<i class="nav-icon far fa-circle text-warning"></i>
											<p>Su-Ket Kematian</p>
										</a>
									</li>
									<li class="nav-item">
										<a href="?page=suket-datang" class="nav-link">
											<i class="nav-icon far fa-circle text-warning"></i>
											<p>Su-Ket Pendatang</p>
										</a>
									</li>
									<li class="nav-item">
										<a href="?page=suket-pindah" class="nav-link">
											<i class="nav-icon far fa-circle text-warning"></i>
											<p>Su-Ket Pindah</p>
										</a>
									</li>
								</ul>
							</li>

							<!-- <li class="nav-item has-treeview">
								<a href="#" class="nav-link">
									<i class="nav-icon fas fa-print"></i>
									<p>
										Kelola Laporan
										<i class="fas fa-angle-left right"></i>
									</p>
								</a>
								<ul class="nav nav-treeview">

									<li class="nav-item">
										<a href="?page=laporan-domisili" class="nav-link">
											<i class="nav-icon far fa-circle text-warning"></i>
											<p>Laporan Penduduk</p>
										</a>
									</li>
									<li class="nav-item">
										<a href="?page=laporan-lahir" class="nav-link">
											<i class="nav-icon far fa-circle text-warning"></i>
											<p>Laporan Lahir</p>
										</a>
									</li>
									<li class="nav-item">
										<a href="?page=laporan-mati" class="nav-link">
											<i class="nav-icon far fa-circle text-warning"></i>
											<p>Laporan Meninggal</p>
										</a>
									</li>
									<li class="nav-item">
										<a href="?page=laporan-datang" class="nav-link">
											<i class="nav-icon far fa-circle text-warning"></i>
											<p>Laporan Pendatang</p>
										</a>
									</li>
									<li class="nav-item">
										<a href="?page=laporan-pindah" class="nav-link">
											<i class="nav-icon far fa-circle text-warning"></i>
											<p>Laporan Pindah</p>
										</a>
									</li>
								</ul>
							</li> -->

							<li class="nav-header">Setting</li>

						<?php
						}
						?>

						<li class="nav-item">
							<a onclick="return confirm('Apakah anda yakin akan keluar ?')" href="logout.php" class="nav-link">
								<i class="nav-icon fas fa-arrow-circle-right"></i>
								<p>
									Logout
								</p>
							</a>
						</li>

				</nav>
				<!-- /.sidebar-menu -->
			</div>
			<!-- /.sidebar -->
		</aside>

		<!-- Content Wrapper. Contains page content -->
		<div class="content-wrapper">
			<!-- Content Header (Page header) -->
			<section class="content-header">
			</section>

			<!-- Main content -->
			<section class="content">
				<!-- /. WEB DINAMIS DISINI ############################################################################### -->
				<div class="container-fluid">

					<?php
					if (isset($_GET['page'])) {
						$hal = $_GET['page'];

						switch ($hal) {

								//Pengguna
							case 'data-pengguna':
								include "admin/pengguna/data_pengguna.php";
								break;
							case 'add-pengguna':
								include "admin/pengguna/add_pengguna.php";
								break;
							case 'edit-pengguna':
								include "admin/pengguna/edit_pengguna.php";
								break;
							case 'del-pengguna':
								include "admin/pengguna/del_pengguna.php";
								break;

								//kartu KK
							case 'data-kartu':
								include "admin/kartu/data_kartu.php";
								break;
							case 'add-kartu':
								include "admin/kartu/add_kartu.php";
								break;
							case 'edit-kartu':
								include "admin/kartu/edit_kartu.php";
								break;
							case 'anggota':
								include "admin/kartu/anggota.php";
								break;
							case 'del-anggota':
								include "admin/kartu/del_anggota.php";
								break;
							case 'del-kartu':
								include "admin/kartu/del_kartu.php";
								break;

								//penduduk
							case 'data-pend':
								include "admin/pend/data_pend.php";
								break;
							case 'add-pend':
								include "admin/pend/add_pend.php";
								break;
							case 'edit-pend':
								include "admin/pend/edit_pend.php";
								break;
							case 'del-pend':
								include "admin/pend/del_pend.php";
								break;
							case 'view-pend':
								include "admin/pend/view_pend.php";
								break;
							case 'upload':
								include "admin/pend/uploader.php";
								break;

								//laki
							case 'data-laki':
								include "admin/laki/data_laki.php";
								break;
							case 'add-laki':
								include "admin/laki/add_laki.php";
								break;
							case 'edit-laki':
								include "admin/laki/edit_laki.php";
								break;
							case 'del-laki':
								include "admin/laki/del_laki.php";
								break;
							case 'view-laki':
								include "admin/laki/view_laki.php";
								break;

								//perempuan
							case 'data-prem':
								include "admin/prem/data_prem.php";
								break;
							case 'add-prem':
								include "admin/prem/add_prem.php";
								break;
							case 'edit-prem':
								include "admin/prem/edit_prem.php";
								break;
							case 'del-prem':
								include "admin/prem/del_prem.php";
								break;
							case 'view-prem':
								include "admin/prem/view_prem.php";
								break;

								//lahir
							case 'data-lahir':
								include "admin/lahir/data_lahir.php";
								break;
							case 'add-lahir':
								include "admin/lahir/add_lahir.php";
								break;
							case 'edit-lahir':
								include "admin/lahir/edit_lahir.php";
								break;
							case 'del-lahir':
								include "admin/lahir/del_lahir.php";
								break;

								//mendu
							case 'data-mendu':
								include "admin/mendu/data_mendu.php";
								break;
							case 'add-mendu':
								include "admin/mendu/add_mendu.php";
								break;
							case 'edit-mendu':
								include "admin/mendu/edit_mendu.php";
								break;
							case 'del-mendu':
								include "admin/mendu/del_mendu.php";
								break;

								//pindah
							case 'data-pindah':
								include "admin/pindah/data_pindah.php";
								break;
							case 'add-pindah':
								include "admin/pindah/add_pindah.php";
								break;
							case 'edit-pindah':
								include "admin/pindah/edit_pindah.php";
								break;
							case 'del-pindah':
								include "admin/pindah/del_pindah.php";
								break;

								//datang
							case 'data-datang':
								include "admin/datang/data_datang.php";
								break;
							case 'add-datang':
								include "admin/datang/add_datang.php";
								break;
							case 'edit-datang':
								include "admin/datang/edit_datang.php";
								break;
							case 'del-datang':
								include "admin/datang/del_datang.php";
								break;

								//suket
							case 'suket-domisili':
								include "surat/suket_domisili.php";
								break;
							case 'suket-lahir':
								include "surat/suket_lahir.php";
								break;
							case 'suket-mati':
								include "surat/suket_mati.php";
								break;
							case 'suket-datang':
								include "surat/suket_datang.php";
								break;
							case 'suket-pindah':
								include "surat/suket_pindah.php";
								break;

								//laporan
								// case 'laporan-domisili':
								// 	include "report/cetak_domisili.php";
								// 	break;
								// case 'laporan-lahir':
								// 	include "report/cetak_lahir.php";
								// 	break;
								// case 'laporan-mati':
								// 	include "report/cetak_mati.php";
								// 	break;
								// case 'laporan-datang':
								// 	include "report/cetak_datang.php";
								// 	break;
								// case 'laporan-pindah':
								// 	include "report/cetak_pindah.php";
								// 	break;

								//stunting
							case 'data-stunting':
								include "admin/stunting/data_stunting.php";
								break;
							case 'add-stunting':
								include "admin/stunting/add_stunting.php";
								break;
							case 'edit-stunting':
								include "admin/stunting/edit_stunting.php";
								break;
							case 'del-stunting':
								include "admin/stunting/del_stunting.php";
								break;
							case 'view-stunting':
								include "admin/stunting/view_stunting.php";
								break;

								//penyintas
							case 'data-penyintas':
								include "admin/penyintas/data_penyintas.php";
								break;
							case 'add-penyintas':
								include "admin/penyintas/add_penyintas.php";
								break;
							case 'edit-penyintas':
								include "admin/penyintas/edit_penyintas.php";
								break;
							case 'del-penyintas':
								include "admin/penyintas/del_penyintas.php";
								break;
							case 'view-penyintas':
								include "admin/penyintas/view_penyintas.php";
								break;

								//berisiko
							case 'data-berisiko':
								include "admin/berisiko/data_berisiko.php";
								break;
							case 'add-berisiko':
								include "admin/berisiko/add_berisiko.php";
								break;
							case 'edit-berisiko':
								include "admin/berisiko/edit_berisiko.php";
								break;
							case 'del-berisiko':
								include "admin/berisiko/del_berisiko.php";
								break;
							case 'view-berisiko':
								include "admin/berisiko/view_berisiko.php";
								break;

								//upload
							case 'import-data':
								include "import.php";
								break;
							case 'export-data':
								include "export.php";
								break;

								//default
							default:
								echo "<center><h1> ERROR !</h1></center>";
								break;
						}
					} else {
						// Auto Halaman Home Pengguna
						if ($data_level == "Administrator") {
							include "home/admin.php";
						} elseif ($data_level == "Kaur Pemerintah") {
							include "home/kaur.php";
						}
					}
					?>

				</div>
			</section>
			<!-- /.content -->
		</div>
		<!-- /.content-wrapper -->

		<footer class="main-footer">
			<div class="float-right d-none d-sm-block">
				Copyright &copy;
				<a target="_blank" href="https://www.instagram.com/kknuh108.watangsoreang/">
					<strong> KKN Posko 11 Watang Soreang</strong>
				</a>
				All rights reserved.
			</div>
			<b>SINTING VERSI 2022</b>
		</footer>

		<!-- Control Sidebar -->
		<aside class="control-sidebar control-sidebar-dark">
			<!-- Control sidebar content goes here -->
		</aside>
		<!-- /.control-sidebar -->
	</div>
	<!-- ./wrapper -->

	<!-- jQuery -->
	<script src="plugins/jquery/jquery.min.js"></script>
	<!-- Bootstrap 4 -->
	<script src="plugins/bootstrap/js/bootstrap.bundle.min.js"></script>
	<!-- Select2 -->
	<script src="plugins/select2/js/select2.full.min.js"></script>
	<!-- DataTables -->
	<script src="plugins/datatables/jquery.dataTables.js"></script>
	<script src="plugins/datatables-bs4/js/dataTables.bootstrap4.js"></script>
	<!-- AdminLTE App -->
	<script src="dist/js/adminlte.min.js"></script>
	<!-- AdminLTE for demo purposes -->
	<script src="dist/js/demo.js"></script>
	<!-- page script -->
	<script src="plugins/jquery-datatable/extensions/export/dataTables.buttons.min.js"></script>
	<script src="plugins/jquery-datatable/extensions/export/buttons.flash.min.js"></script>
	<script src="plugins/jquery-datatable/extensions/export/jszip.min.js"></script>
	<script src="plugins/jquery-datatable/extensions/export/pdfmake.min.js"></script>
	<script src="plugins/jquery-datatable/extensions/export/vfs_fonts.js"></script>
	<script src="plugins/jquery-datatable/extensions/export/buttons.html5.min.js"></script>
	<script src="plugins/jquery-datatable/extensions/export/buttons.print.min.js"></script>

	<script>
		$(function() {
			$("#example1").DataTable();
			$('#example2').DataTable({
				"paging": true,
				"lengthChange": false,
				"searching": false,
				"ordering": true,
				"info": true,
				"autoWidth": false,
			});
		});
	</script>

	<script>
		$(function() {
			//Initialize Select2 Elements
			$('.select2').select2()

			//Initialize Select2 Elements
			$('.select2bs4').select2({
				theme: 'bootstrap4'
			})
		})
	</script>

</body>

</html>