<!doctype html>
<html lang="en">

<head>
	<base href="<?= base_url() ?>">
	<meta charset="UTF-8">
	<title>Tampil Data</title>
</head>

<body>
	<div class="content-wrapper">
		<section class="content-header">
			<h1>
				Data Tables
<!-- 				<small>advanced tables</small>
 -->			</h1>
			<ol class="breadcrumb">
				<li><a href="#"><i class="fa fa-dashboard"></i> Home</a></li>
				<li><a href="#">Tables</a></li>
				<li class="active">Data tables</li>
			</ol>
		</section>
		<section class="content">
			<div class="row">
				<div class="box">
					<div class="box-header">
						<a href="crud/tambah" class="btn btn-primary">Tambah Data</a>
						<a href="<?= base_url(); ?>crud/excel" class="btn btn-primary">Export Excel</a>
						<!-- <a href="crud/tampil2" class="btn btn-primary">Riwayat Log</a> -->
						<div class="box-body">
							<table id="example1" class="table table-bordered table-striped">
								<thead>
									<tr>
										<th>ID</th>
										<th>Nim</th>
										<th>Nama</th>
										<th>Jenis Kelamin</th>
										<th>Alamat</th>
										<th>Nomer Hp</th>
										<th>Edit</th>
										<th>Hapus</th>
										</tr>
								</thead>
								<?php $id=1;
								foreach ($isi as $k0 => $v0) {
									?>
									<tr>
										<td><?php echo $v0['idmhs'] ?></td>
										<td><?php echo $v0['nim'] ?></td>
										<td><?php echo $v0['nama'] ?></td>
										<td><?php echo $v0['jenis_kelamin'] ?></td>
										<td><?php echo $v0['alamat'] ?></td>
										<td><?php echo $v0['nomer_hp'] ?></td>
										<td><a href='<?php echo "crud/ubah/".$v0["idmhs"] ?>'>ubah<a/></td>
										<td><a href='<?php echo "crud/hapus/".$v0["idmhs"] ?>'>hapus<a/></td>
									</tr>
								<?php } ?>
							</table>
		</section>
	</div>
	</div>
	</div>
	</div>


</body>