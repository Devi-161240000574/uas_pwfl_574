<!doctype html>
<html lang="en">

<head>
	<base href="<?= base_url() ?>">
	<meta charset="UTF-8">
</head>

<div class="content-wrapper">
	<section class="content-header">
		<h1>
			Data Tables
		</h1>
		<ol class="breadcrumb">
			<li><a href="#"><i class="fa fa-dashboard"></i> Home</a></li>
			<li><a href="#">Tables</a></li>
		</ol>
	</section>
	<section class="content">
		<div class="row">
			<div class="box">
				<div class="box-header">
	<?php echo "<h3>" . $judul . "</h3>"; ?>

		<form method="post" action="<?php echo base_url('crud/update')?>" enctype="multipart/form-data">
	<table>
		<input type="hidden" name="idmhs" value="<?php echo $idmhs ?>">
			<tr>
		<td>NIM</td>
		<td>:</td>
		<td><input type="text" name="nim" value="<?php echo $nim ?>"></td>
			</tr>

			<tr>
		<td>Nama</td>
		<td>:</td>
		<td><input type="text" name="nama" value="<?php echo $nama ?>"></td>
			</tr>

		<tr>
		<td>Jenis Kelamin</td>
		<td>:</td>
		<td>
			<input type="radio" name="jenis_kelamin" value="Laki-laki" <?php echo $jenis_kelamin == 'L' ? 'chechked' : '' ?>>Laki-laki
			<input type="radio" name="jenis_kelamin" value="Perempuan" <?php echo $jenis_kelamin == 'P' ? 'chechked' : '' ?>>Perempuan
		</td>
		</tr>

		<tr>
		<td>Alamat</td>
		<td>:</td> 
		<td><input type="text" name="alamat" value="<?php echo $alamat ?>"></td>
		</tr>
		
		<tr>
		<td>Nomer Hp</td>
		<td>:</td> 
		<td><input type="text" name="nomer_hp" value="<?php echo $nomer_hp ?>"></td>
		</tr>

		<tr>
		<td>
		<input type="submit" name="simpan" value="Update">
		<a href="<?php echo base_url().'crud/tampilan';?>"><input type="button" value="Batal"></a>
		</td>
		</tr>
	</table>
	</form>
</div>
</div>
</section>
</div>