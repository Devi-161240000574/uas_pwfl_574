<?php


class M_crud extends CI_Model
{
	
	public function retrive()
	{
		$this->db->select('idmhs, nim, nama, jenis_kelamin, alamat, nomer_hp');
		$this->db->from('tampil');
		$q = $this->db->get();
		if ($q->num_rows() > 0) {
			$hasil = $q->result_array();
		}else{
			$hasil = array();
		}
		$q->free_result();

		return $hasil;
	}

	public function retrive2()
	{
		$this->db->select('idmhs, nim, no_hp_lama, no_hp_baru, tgl_diubah');
		$this->db->from('histori');
		$q = $this->db->get();
		if ($q->num_rows() > 0) {
			$hasil = $q->result_array();
		}else{
			$hasil = array();
		}
		$q->free_result();

		return $hasil;
	}

	public function donut()
	{
		return $this->db->where(array('jenis_kelamin'=>'Laki-laki'))->count_all_results('mhs');
	}

	public function donut2()
	{
		return $this->db->where(array('jenis_kelamin'=>'Perempuan'))->count_all_results('mhs');
	}

	public function simpan_data($data)
	{
		$q = $this->db->insert('mhs', $data);
		
	}

	public function getRow($idmhs)
	{
		$q = $this->db->where('idmhs', $idmhs)->get('mhs');
		return $q;
	}

	public function update_data()
	{
		$idmhs = $this->input->post('idmhs');
		$data = array(
			'nim' => $this->input->post('nim'),
			'nama' => $this->input->post('nama'),
			'jenis_kelamin' => $this->input->post('jenis_kelamin'),
			'alamat' => $this->input->post('alamat'),
			'nomer_hp' => $this->input->post('nomer_hp')
		);

		$q = $this->db->where('idmhs', $idmhs)->update('mhs', $data);
		if ($q) {
			redirect('Crud/tampilan');
		}else{
			redirect('Crud/ubah');
			echo 'Halaman Tersimpan';
		}
	}

	public function hapus_data()
	{
		$idmhs = $this->uri->segment(3);
		$q = $this->db->where('idmhs', $idmhs)->delete('mhs');
		if ($q) {
			redirect('Crud/tampilan');
		}
	}
}