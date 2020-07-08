<?php
defined('BASEPATH') OR exit('No direct script access allowed');

class Phpexcel_model extends CI_Model {

	public function upload_data($filename){
        ini_set('memory_limit', '-1');
        $inputFileName = './assets/uploads/'.$filename;
        try {
        $objPHPExcel = PHPExcel_IOFactory::load($inputFileName);
        } catch(Exception $e) {
        die('Error loading file :' . $e->getMessage());
        }

        $worksheet = $objPHPExcel->getActiveSheet()->toArray(null,true,true,true);
        $numRows = count($worksheet);

        for ($i=1; $i < ($numRows+1) ; $i++) { 
        	$tgl_asli = str_replace('/', '-', $worksheet[$i]['B']);
	        $exp_tgl_asli = explode('-', $tgl_asli);
	        $exp_tahun = explode(' ', $exp_tgl_asli[2]);
	        $tgl_sql = $exp_tahun[0].'-'.$exp_tgl_asli[0].'-'.$exp_tgl_asli[1].' '.$exp_tahun[1];

	        $ins = array(
	        		"nama"          => $worksheet[$i]["A"],
	        		"waktu_absen"   => $tgl_sql
	        	   );

	        $this->db->insert('data', $ins);
	    }
    }

}

/* End of file Phpexcel_model.php */
/* Location: ./application/models/Phpexcel_model.php */