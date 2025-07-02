<?php

namespace App\Controllers;

use App\Controllers\BaseController;
use PhpOffice\PhpSpreadsheet\Spreadsheet;
use PhpOffice\PhpSpreadsheet\Writer\Xls;
use PhpOffice\PhpSpreadsheet\Writer\Xlsx;
use PhpOffice\PhpSpreadsheet\IOFactory;
use PhpOffice\PhpSpreadsheet\Shared;


public function sendWabcviaexcel()
    {
        // ob_implicit_flush();
        // for ($i = 0; $i < 10; $i++) {
        //     sleep(1);
        // }
        $validationRule = [
            'file_excel' => [
                'label' => 'Image File',
                'rules' => [
                    'uploaded[file_excel]',
                    'is_image[file_excel]',
                    'mime_in[file_excel]',
                    'max_size[file_excel,100]',
                    'ext_in[file_excel,xlsx,xls]',
                    'max_dims[file_excel,1024,768]',
                ],
            ],
        ];
        if (!$this->validate($validationRule)) {
            $output = ['errors' => $this->validator->getErrors()];

            // return view('upload_form', $data);
        }

        $file = $this->request->getFile('file_excel');
        // Tentukan path penyimpanan di folder public/uploadsexcel
        $uploadPath = FCPATH . 'uploadsexcel/'; // FCPATH mengarah ke folder public
        $originalname = $file->getClientName();

        // Ambil file Flyer (jika ada)
        $fileFlayer = $this->request->getFile('file_flayer');
        $uploadPathFlayer = FCPATH . 'uploadsflayer/';
        $urlFlayer = ''; // Default kosong
        if ($fileFlayer && $fileFlayer->isValid() && !$fileFlayer->hasMoved()) {
            $originalNameFlayer = $fileFlayer->getClientName();
            $fileFlayer->move($uploadPathFlayer, $originalNameFlayer);

            // Pastikan ini adalah URL yang bisa diakses publik
            // $urlFlayer = base_url('uploadsflayer/' . $originalNameFlayer);
            $urlFlayer = "http://192.168.1.202/pkrs/public/uploadsflayer/". $originalNameFlayer;
        }

        // Menggunakan PhpSpreadsheet untuk membaca file
        if ($file->isValid() && !$file->hasMoved()) {
            $file->move($uploadPath, $originalname);  // Pindahkan file ke public/uploadsexcel
            $realPath = $uploadPath . $originalname;  // Path lengkap untuk file yang dipindahkan

            // Lanjutkan dengan proses membaca file menggunakan PhpSpreadsheet
            try {
                $inputFileType = 'Xlsx';
                $reader = IOFactory::createReader($inputFileType);
                // Ambil semua sheet yang ada di file Excel
                $worksheetData = $reader->listWorksheetInfo($realPath);

                $sheetName = $this->request->getPost('sheet');
                $datapesan = json_decode($this->request->getPost('datapesan'))->message;
                $sheetExists = false;

                // Validasi apakah sheetName yang diinput ada di dalam file
                foreach ($worksheetData as $worksheet) {
                    if ($worksheet['worksheetName'] === $sheetName) {
                        $sheetExists = true;
                        break;
                    }
                }

                // Jika sheet tidak ditemukan, berikan response error
                if (!$sheetExists) {
                    // Hapus file yang sudah di-upload
                    unlink($realPath);
                    return $this->response->setJSON(['error' => 'Sheet yang diinput tidak ditemukan di file Excel.']);
                }

                $reader->setLoadSheetsOnly($this->request->getPost('sheet'));
                // Nama sheet yang diinput oleh pengguna


                $spreadsheet = $reader->load($realPath);

                $worksheet = $spreadsheet->getActiveSheet();
                $spreadSheetAry = $worksheet->toArray();

                $datanohp = [];
                $datanbc = [];
                $datapasiens = [];
                $datadokters = [];

                $sheetCount = count($spreadSheetAry);

                for ($i = 0; $i <= $sheetCount; $i++) {
                    for ($i = 0; $i < $sheetCount; $i++) {
                        if (empty($spreadSheetAry[$i])) continue; // Skip jika baris kosong

                        $nama_dokter = isset($spreadSheetAry[$i][0]) ? trim($spreadSheetAry[$i][0]) : null;
                        $nama = isset($spreadSheetAry[$i][2]) ? trim($spreadSheetAry[$i][2]) : null;
                        $nohp = isset($spreadSheetAry[$i][19]) && is_numeric($spreadSheetAry[$i][19])
                            ? '0' . $spreadSheetAry[$i][19]
                            : null;

                        // Simpan hanya jika nama dan nohp tersedia
                        if ($nama && $nohp) {
                            $tgl_daftar = isset($spreadSheetAry[$i][12]) && is_numeric($spreadSheetAry[$i][12])
                                ? date('Y-m-d', \PhpOffice\PhpSpreadsheet\Shared\Date::excelToTimestamp($spreadSheetAry[$i][12]))
                                : '1970-01-01';

                            $datanbc[] = [
                                'nama_dokter' => $nama_dokter,
                                'nama' => $nama,
                                'nohp' => $nohp,
                                'tgl_daftar' => $tgl_daftar,
                            ];
                        }
                    }
                }

                $success = 0;
                $error = 0;
                $no = 0;
                $result_log = [];  // Array untuk menyimpan log pengiriman

                foreach ($datanbc as $key => $value) {
                    if (empty($value['nohp'])) continue;  // Lewati jika no HP kosong

                    $nama_pasien = $value['nama'];
                    $nama_dokter = $value['nama_dokter'];

                    // Ganti placeholder di pesan
                    // $txt = str_replace("\n", "\\n", $datapesan);
                    $txt = str_replace(
                        ['{{namapasien}}', '{{namadr}}'],
                        [$nama_pasien, $nama_dokter],
                        $datapesan
                    );
                    $no++;
                    // Delay 5 detik untuk menghindari rate-limit (opsional)
                    sleep(5);
                    try {
                        $no_ho = $value['nohp'];

                        if ($urlFlayer) {
                            // Jika ada file flyer, kirim sebagai media
                            $postData = json_encode([
                                "number" => $no_ho,
                                "caption" => $txt,
                                // "file" => "http://localhost/apicliente-klaim/public/flayer1.jpeg"
                                "file" => $urlFlayer
                            ]);

                            $curlUrl = 'https://192.168.1.203:1243/send-media';
                        } else {
                            // Jika tidak ada file flyer, kirim sebagai teks biasa
                            $postData = json_encode([
                                "number" => $no_ho,
                                "message" => $txt
                            ]);

                            $curlUrl = 'https://192.168.1.203:1243/send-message';
                        }

                        $curl = curl_init();
                        curl_setopt_array($curl, [
                            CURLOPT_URL => $curlUrl,
                            CURLOPT_RETURNTRANSFER => true,
                            CURLOPT_ENCODING => '',
                            CURLOPT_MAXREDIRS => 30,
                            CURLOPT_TIMEOUT => 0,
                            CURLOPT_FOLLOWLOCATION => true,
                            CURLOPT_HTTP_VERSION => CURL_HTTP_VERSION_1_1,
                            CURLOPT_SSL_VERIFYPEER => false,
                            CURLOPT_SSL_VERIFYHOST => false,
                            CURLOPT_CUSTOMREQUEST => 'POST',
                            CURLOPT_POSTFIELDS => $postData,
                            CURLOPT_HTTPHEADER => [
                                'Content-Type: application/json'
                            ],
                        ]);

                        $response = curl_exec($curl);
                        $curl_error = curl_error($curl);
                        curl_close($curl);

                        if ($curl_error) {
                            $error++;
                            $result_log[] = [
                                'no_hp' => $no_ho,
                                'status' => 'gagal',
                                'error' => $curl_error
                            ];
                        } else {
                            $responseData = json_decode($response);
                            if ($responseData->status) {
                                $success++;
                                $result_log[] = [
                                    'no_hp' => $no_ho,
                                    'status' => 'berhasil'
                                ];
                            } else {
                                $error++;
                                $result_log[] = [
                                    'no_hp' => $no_ho,
                                    'status' => 'gagal',
                                    'error' => $responseData->message
                                ];
                            }
                        }
                    } catch (\Throwable $t) {
                        $error++;
                        $result_log[] = [
                            'no_hp' => $no_ho ?? 'tidak diketahui',
                            'status' => 'gagal',
                            'error' => $t->getMessage()
                        ];
                        continue;
                    }
                }

                // Hapus file setelah selesai
                unlink($realPath);

                // Kirim response JSON
                $output = [
                    'datapesan' => $datapesan,
                    'berhasil' => $success,
                    'gagal' => $error,
                    'log_pengiriman' => $result_log
                ];

                return $this->response->setJSON($output);
            } catch (\Exception $e) {
                return $this->response->setJSON(['error' => $e->getMessage()]);
            }
        } else {
            return $this->response->setJSON(['error' => 'File sudah dipindahkan']);
        }
    }