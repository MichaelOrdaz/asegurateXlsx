<?php
namespace App;

use PhpOffice\PhpSpreadsheet\Spreadsheet;
use PhpOffice\PhpSpreadsheet\Reader\IReader;
use PhpOffice\PhpSpreadsheet\Reader\Xlsx as ReaderXlsx;
use PhpOffice\PhpSpreadsheet\Writer\IWriter;
use PhpOffice\PhpSpreadsheet\Writer\Xlsx as WriterXlsx;

class Xlsx
{
  public static function create($data)
  {
    extract($data);

    //leer excel template
    $reader = new ReaderXlsx();
    $spreadsheet = $reader->load(APP_PATH . '/files/template/certificado_para_bici.xlsx');
    
    //editando sheet
    $worksheet = $spreadsheet->getActiveSheet();
    
    $worksheet->setCellValue('C7', $solicitante->nombre);
    $worksheet->setCellValue('C8', $solicitante->calle);
    $worksheet->setCellValue('C9', $solicitante->colonia);
    $worksheet->setCellValue('C10', $solicitante->municipio);
    $worksheet->setCellValue('C11', $solicitante->email);
    
    $worksheet->setCellValue('E7', $solicitante->rfc);
    $worksheet->setCellValue('E8', $solicitante->cp);
    $worksheet->setCellValue('E9', $solicitante->interior . '/' . $solicitante->exterior);
    $worksheet->setCellValue('E10', $solicitante->estado);
    $worksheet->setCellValue('E11', $solicitante->nacimiento);
    
    $worksheet->setCellValue('C15', $expedicionPoliza->fecha_solicitud);
    $worksheet->setCellValue('E16', $expedicionPoliza->fecha_vigencia);
    
    $worksheet->setCellValue('C21', 'No. Serie: ' . $row->serie . '.' . $row->descripcion);
    
    $worksheet->setCellValue('D26', $prima->monto);
    
    $cobertura = COBERTURAS[$paquete];
    
    $worksheet->setCellValue('D35', $cobertura['danio_material']['suma']);
    $worksheet->setCellValue('E35', $cobertura['danio_material']['deducible']);
    
    $worksheet->setCellValue('D36', $cobertura['escombros']['suma']);
    $worksheet->setCellValue('E36', $cobertura['escombros']['deducible']);
    
    $worksheet->setCellValue('D37', $cobertura['gastos']['suma']);
    $worksheet->setCellValue('E37', $cobertura['gastos']['deducible']);
    
    $worksheet->setCellValue('D38', $cobertura['robo_casa']['suma']);
    $worksheet->setCellValue('E38', $cobertura['robo_casa']['deducible']);
    
    $worksheet->setCellValue('D39', $cobertura['robo_calle']['suma']);
    $worksheet->setCellValue('E39', $cobertura['robo_calle']['deducible']);
    
    $worksheet->setCellValue('D40', $cobertura['muerte']['suma']);
    $worksheet->setCellValue('E40', $cobertura['muerte']['deducible']);
    
    //precio con formato 1,372.06 M.N.
    
    //escribir excel
    $writer = new WriterXlsx($spreadsheet);
    $path = APP_PATH . '/files/odts/' . $filename . '.xlsx';
    $writer->save($path);
    return $path;
  }

}