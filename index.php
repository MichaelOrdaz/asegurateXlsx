<?php
require 'vendor/autoload.php';

use App\Xlsx;

$solicitante = [
  'nombre' => 'Michael Ordaz Martinez',
  'calle' => 'Francisco VIlla mz 2 lote 1',
  'colonia' => 'Amplición Ignacio Mariscal',
  'municipio' => 'Puebla',
  'email' => 'ordazmartinezmichael@gmail.com',
  'rfc' => 'LOREMIPUSM76',
  'cp' => '72014',
  'interior' => '3S',
  'exterior' => '345',
  'estado' => 'Puebla',
  'nacimiento' => '1993-11-28',
];
$solicitante = (object) $solicitante;

$now = new DateTime();
$nextYear = clone $now;
$nextYear->add(new DateInterval('P1Y'));

$expedicionPoliza = [
  'fecha_solicitud' => $now->format('d-m-Y'),
  'fecha_vigencia' => $nextYear->format('d-m-Y'),
];

$expedicionPoliza = (object) $expedicionPoliza;

$filename = $solicitante->nombre . '-' . uniqid();

$row = [
  'descripcion' => 'Lorem ipsum dolor sit amet consectetur adipisicing elit. Tempore sequi asperiores dolorum iusto vero nesciunt delectus esse placeat, at itaque, inventore ad rerum magni impedit recusandae quos obcaecati corrupti! Sed?',
  'serie' => uniqid(),
];
$row = (object) $row;

$prima = ['monto' => '1,827.90 M.N.'];
$prima = (object) $prima;

$paquete = strtolower('bronce');
// datos que recibira el excel arriba

Xlsx::create(compact('solicitante', 'expedicionPoliza', 'filename', 'row', 'prima', 'paquete'));