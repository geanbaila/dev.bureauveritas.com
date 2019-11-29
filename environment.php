<?php
define("DEBUG", true);//false
$GLOBALS["HOST"] = "localhost";
$GLOBALS["USUARIO"] = "root";
$GLOBALS["PASSWORD"] = "root";
$GLOBALS["DATABASE"] = "pe_bv_scs_homo";

$GLOBALS["RESPUESTAS"]=array(
	2=>array(1=>"Sí",2=>"No"),
	7=>array(1=>"Sí",2=>"No",3=>"No aplica")
);

$GLOBALS["GRIS"]="CECECE";
/*cambio del hexadecimal a pedido del cliente..*/
$GLOBALS["NEGRO"] = "4B505F";
$GLOBALS["AZUL"] = "4B505F";
$GLOBALS["AZUL_MARINO"] = "4B505F";
$GLOBALS["ROJO"] = "4B505F";
$GLOBALS["VERDE"] = "BDECD3";
$GLOBALS["BLANCO"] = "FFFFFF";

/*$GLOBALS["NEGRO"] = "4B505F";
$GLOBALS["AZUL"] = "4F5F98";
$GLOBALS["AZUL_MARINO"] = "A87D32";
$GLOBALS["ROJO"] = "FF5F98";*/

$GLOBALS["categorias"] = array("I.","II.","III.","IV.","V.","VI.","VII.","VIII.","IX.","X.","XI.","XII.","XIII.","XIV.","XV.","XVI.","XVII.","XVIII.","XIX.","XX.","XXI.","XXII.","XXIII.","XXIV.","XXV.","XXVI.","XXVII.","XXVIII.","XXIX.","XXX.","XXXI.","XXXII.","XXXIII.","XXXIV.","XXXV.","XXXVI.","XXXVII.","XXXVIII.","XXXIX.","XL.","XLI.","XLII.","XLIII.","XLIV.");

define('ANCHO_ALTERNATIVAS', 1600);
define('ANCHO_PUNTUACION', 200);
define('ORDENAR_DESCENDENTE', true);
define('REEMPLAZAR_VACIOS', '-');
define('PREGUNTA_ALINEADA', 'left');
define('PREGUNTA_CERRADA_SIMPLE', 2);
define('PREGUNTA_CERRADA_COMPLEJA', 7);
define('PREGUNTA_CABECERA', 21);
define('PREGUNTA_MULTIPLE', 1);
define('PREGUNTA_FINAL', 2);
define('PUNTAJE_CALIFICADO', 1);
define('PUNTAJE_INFORMATIVO', 0);
define('PUNTAJE_ACUMULADO', 2);
define('SEPARADOR', '');
define('INFORMATIVO', '');
define("PUBLIC_RESOURCES_INFORMES", __DIR__.'/resources/');
define("PUBLIC_RESOURCES_FOTOS", __DIR__.'/../userfiles/cms/homologacion/foto/');
define("IMAGE_CELL_WIDTH", 2000);
define("IMAGE_WIDTH",150);
define("IMAGE_HEIGHT",150);
define("INFORME_INICIAL", __DIR__.'/resources/archivo_base.docx');
define("INFORME_PARTE_UNO", 'parte_1.docx');
define("INFORME_PARTE_DOS",'parte_2.docx');
define("INFORME_FINAL", __DIR__.'/resources/informe_general.docx');
define("OCULTAR_SCORE_CABECERAS", true);
define("OCULTAR_SCORE_SUB_CONTENEDORES", true);
 
$GLOBALS["WIDTH_FOR_QUESTION"]=[
1=>ANCHO_ALTERNATIVAS,
2=>ANCHO_ALTERNATIVAS,
3=>ANCHO_ALTERNATIVAS,
4=>ANCHO_ALTERNATIVAS,
5=>ANCHO_ALTERNATIVAS,
6=>ANCHO_ALTERNATIVAS,
7=>ANCHO_ALTERNATIVAS,
20=>ANCHO_ALTERNATIVAS,
21=>ANCHO_ALTERNATIVAS,
22=>ANCHO_ALTERNATIVAS,
23=>ANCHO_ALTERNATIVAS,
24=>ANCHO_ALTERNATIVAS,
25=>ANCHO_ALTERNATIVAS,
0=>ANCHO_ALTERNATIVAS	
];