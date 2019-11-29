<?php

require_once __DIR__ . '/vendor/autoload.php';

use \DocxMerge\DocxMerge;

require_once __DIR__ . '/environment.php';
require_once __DIR__ . '/util.php';
$homologacionID = ($_REQUEST["homologacionID"] + 0);
$codigo_informe = uniqid();
$cellRowSpan = array('vMerge' => 'restart', 'valign' => 'center', 'bgColor' => $GLOBALS["BLANCO"]);
$cellRowContinue = array('vMerge' => 'continue', 'valign' => 'center', 'bgColor' => $GLOBALS["BLANCO"]);
$cellHCentered = array('alignment' => \PhpOffice\PhpWord\SimpleType\Jc::CENTER);
$cellVCentered = array('valign' => 'center');

$sql = "SELECT *,
	(select concat(firstName,' ',lastName) from adm_user where userID = crm_homologacion.userID ) as _auditedBy,
	date_format(crm_homologacion.registerUpdate,'%d/%m/%Y') as fecha_emision_evaluacion,
	(SELECT countryName  FROM ubg_country WHERE countryID = (select country from crm_proveedor where proveedorID = (select proveedorID from crm_requerimiento where requerimientoID = crm_homologacion.requerimientoID) )) as countryName,
	(SELECT nombre  FROM crm_ubigeo WHERE cod_dpto = (select department from crm_proveedor where proveedorID = (select proveedorID from crm_requerimiento where requerimientoID = crm_homologacion.requerimientoID)) 
		and cod_prov like '00' 
		and cod_dist like '00') as nombreDepartment,
	(SELECT nombre  FROM crm_ubigeo WHERE cod_dpto = (select department from crm_proveedor where proveedorID = (select proveedorID from crm_requerimiento where requerimientoID = crm_homologacion.requerimientoID)) 
		and cod_prov = (select province from crm_proveedor where proveedorID = (select proveedorID from crm_requerimiento where requerimientoID = crm_homologacion.requerimientoID)) 
		and cod_dist like '00') as nombreProvince,
	(SELECT nombre  FROM crm_ubigeo WHERE cod_dpto = (select department from crm_proveedor where proveedorID = (select proveedorID from crm_requerimiento where requerimientoID = crm_homologacion.requerimientoID))
		and cod_prov = (select province from crm_proveedor where proveedorID = (select proveedorID from crm_requerimiento where requerimientoID = crm_homologacion.requerimientoID))
		and cod_dist like (select district from crm_proveedor where proveedorID = (select proveedorID from crm_requerimiento where requerimientoID = crm_homologacion.requerimientoID))) as nombreDistrict, 
	(SELECT titleForm  FROM crm_form_propuesta WHERE propxformID = (select propxformID from crm_requerimiento where requerimientoID = crm_homologacion.requerimientoID)) as titleForm,
	(SELECT proposalNumber from crm_propuesta where propuestaID = (SELECT propuestaID FROM crm_form_propuesta where propxformID = (select propxformID from crm_requerimiento where requerimientoID = crm_homologacion.requerimientoID))) as proposalNumber,
        
	(select homologacionID from crm_datos_homo where homologacionID = crm_homologacion.homologacionID) as homologacionID,
	(select rcc from crm_datos_homo where homologacionID = crm_homologacion.homologacionID) as rcc,
	(select clasif_crediticia from crm_datos_homo where homologacionID = crm_homologacion.homologacionID) as clasif_crediticia,
	(select cuentas_cerradas from crm_datos_homo where homologacionID = crm_homologacion.homologacionID) as cuentas_cerradas,
	(select boletines from crm_datos_homo where homologacionID = crm_homologacion.homologacionID) as boletines,
	(select consolidado from crm_datos_homo where homologacionID = crm_homologacion.homologacionID) as consolidado,
	(select negativo_sunat from crm_datos_homo where homologacionID = crm_homologacion.homologacionID) as negativo_sunat,
	(select comercio_exterior from crm_datos_homo where homologacionID = crm_homologacion.homologacionID) as comercio_exterior,
	(select rectifica from crm_datos_homo where homologacionID = crm_homologacion.homologacionID) as rectifica,
	(select deuda_previsional from crm_datos_homo where homologacionID = crm_homologacion.homologacionID) as deuda_previsional,
	(select observaciones from crm_datos_homo where homologacionID = crm_homologacion.homologacionID) as observaciones,
	(select recomendaciones from crm_datos_homo where homologacionID = crm_homologacion.homologacionID) as recomendaciones,
	(select registerDate from crm_datos_homo where homologacionID = crm_homologacion.homologacionID) as registerDate,
	(select registerUpdate from crm_datos_homo where homologacionID = crm_homologacion.homologacionID) as registerUpdate,

	(select requerimientoID from crm_requerimiento where requerimientoID = crm_homologacion.requerimientoID) as requerimientoID,
	(select period from crm_requerimiento where requerimientoID = crm_homologacion.requerimientoID) as period,
	(select propxformID from crm_requerimiento where requerimientoID = crm_homologacion.requerimientoID) as propxformID,
	(select proveedorID from crm_requerimiento where requerimientoID = crm_homologacion.requerimientoID) as proveedorID,
	(select observation from crm_requerimiento where requerimientoID = crm_homologacion.requerimientoID) as observation,
	(select amount from crm_requerimiento where requerimientoID = crm_homologacion.requerimientoID) as amount,
	(select threeDay from crm_requerimiento where requerimientoID = crm_homologacion.requerimientoID) as threeDay,
	(select nineDay from crm_requerimiento where requerimientoID = crm_homologacion.requerimientoID) as nineDay,
	(select fourteenDay from crm_requerimiento where requerimientoID = crm_homologacion.requerimientoID) as fourteenDay,
	(select alert from crm_requerimiento where requerimientoID = crm_homologacion.requerimientoID) as alert,
	(select state from crm_requerimiento where requerimientoID = crm_homologacion.requerimientoID) as state,
	(select registerDate from crm_requerimiento where requerimientoID = crm_homologacion.requerimientoID) as registerDate,
	(select registerExpire from crm_requerimiento where requerimientoID = crm_homologacion.requerimientoID) as registerExpire,
	(select registerUpdate from crm_requerimiento where requerimientoID = crm_homologacion.requerimientoID) as registerUpdate,

	(select proveedorID from crm_proveedor where proveedorID = (select proveedorID from crm_requerimiento where requerimientoID = crm_homologacion.requerimientoID) ) as proveedorID,
	(select documentNumber from crm_proveedor where proveedorID = (select proveedorID from crm_requerimiento where requerimientoID = crm_homologacion.requerimientoID) ) as documentNumber,
	(select user from crm_proveedor where proveedorID = (select proveedorID from crm_requerimiento where requerimientoID = crm_homologacion.requerimientoID) ) as user,
	(select pass from crm_proveedor where proveedorID = (select proveedorID from crm_requerimiento where requerimientoID = crm_homologacion.requerimientoID) ) as pass,
	(select typeProvider from crm_proveedor where proveedorID = (select proveedorID from crm_requerimiento where requerimientoID = crm_homologacion.requerimientoID) ) as typeProvider,
	 
	(select businessName from crm_cliente where clienteID = (select clienteID from crm_propuesta where propuestaID = (select propuestaID from crm_form_propuesta  where propxformID = (SELECT propxformID FROM pe_bv_scs_homo.crm_requerimiento WHERE requerimientoID = crm_homologacion.requerimientoID)))
	) as cliente_businessName,

	(select businessName from crm_proveedor where proveedorID = (select proveedorID from crm_requerimiento where requerimientoID = crm_homologacion.requerimientoID) ) as businessName,
	(select ciudad from crm_proveedor where proveedorID = (select proveedorID from crm_requerimiento where requerimientoID = crm_homologacion.requerimientoID) ) as ciudad,
	(select address from crm_proveedor where proveedorID = (select proveedorID from crm_requerimiento where requerimientoID = crm_homologacion.requerimientoID) ) as address,
	(select phone from crm_proveedor where proveedorID = (select proveedorID from crm_requerimiento where requerimientoID = crm_homologacion.requerimientoID) ) as phone,
	(select phone from crm_proveedor where proveedorID = (select proveedorID from crm_requerimiento where requerimientoID = crm_homologacion.requerimientoID) ) as telefono_proveedor,
	(select email from crm_proveedor where proveedorID = (select proveedorID from crm_requerimiento where requerimientoID = crm_homologacion.requerimientoID) ) as email,
	(select other from crm_proveedor where proveedorID = (select proveedorID from crm_requerimiento where requerimientoID = crm_homologacion.requerimientoID) ) as other,
	(select bienID from crm_proveedor where proveedorID = (select proveedorID from crm_requerimiento where requerimientoID = crm_homologacion.requerimientoID) ) as bienID,
	(select servicioID from crm_proveedor where proveedorID = (select proveedorID from crm_requerimiento where requerimientoID = crm_homologacion.requerimientoID) ) as servicioID,
	(select country from crm_proveedor where proveedorID = (select proveedorID from crm_requerimiento where requerimientoID = crm_homologacion.requerimientoID) ) as country,
	(select department from crm_proveedor where proveedorID = (select proveedorID from crm_requerimiento where requerimientoID = crm_homologacion.requerimientoID) ) as department,
	(select province from crm_proveedor where proveedorID = (select proveedorID from crm_requerimiento where requerimientoID = crm_homologacion.requerimientoID) ) as province,
	(select district from crm_proveedor where proveedorID = (select proveedorID from crm_requerimiento where requerimientoID = crm_homologacion.requerimientoID) ) as district,
	(select postalCode from crm_proveedor where proveedorID = (select proveedorID from crm_requerimiento where requerimientoID = crm_homologacion.requerimientoID) ) as postalCode,
	(select fax from crm_proveedor where proveedorID = (select proveedorID from crm_requerimiento where requerimientoID = crm_homologacion.requerimientoID) ) as fax,
	(select contacts from crm_proveedor where proveedorID = (select proveedorID from crm_requerimiento where requerimientoID = crm_homologacion.requerimientoID) ) as contacts,
	(select legalDirection from crm_proveedor where proveedorID = (select proveedorID from crm_requerimiento where requerimientoID = crm_homologacion.requerimientoID) ) as legalDirection,
	(select departmentLegal from crm_proveedor where proveedorID = (select proveedorID from crm_requerimiento where requerimientoID = crm_homologacion.requerimientoID) ) as departmentLegal,
	(select provinceLegal from crm_proveedor where proveedorID = (select proveedorID from crm_requerimiento where requerimientoID = crm_homologacion.requerimientoID) ) as provinceLegal,
	(select districtLegal from crm_proveedor where proveedorID = (select proveedorID from crm_requerimiento where requerimientoID = crm_homologacion.requerimientoID) ) as districtLegal,
	(select legalRepresentative from crm_proveedor where proveedorID = (select proveedorID from crm_requerimiento where requerimientoID = crm_homologacion.requerimientoID) ) as legalRepresentative,
	(select commercialContactName from crm_proveedor where proveedorID = (select proveedorID from crm_requerimiento where requerimientoID = crm_homologacion.requerimientoID) ) as commercialContactName,
	(select commercialContactPhone from crm_proveedor where proveedorID = (select proveedorID from crm_requerimiento where requerimientoID = crm_homologacion.requerimientoID) ) as commercialContactPhone,
	(select commercialContactCellphone from crm_proveedor where proveedorID = (select proveedorID from crm_requerimiento where requerimientoID = crm_homologacion.requerimientoID) ) as commercialContactCellphone,
	(select commercialContactEmail from crm_proveedor where proveedorID = (select proveedorID from crm_requerimiento where requerimientoID = crm_homologacion.requerimientoID) ) as commercialContactEmail,
	(select generalManagerName from crm_proveedor where proveedorID = (select proveedorID from crm_requerimiento where requerimientoID = crm_homologacion.requerimientoID) ) as generalManagerName,
	(select generalManagerPhone from crm_proveedor where proveedorID = (select proveedorID from crm_requerimiento where requerimientoID = crm_homologacion.requerimientoID) ) as generalManagerPhone,
	(select generalManagerCellphone from crm_proveedor where proveedorID = (select proveedorID from crm_requerimiento where requerimientoID = crm_homologacion.requerimientoID) ) as generalManagerCellphone,
	(select generalManagerEmail from crm_proveedor where proveedorID = (select proveedorID from crm_requerimiento where requerimientoID = crm_homologacion.requerimientoID) ) as generalManagerEmail,
	(select numberCollaborateAdmin from crm_proveedor where proveedorID = (select proveedorID from crm_requerimiento where requerimientoID = crm_homologacion.requerimientoID) ) as numberCollaborateAdmin,
	(select numberCollaborateOper from crm_proveedor where proveedorID = (select proveedorID from crm_requerimiento where requerimientoID = crm_homologacion.requerimientoID) ) as numberCollaborateOper,
	(select workShifts from crm_proveedor where proveedorID = (select proveedorID from crm_requerimiento where requerimientoID = crm_homologacion.requerimientoID) ) as workShifts,
	(select businessAction1 from crm_proveedor where proveedorID = (select proveedorID from crm_requerimiento where requerimientoID = crm_homologacion.requerimientoID) ) as businessAction1,
	(select percentageParticipant1 from crm_proveedor where proveedorID = (select proveedorID from crm_requerimiento where requerimientoID = crm_homologacion.requerimientoID) ) as percentageParticipant1,
	(select businessAction2 from crm_proveedor where proveedorID = (select proveedorID from crm_requerimiento where requerimientoID = crm_homologacion.requerimientoID) ) as businessAction2,
	(select percentageParticipant2 from crm_proveedor where proveedorID = (select proveedorID from crm_requerimiento where requerimientoID = crm_homologacion.requerimientoID) ) as percentageParticipant2,
	(select businessAction3 from crm_proveedor where proveedorID = (select proveedorID from crm_requerimiento where requerimientoID = crm_homologacion.requerimientoID) ) as businessAction3,
	(select percentageParticipant3 from crm_proveedor where proveedorID = (select proveedorID from crm_requerimiento where requerimientoID = crm_homologacion.requerimientoID) ) as percentageParticipant3,
	(select activityDate from crm_proveedor where proveedorID = (select proveedorID from crm_requerimiento where requerimientoID = crm_homologacion.requerimientoID) ) as activityDate,
	(select partnerships from crm_proveedor where proveedorID = (select proveedorID from crm_requerimiento where requerimientoID = crm_homologacion.requerimientoID) ) as partnerships,
	(select ecoActivity from crm_proveedor where proveedorID = (select proveedorID from crm_requerimiento where requerimientoID = crm_homologacion.requerimientoID) ) as ecoActivity,
	(select retentionIgv from crm_proveedor where proveedorID = (select proveedorID from crm_requerimiento where requerimientoID = crm_homologacion.requerimientoID) ) as retentionIgv,
	(select observation from crm_proveedor where proveedorID = (select proveedorID from crm_requerimiento where requerimientoID = crm_homologacion.requerimientoID) ) as observation,
	(select registration from crm_proveedor where proveedorID = (select proveedorID from crm_requerimiento where requerimientoID = crm_homologacion.requerimientoID) ) as registration,
	(select testConstitution from crm_proveedor where proveedorID = (select proveedorID from crm_requerimiento where requerimientoID = crm_homologacion.requerimientoID) ) as testConstitution,
	(select firm from crm_proveedor where proveedorID = (select proveedorID from crm_requerimiento where requerimientoID = crm_homologacion.requerimientoID) ) as firm,
	(select representation from crm_proveedor where proveedorID = (select proveedorID from crm_requerimiento where requerimientoID = crm_homologacion.requerimientoID) ) as representation,
	(select licence from crm_proveedor where proveedorID = (select proveedorID from crm_requerimiento where requerimientoID = crm_homologacion.requerimientoID) ) as licence,
	(select certInspeccion from crm_proveedor where proveedorID = (select proveedorID from crm_requerimiento where requerimientoID = crm_homologacion.requerimientoID) ) as certInspeccion,
	(select registerMine from crm_proveedor where proveedorID = (select proveedorID from crm_requerimiento where requerimientoID = crm_homologacion.requerimientoID) ) as registerMine,
	(select cargaMasiva from crm_proveedor where proveedorID = (select proveedorID from crm_requerimiento where requerimientoID = crm_homologacion.requerimientoID) ) as cargaMasiva,
	(select registerDate from crm_proveedor where proveedorID = (select proveedorID from crm_requerimiento where requerimientoID = crm_homologacion.requerimientoID) ) as registerDate,
	(select registerUpdate from crm_proveedor where proveedorID = (select proveedorID from crm_requerimiento where requerimientoID = crm_homologacion.requerimientoID) ) as registerUpdate,
	(select state from crm_proveedor where proveedorID = (select proveedorID from crm_requerimiento where requerimientoID = crm_homologacion.requerimientoID) ) as state,
	1
	FROM crm_homologacion 	
	WHERE crm_homologacion.homologacionID = " . $homologacionID . " 
	";
$database = new DataBase();
$result = $database->getResult($sql);
if ($result["num_rows"] > 0):
    $templateProcessor = new PhpOffice\PhpWord\TemplateProcessor(INFORME_INICIAL);
    foreach ($result["data"] as $data):
        $nivel = $data["nivel"];
        $puntajeFinalTotal = $data["puntajeFinal"];
        $observaciones = $data["observaciones"];
        $recomendaciones = $data["recomendaciones"];
        $propxformID = ($data["propxformID"] + 0);
        $proposalNumber = clearText($data["proposalNumber"]);
        $proposalNumber = ($proposalNumber == "") ? REEMPLAZAR_VACIOS : $proposalNumber;
        $templateProcessor->setValue("proposalNumber", $proposalNumber);

        //DATOS GENERALES
        //razón social
        $auditedBy = clearText($data["_auditedBy"]);
        $auditedBy = ($auditedBy == "") ? REEMPLAZAR_VACIOS : $auditedBy;
        $businessName = clearText($data["businessName"]);
        $businessName = ($businessName == "") ? REEMPLAZAR_VACIOS : $businessName;
        $templateProcessor->setValue("businessName", $businessName);
        //ruc
        $documentNumber = clearText($data["documentNumber"]);
        $documentNumber = ($documentNumber == "") ? REEMPLAZAR_VACIOS : $documentNumber;
        $templateProcessor->setValue("documentNumber", $documentNumber);
        //actividad económica
        $ecoActivity = clearText($data["ecoActivity"]);
        $ecoActivity = ($ecoActivity == "") ? REEMPLAZAR_VACIOS : $ecoActivity;
        $templateProcessor->setValue("ecoActivity", $ecoActivity);
        //pais
        $countryName = clearText($data["countryName"]);
        $countryName = ($countryName == "") ? REEMPLAZAR_VACIOS : $countryName;
        $templateProcessor->setValue("countryName", $countryName);
        //código postal
        $postalCode = clearText($data["postalCode"]);
        $postalCode = ($postalCode == "") ? REEMPLAZAR_VACIOS : $postalCode;
        $templateProcessor->setValue("postalCode", $postalCode);
        //departmento
        $nombreDepartment = clearText($data["nombreDepartment"]);
        $nombreDepartment = ($nombreDepartment == "") ? REEMPLAZAR_VACIOS : $nombreDepartment;
        $templateProcessor->setValue("nombreDepartment", $nombreDepartment);
        //provincia
        $nombreProvince = clearText($data["nombreProvince"]);
        $nombreProvince = ($nombreProvince == "") ? REEMPLAZAR_VACIOS : $nombreProvince;
        $templateProcessor->setValue("nombreProvince", $nombreProvince);
        //distrito
        $nombreDistrict = clearText($data["nombreDistrict"]);
        $nombreDistrict = ($nombreDistrict == "") ? REEMPLAZAR_VACIOS : $nombreDistrict;
        $templateProcessor->setValue("nombreDistrict", $nombreDistrict);
        //email
        $commercialContactEmail = clearText($data["commercialContactEmail"]);
        $commercialContactEmail = ($commercialContactEmail == "") ? REEMPLAZAR_VACIOS : $commercialContactEmail;
        $templateProcessor->setValue("commercialContactEmail", $commercialContactEmail);
        //agente de retención de IGV
        $templateProcessor->setValue("AGENTE_RETENCION", "-");
        //teléfono
        $commercialContactCellphone = clearText($data["commercialContactCellphone"]);
        $commercialContactCellphone = ($commercialContactCellphone == "") ? REEMPLAZAR_VACIOS : $commercialContactCellphone;
        $templateProcessor->setValue("telefono_proveedor", $commercialContactCellphone);
        //fecha de visita
        $fecha_emision_evaluacion = clearText($data["fecha_emision_evaluacion"]);
        $fecha_emision_evaluacion = ($fecha_emision_evaluacion == "") ? REEMPLAZAR_VACIOS : $fecha_emision_evaluacion;
        $templateProcessor->setValue("FECHA_VISITA", $fecha_emision_evaluacion);
        //dirección legal
        $legalDirection = clearText($data["legalDirection"]);
        $legalDirection = ($legalDirection == "") ? REEMPLAZAR_VACIOS : $legalDirection;
        $templateProcessor->setValue("legalDirection", $legalDirection);
        //representante legal
        $legalRepresentative = clearText($data["legalRepresentative"]);
        $legalRepresentative = ($legalRepresentative == "") ? REEMPLAZAR_VACIOS : $legalRepresentative;
        $templateProcessor->setValue("legalRepresentative", $legalRepresentative);
        //dirección visita
        $address = clearText($data["address"]);
        $address = ($address == "") ? REEMPLAZAR_VACIOS : $address;
        $templateProcessor->setValue("address", $address);

        $businessAction1 = clearText($data["businessAction1"]);
        $businessAction1 = ($businessAction1 == "") ? REEMPLAZAR_VACIOS : $businessAction1;
        $templateProcessor->setValue("businessAction1", $businessAction1);

        $businessAction2 = clearText($data["businessAction2"]);
        $businessAction2 = ($businessAction2 == "") ? REEMPLAZAR_VACIOS : $businessAction2;
        $templateProcessor->setValue("businessAction2", $businessAction2);

        $businessAction3 = clearText($data["businessAction3"]);
        $businessAction3 = ($businessAction3 == "") ? REEMPLAZAR_VACIOS : $businessAction3;
        $templateProcessor->setValue("businessAction3", $businessAction3);

        $percentageParticipant1 = clearText($data["percentageParticipant1"]);
        $percentageParticipant1 = ($percentageParticipant1 == "") ? REEMPLAZAR_VACIOS : $percentageParticipant1;
        $templateProcessor->setValue("percentageParticipant1", $percentageParticipant1);

        $percentageParticipant2 = clearText($data["percentageParticipant2"]);
        $percentageParticipant2 = ($percentageParticipant2 == "") ? REEMPLAZAR_VACIOS : $percentageParticipant2;
        $templateProcessor->setValue("percentageParticipant2", $percentageParticipant2);

        $percentageParticipant3 = clearText($data["percentageParticipant3"]);
        $percentageParticipant3 = ($percentageParticipant3 == "") ? REEMPLAZAR_VACIOS : $percentageParticipant3;
        $templateProcessor->setValue("percentageParticipant3", $percentageParticipant3);

        $registration = clearText($data["registration"]);
        $registration = ($registration == "") ? REEMPLAZAR_VACIOS : $registration;
        $templateProcessor->setValue("registration", $registration);

        $testConstitution = clearText($data["testConstitution"]);
        $testConstitution = ($testConstitution == "") ? REEMPLAZAR_VACIOS : $testConstitution;
        $templateProcessor->setValue("testConstitution", $testConstitution);

        $representation = clearText($data["representation"]);
        $representation = ($representation == "") ? REEMPLAZAR_VACIOS : $representation;
        $templateProcessor->setValue("representation", $representation);

        $licence = clearText($data["licence"]);
        $licence = ($licence == "") ? REEMPLAZAR_VACIOS : $licence;
        $templateProcessor->setValue("licence", $licence);

        $certInspeccion = clearText($data["certInspeccion"]);
        $certInspeccion = ($certInspeccion == "") ? REEMPLAZAR_VACIOS : $certInspeccion;
        $templateProcessor->setValue("certInspeccion", $certInspeccion);

        $registerMine = clearText($data["registerMine"]);
        $registerMine = ($registerMine == "") ? REEMPLAZAR_VACIOS : $registerMine;
        $templateProcessor->setValue("registerMine", $registerMine);


        //CONTACTO
        $commercialContactName = clearText($data["commercialContactName"]);
        $commercialContactName = ($commercialContactName == "") ? REEMPLAZAR_VACIOS : $commercialContactName;
        $templateProcessor->setValue("commercialContactName", $commercialContactName);

        $fecha_emision_evaluacion = clearText($data["fecha_emision_evaluacion"]);
        $fecha_emision_evaluacion = ($fecha_emision_evaluacion == "") ? REEMPLAZAR_VACIOS : $fecha_emision_evaluacion;
        $templateProcessor->setValue("FECHA_EMISION_EVALUACION", $fecha_emision_evaluacion);

        $cliente_businessName = clearText($data["cliente_businessName"]);
        $cliente_businessName = ($cliente_businessName == "") ? REEMPLAZAR_VACIOS : $cliente_businessName;
        $templateProcessor->setValue("cliente_businessName", $cliente_businessName); //

        $templateProcessor->setValue("ENTE_DEFINIDOR_ESTANDARES", $cliente_businessName);

        $scope = clearText($data["scope"]);
        $scope = ($scope == "") ? REEMPLAZAR_VACIOS : $scope;
        $templateProcessor->setValue("scope", $scope);

        $retentionIgv = clearText($data["retentionIgv"]);
        $retentionIgv = ($retentionIgv == "") ? REEMPLAZAR_VACIOS : $retentionIgv;
        $templateProcessor->setValue("RETENCION_IGV", $retentionIgv);


        //PRODUCTOS:
        $rcc = clearText($data["rcc"]);
        $rcc = ($rcc == "") ? REEMPLAZAR_VACIOS : $rcc;
        $templateProcessor->setValue("rcc", $rcc);

        $negativo_sunat = clearText($data["negativo_sunat"]);
        $negativo_sunat = ($negativo_sunat == "") ? REEMPLAZAR_VACIOS : $negativo_sunat;
        $templateProcessor->setValue("negativo_sunat", $negativo_sunat);

        $cuentas_cerradas = clearText($data["cuentas_cerradas"]);
        $cuentas_cerradas = ($cuentas_cerradas == "") ? REEMPLAZAR_VACIOS : $cuentas_cerradas;
        $templateProcessor->setValue("cuentas_cerradas", $cuentas_cerradas);

        $clasif_crediticia = clearText($data["clasif_crediticia"]);
        $clasif_crediticia = ($clasif_crediticia == "") ? REEMPLAZAR_VACIOS : $clasif_crediticia;
        $templateProcessor->setValue("clasif_crediticia", $clasif_crediticia);

        $rectifica = clearText($data["rectifica"]);
        $rectifica = ($rectifica == "") ? REEMPLAZAR_VACIOS : $rectifica;
        $templateProcessor->setValue("rectifica", $rectifica);

        $comercio_exterior = clearText($data["comercio_exterior"]);
        $comercio_exterior = ($comercio_exterior == "") ? REEMPLAZAR_VACIOS : $comercio_exterior;
        $templateProcessor->setValue("comercio_exterior", $comercio_exterior);

        $boletines = clearText($data["boletines"]);
        $boletines = ($boletines == "") ? REEMPLAZAR_VACIOS : $boletines;
        $templateProcessor->setValue("boletines", $boletines);

        $deuda_previsional = clearText($data["deuda_previsional"]);
        $deuda_previsional = ($deuda_previsional == "") ? REEMPLAZAR_VACIOS : $deuda_previsional;
        $templateProcessor->setValue("deuda_previsional", $deuda_previsional);

        $consolidado = clearText($data["consolidado"]);
        $consolidado = ($consolidado == "") ? REEMPLAZAR_VACIOS : $consolidado;
        $templateProcessor->setValue("consolidad_morisidad", $consolidado);

        $templateProcessor->saveAs(PUBLIC_RESOURCES_INFORMES . $codigo_informe . INFORME_PARTE_UNO);
    endforeach;
else:
    exit("No hay datos suficientes para generar el informe");
endif;


$sql = "SELECT  crm_checklist.checkID,crm_checklist.precheckID FROM crm_checklist  INNER JOIN crm_check_homo on (crm_check_homo.checkID=crm_checklist.checkID)  WHERE 1=1 and crm_checklist.state=1  and crm_check_homo.homologacionID = " . $homologacionID . "  ORDER BY crm_checklist.checkID ASC";
$result = $database->getResult($sql);
if ($result["num_rows"] > 0):
    $checkID = "";
    $precheckID = "";
    foreach ($result["data"] as $data):
        $checkID.= "," . $data["checkID"];
        $precheckID.= "," . $data["precheckID"];
    endforeach;
    $checkID = substr($checkID, 1);
    $precheckID = substr($precheckID, 1);
    $whereIn = array();
    array_push($whereIn, ["checkID" => $checkID, "precheckID" => $precheckID]);
    $sql = "SELECT *,
	(SELECT scoreRes FROM crm_general_homo WHERE checkID = crm_checklist.checkID and homologacionID = " . $homologacionID . ") scoreRes,
	(SELECT scoreAcu FROM crm_general_homo WHERE checkID = crm_checklist.checkID and homologacionID = " . $homologacionID . ") scoreAcu,
	(SELECT observation FROM crm_general_homo WHERE homologacionID = " . $homologacionID . " AND checkID= crm_checklist.checkID) observation
	FROM crm_checklist 
	WHERE 1=1
	and precheckID = 0 
	and state=1 
	and checkID in (" . $precheckID . ")
	ORDER BY checkID ASC";
    $result = $database->getResult($sql);
    //'breakType' => 'oddPage',
    //'breakType' => 'continuous',
    $margenes = array('marginLeft' => 600, 'marginRight' => 600, 'marginTop' => 1000, 'marginBottom' => 2700);
    if ($result["num_rows"] > 0):
        $phpWord = new \PhpOffice\PhpWord\PhpWord();
        $h0 = 'h0';
        $h1 = 'h1';
        $h2 = 'h2';
        $h3 = 'h3';
        $h4 = 'h4';
        $phpWord->addFontStyle($h0, array('name' => 'Arial', 'size' => 13, 'color' => 'CC0000', 'bold' => true, 'align' => 'center', 'marginLeft' => 45));
        $phpWord->addFontStyle($h1, array('name' => 'Arial', 'size' => 8, 'color' => '1B2232', 'bold' => true, 'align' => 'center'));
        $phpWord->addFontStyle($h2, array('name' => 'Arial', 'size' => 8, 'color' => '1B2232', 'bold' => false, 'align' => 'center'));
        $phpWord->addFontStyle($h3, array('name' => 'Arial', 'size' => 8, 'color' => '1B2232', 'bold' => false, 'align' => 'center'));
        $phpWord->addFontStyle($h4, array('name' => 'Arial', 'size' => 8, 'color' => '1B2232', 'bold' => false, 'align' => 'center'));
        $styleTable = array(
            'borderSize' => 2,
            'borderColor' => $GLOBALS["GRIS"],
            'cellMarginRight' => 2,
            'cellMarginTop' => 0,
            'cellMarginBottom' => 0,
            'cellMarginLeft' => 1,
            'name' => 'Arial',
            'size' => 8,
            'align' => 'center',
            'width' => 100 * 50
        );
        $styleFirstRow = array(
            'borderBottomSize' => 18,
            'borderBottomColor' => '00000',
            'cellMarginRight' => 10,
            'cellMarginTop' => 0,
            'cellMarginBottom' => 0,
            'cellMarginLeft' => 10
        );

        $phpWord->addTableStyle('tablacentrada', $styleTable, $styleFirstRow);
        $whereIn_ = array_shift($whereIn);
        $prg_categorias = $result["data"];
        foreach ($result["data"] as $i => $data):
            //se recorren las categorías
            $section = $phpWord->addSection($margenes);
            $scoreRes_subtotal = $data["scoreRes"];
            $scoreAcu_subtotal = $data["scoreAcu"];
            $title_subtotal = "Subtotal "; //$data["title"];
            $comentarios = $data["observation"];

            if ($i == 0) {
                $section->addText("5. FORMULARIO DE EVALUACIÓN", $h0);
                $section->addTextBreak();
            }
            $section->addText(clearText($GLOBALS["categorias"][$i] . $data["title"], true), $h0);
            $sql = "select 
			crm_checklist.checkID,
			crm_checklist.precheckID,
			crm_checklist.formID,
			crm_checklist.typeCheck,
			crm_checklist.title,
			crm_checklist.question1,
			crm_checklist.question2,
			crm_checklist.question3,
			crm_checklist.question4,
			crm_checklist.question5,
			crm_checklist.text1,
			crm_checklist.text2,
			crm_checklist.text3,
			crm_checklist.text4,
			crm_checklist.text5,
			crm_checklist.score,
			crm_checklist.numScore,
			crm_checklist.information,
			crm_checklist.registerDate,
			crm_checklist.state,
			crm_check_homo.checkHomoID,
			crm_check_homo.homologacionID,
			crm_check_homo.response1,
			crm_check_homo.response2,
			crm_check_homo.response3,
			crm_check_homo.response4,
			crm_check_homo.response5,
			crm_check_homo.registerDate,
			crm_check_homo.registerUpdate,
			crm_check_homo.score as final_score
			FROM crm_checklist
			INNER JOIN crm_check_homo ON crm_checklist.checkID = crm_check_homo.checkID
			WHERE 1=1
			AND crm_checklist.precheckID = " . $data["checkID"] . "
			AND crm_check_homo.homologacionID = " . $homologacionID . "
			AND crm_checklist.state=1
			ORDER BY crm_checklist.checkID asc";
            $result = $database->getResult($sql);
            $tree = 0;
            $table = $section->addTable('tablacentrada');
            $table->addRow();
            $table->addCell(1000 + (5 * ANCHO_ALTERNATIVAS) + ANCHO_PUNTUACION, array('gridSpan' => 7, 'borderLeftSize' => 0, 'borderRightSize' => 0, 'borderTopSize' => 0, 'borderBottomSize' => 0, 'borderColor' => $GLOBALS["BLANCO"], 'valign' => 'center'))
                    ->addText(SEPARADOR);
            if ($result["num_rows"] > 0):
                foreach ($result["data"] as $y => $data):
                    //se recorren las preguntas y cabeceras
                    $alternativas = getAlternativas($data);
                    $title = clearText($data["title"]);
                    //Gean: Victor, debe quitar de la base de datos aquellas preguntas 21 que no tiene titutlo de pregunta/alternativa
                    //typeCheck, toma valores entre 1 y 2
                    $isleyenda = false;
                    $rowspan = false;
                    $score_rowspan = 0;
                    if ($data["typeCheck"] == PREGUNTA_FINAL)://2.

                        if ($alternativas == 0 && $y > 0): //leyendas informativas
                            $isleyenda = true;
                            $align = 'left';
                            $italic = false;
                            $color = $GLOBALS["NEGRO"];

                        elseif ($alternativas == 0 && $y == 0): //simular cabeceras de categoría
                            $align = 'left';
                            $italic = true;
                            $color = $GLOBALS["AZUL"]; //AZUL_MARINO
                        else:
                            $questions = array($data["question1"], $data["question2"], $data["question3"], $data["question4"], $data["question5"]);

                            if (in_array(PREGUNTA_CABECERA, $questions))://artificio para pintar las cabeceras
                                //imprimir espacio en blanco porque no tiene cabecera y se tiene categoría previa
                                /* $table->addRow();
                                  $table->addCell(9000+ANCHO_PUNTUACION, array('gridSpan' => 7, 'borderLeftSize' => 0,'borderRightSize' => 0,'borderTopSize' => 0, 'borderBottomSize' => 0, 'borderColor'=>$GLOBALS["BLANCO"],'valign'=>'center'))
                                  ->addText(SEPARADOR);//qwe "--separador de subpregunta--". */
                                $align = PREGUNTA_ALINEADA; //'right';
                                $italic = false;
                                $color = $GLOBALS["ROJO"];
                            else://otros titulos ya en la respuesta
                                $align = PREGUNTA_ALINEADA; //'right';
                                $italic = false;
                                $color = $GLOBALS["NEGRO"];
                            endif;
                        endif;
                        $title = ($isleyenda) ? "**" . ucwords(mb_strtolower($title)) : $title; //leyenda
                        $cellColSpan = array('gridSpan' => 6 - $alternativas, 'valign' => 'center');

                        $iscabecera = true;
                        $mostrar_no_aplica = false; //se declara fuera ya que basta que uno sea no aplica para cambiar la respuesta del porcentaje
                        if ($alternativas > 0):
                            $bgColor = $GLOBALS["BLANCO"];
                            if ($isleyenda):
                            /* $table->addRow();
                              $table->addCell(9000+ANCHO_PUNTUACION, array('gridSpan' => 7, 'borderLeftSize' => 0,'borderRightSize' => 0,'borderTopSize' => 0, 'borderBottomSize' => 0, 'borderColor'=>$GLOBALS["BLANCO"],'valign'=>'center'))->addText(SEPARADOR); */
                            endif;
                            $table->addRow();
                            $table->addCell(1000 + ((5 - $alternativas) * ANCHO_ALTERNATIVAS), $cellColSpan)
                                    ->addText($title, ['name' => 'Arial', 'size' => 8, 'italic' => $italic, 'color' => $color, 'bold' => false, 'valign' => 'center'], ['align' => $align]
                            );


                            //tiene alternativas
                            //for($i=$alternativas;$i>0;$i--):
                            $mostrar_no_aplica = false;
                            for ($i = 1; $i <= $alternativas; $i++):
                                $value = "";
                                if ($data["question" . $i] == PREGUNTA_CABECERA):
                                    $value = $data["text" . $i];
                                elseif ($data["question" . $i] == PREGUNTA_CERRADA_SIMPLE):
                                    $iscabecera = false;
                                    if (array_key_exists($data["response" . $i], $GLOBALS["RESPUESTAS"][$data["question" . $i]])):
                                        $value = $GLOBALS["RESPUESTAS"][$data["question" . $i]][$data["response" . $i]];
                                    endif;
                                elseif ($data["question" . $i] == PREGUNTA_CERRADA_COMPLEJA):
                                    $iscabecera = false;

                                    if (array_key_exists($data["response" . $i], $GLOBALS["RESPUESTAS"][$data["question" . $i]])):
                                        $value = $GLOBALS["RESPUESTAS"][$data["question" . $i]][$data["response" . $i]];
                                        if ($data["response" . $i] == 3) {
                                            $mostrar_no_aplica = true;
                                        }
                                    endif;

                                else:
                                    $iscabecera = false;
                                    $value = $data["response" . $i];
                                endif;
                                $value = clearText($value);
                                $value = ($value == "") ? REEMPLAZAR_VACIOS : $value;
                                $table->addCell($GLOBALS["WIDTH_FOR_QUESTION"][$data["question" . $i]], ['valign' => 'center'])
                                        ->addText($value, ["size" => 8, "color" => $color], ["align" => "center"]);
                            endfor;
                        else:
                            $bgColor = $GLOBALS["BLANCO"];
                            if ($isleyenda):
                            /* $table->addRow();
                              $table->addCell(9000+ANCHO_PUNTUACION, array('gridSpan' => 7, 'borderLeftSize' => 0,'borderRightSize' => 0,'borderTopSize' => 0, 'borderBottomSize' => 0, 'borderColor'=>$GLOBALS["BLANCO"],'valign'=>'center'))->addText(SEPARADOR); */
                            endif;

                            //deja espacio par el puntaje. no tiene alternativas, eso quiere decir que pintó el título con un colspan de 6.
                            $table->addRow();
                            $table->addCell(1000 + (5 * ANCHO_ALTERNATIVAS), $cellColSpan)
                                    ->addText($title, ['name' => 'Arial', 'size' => 8, 'italic' => $italic, 'color' => $color, 'bold' => false, 'valign' => 'center'], ['align' => $align]
                            );

                        endif;

                        if ($data["score"] == PUNTAJE_CALIFICADO && $isleyenda != true && $iscabecera != OCULTAR_SCORE_CABECERAS):
                            $ss = ($mostrar_no_aplica) ? $GLOBALS["RESPUESTAS"][7][3] : getScore($data["final_score"], $data["numScore"]);
                            $table->addCell(ANCHO_PUNTUACION, array('valign' => 'center', 'bgColor' => $bgColor))
                                    ->addText($ss, ['size' => 8, 'color' => $GLOBALS["NEGRO"], 'bold' => true], ['align' => 'center']);
                        elseif ($data["score"] == PUNTAJE_ACUMULADO && $isleyenda != true && $iscabecera != OCULTAR_SCORE_CABECERAS):
                            $ss = ($mostrar_no_aplica) ? $GLOBALS["RESPUESTAS"][7][3] : INFORMATIVO;
                            $table->addCell(ANCHO_PUNTUACION, array('valign' => 'center', 'bgColor' => $bgColor))
                                    ->addText($ss, ['size' => 8, 'color' => $GLOBALS["NEGRO"], 'bold' => true], ['align' => 'center']);
                        else:
                            $ss = ($data["score"] == PUNTAJE_ACUMULADO) ? INFORMATIVO : getScore($data["final_score"], $data["numScore"]);
                            $table->addCell(ANCHO_PUNTUACION, array('valign' => 'center', 'bgColor' => $bgColor))
                                    ->addText($ss, ['size' => 8, 'color' => $GLOBALS["NEGRO"], 'bold' => true], ['align' => 'center']);
                        endif;



                    elseif ($data["typeCheck"] == PREGUNTA_MULTIPLE)://OK
                        //no aplica el uso de leyenda=true
                        $somelikeu = 2;
                        $title = ($title == "") ? "Título para el contenedor" : $title;
                        if ($alternativas == 0):

                            //hacer un espacio para imprimir entre contenedores
                            $table->addRow();
                            $table->addCell(1000 + (5 * ANCHO_ALTERNATIVAS) + ANCHO_PUNTUACION, array('gridSpan' => 7, 'borderLeftSize' => 0, 'borderRightSize' => 0, 'borderTopSize' => 0, 'borderBottomSize' => 0, 'borderColor' => $GLOBALS["BLANCO"], 'valign' => 'center'))
                                    ->addText(SEPARADOR);

                            //imprimir cabeceras pero verificar si tiene puntuación 
                            if ($data["score"] == PUNTAJE_CALIFICADO)://1
                                $rowspan = true;
                                $score_rowspan = getScore($data["final_score"], $data["numScore"]);
                                $table->addRow();
                                $table->addCell(1000 + (5 * ANCHO_ALTERNATIVAS), array('vMerge' => 'restart', 'valign' => 'center', 'gridSpan' => 6, 'bgColor' => $GLOBALS["GRIS"]))
                                        ->addText($title, array('name' => 'Arial', 'size' => 8, 'color' => $GLOBALS["NEGRO"], 'bold' => true), ['align' => 'left']
                                );

                                //$table->addCell(ANCHO_PUNTUACION,$cellRowContinue)
                                $table->addCell(ANCHO_PUNTUACION, array('vMerge' => 'restart', "valign" => "center", 'bgColor' => $GLOBALS["BLANCO"]))
                                        ->addText($score_rowspan, ['size' => 8, 'color' => $GLOBALS["NEGRO"], 'bold' => true], ['align' => 'center']); //mix score 
                            elseif ($data["score"] == PUNTAJE_ACUMULADO): //2
                                $rowspan = false;
                                $score_rowspan = INFORMATIVO;
                                $table->addRow();
                                $table->addCell(1000 + (5 * ANCHO_ALTERNATIVAS), array('gridSpan' => 6, 'valign' => 'center', 'bgColor' => $GLOBALS["GRIS"]))
                                        ->addText($title, array('name' => 'Arial', 'size' => 8, 'color' => $GLOBALS["NEGRO"], 'bold' => true), ['align' => 'left']
                                ); //no mostrar informativo en los primeros contenedores en gris según solicitud de victor
                                $table->addCell(ANCHO_PUNTUACION, array('vMerge' => 'restart', "valign" => "center", 'bgColor' => $GLOBALS["BLANCO"]))
                                        ->addText($score_rowspan, ['size' => 8, 'color' => $GLOBALS["NEGRO"], 'bold' => true], ['align' => 'center']); 
                            else:
                                //score inesperado..
                                $rowspan = true;
                                $score_rowspan = getScore($data["final_score"], $data["numScore"]);
                                $table->addRow();
                                $table->addCell(1000 + (5 * ANCHO_ALTERNATIVAS), array('gridSpan' => 6, 'valign' => 'center', 'bgColor' => $GLOBALS["GRIS"]))
                                        ->addText($title, array('name' => 'Arial', 'size' => 8, 'color' => $GLOBALS["NEGRO"], 'bold' => true), ['align' => 'left']
                                );

                                $table->addCell(ANCHO_PUNTUACION, array('vMerge' => 'restart', 'valign' => 'center', 'bgColor' => $GLOBALS["BLANCO"]))
                                        ->addText($score_rowspan, ['size' => 8, 'color' => $GLOBALS["NEGRO"], 'bold' => true], ['align' => 'center']); //mix score	

                            endif;
                            //traer subpreguntas 
                            getCaffe($homologacionID, $data["checkID"], $somelikeu, $table, $section, $rowspan, $score_rowspan);
                        else:
                            exit("Advertencia, imposible continuar, hay inconsistencia en la definición de su formulario <b>" . $title . "</b>");
                        endif;

                    else:
                        exit("Advertencia, valor no permitido para typeCheck. valor=" . $data["typeCheck"]);
                    endif;
                endforeach;
            endif;

            //ESCRIBIR SUBTOTAL
            $score_subtotal = getScore($scoreRes_subtotal, $scoreAcu_subtotal);
            $score_subtotal = ($score_subtotal == "") ? "INFORMATIVO" : $score_subtotal;

            $table->addRow();

            $table->addCell(8000, array('gridSpan' => 6, 'borderLeftSize' => 0, 'borderRightSize' => 0, 'borderTopSize' => 0, 'borderBottomSize' => 0, 'borderColor' => $GLOBALS["BLANCO"], 'valign' => 'center'))
                    ->addText(clearText($title_subtotal), ['size' => 8, "bold" => true, 'color' => $GLOBALS["NEGRO"]], ["align" => "right"]);

            $table->addCell(1000, ['valign' => 'center'])
                    ->addText($score_subtotal, ['size' => 8, "bold" => true, 'color' => $GLOBALS["NEGRO"]], ["align" => "center"]);





            //ESCRIBIR LOS COMENTARIOS
            $paragraphStyleName = 'P-Style';
            $phpWord->addParagraphStyle($paragraphStyleName, array('spaceAfter' => 70, 'size' => 8, 'color' => $GLOBALS["NEGRO"], 'align' => 'both'));
            $predefinedMultilevelStyle = array('listType' => \PhpOffice\PhpWord\Style\ListItem::TYPE_BULLET_FILLED);
            $section->addTextBreak();
            $tableComentarios = $table;
            $tableComentarios->addRow();
            $tableComentarios->addCell(1000 + (5 * ANCHO_ALTERNATIVAS) + ANCHO_PUNTUACION, array('gridSpan' => 7, 'borderLeftSize' => 0, 'borderRightSize' => 0, 'borderTopSize' => 0, 'borderBottomSize' => 0, 'borderColor' => $GLOBALS["BLANCO"], 'valign' => 'center'))->addText(SEPARADOR);
            $tableComentarios->addRow();
            $tableComentarios->addCell(1000 + (5 * ANCHO_ALTERNATIVAS) + ANCHO_PUNTUACION, array('gridSpan' => 7, 'valign' => 'center', 'bgColor' => $GLOBALS["GRIS"]))->addText(clearText("Comentarios:"), ["italic" => false, "bold" => true, "size" => 8, 'color' => $GLOBALS["NEGRO"]], ["align" => "both"]);
            $prg_comentarios = explode("\n", $comentarios);

            foreach ($prg_comentarios as $key => $comentario):
                $comentario = str_replace(array("- ", "-	", "-		", "•", "•	"), array("", "", "", "", ""), $comentario);
                $comentario = clearText($comentario);
                if ($comentario != "" && preg_match('/\s/', $comentario))://evitar imprimir espacios en blanco
                    if (mb_substr($comentario, -1) == ":"):
                        $tableComentarios->addRow();
                        $tableComentarios->addCell(1000 + (5 * ANCHO_ALTERNATIVAS) + ANCHO_PUNTUACION, array('gridSpan' => 7, 'valign' => 'center',
                            'borderTopColor' => $GLOBALS["BLANCO"],
                            'borderLeftColor' => $GLOBALS["GRIS"],
                            'borderRightColor' => $GLOBALS["GRIS"],
                            'borderBottomColor' => $GLOBALS["BLANCO"],
                            'borderTopSize' => 0, 'borderLeftSize' => 1, 'borderRightSize' => 1, 'borderBottomSize' => 0
                        ))->addText($comentario, ["size" => 8, 'color' => $GLOBALS["NEGRO"]], []);
                    else:
                        $tableComentarios->addRow();
                        $tableComentarios->addCell(1000 + (5 * ANCHO_ALTERNATIVAS) + ANCHO_PUNTUACION, array('gridSpan' => 7, 'valign' => 'center',
                                    'borderTopColor' => $GLOBALS["BLANCO"],
                                    'borderLeftColor' => $GLOBALS["GRIS"],
                                    'borderRightColor' => $GLOBALS["GRIS"],
                                    'borderBottomColor' => $GLOBALS["BLANCO"],
                                    'borderTopSize' => 0, 'borderLeftSize' => 1, 'borderRightSize' => 1, 'borderBottomSize' => 0
                                ))
                                ->addListItem($comentario, 0, ["size" => 8, 'color' => $GLOBALS["NEGRO"]], $predefinedMultilevelStyle, $paragraphStyleName);
                    endif;

                endif;
            endforeach;
            $table->addRow();
            $table->addCell(1000 + (5 * ANCHO_ALTERNATIVAS) + ANCHO_PUNTUACION, array('gridSpan' => 7,
                'borderTopSize' => 1, 'borderLeftSize' => 0, 'borderRightSize' => 0, 'borderBottomSize' => 0,
                'borderTopColor' => $GLOBALS["GRIS"],
                'borderLeftColor' => $GLOBALS["BLANCO"],
                'borderRightColor' => $GLOBALS["BLANCO"],
                'borderBottomColor' => $GLOBALS["BLANCO"]))->addText(SEPARADOR); //quitar caja que va debajo del comentario
        endforeach;

        
        //ESCRIBIR CUADRO DE RESULTADOS
        $section = $phpWord->addSection($margenes);
        $section->addText('6. RESUMEN DE RESULTADOS', array('name' => 'Arial', 'size' => 13, 'color' => 'CC0000', 'bold' => true, 'valign' => 'center', 'marginLeft' => 45));
        $section->addTextBreak();
        $table = $section->addTable('tablacentrada');
        $table->addRow();

        $generalCell = array('borderSize' => 1, 'bgColor' => $GLOBALS["BLANCO"]);
        $selectedCell = array('borderSize' => 1, 'bgColor' => $GLOBALS["VERDE"]);
        $sql = "SELECT *,(select min(maximo) from crm_form_propuesta_nivel where propxformID=" . $propxformID . ") as minimo_aprobatorio FROM crm_form_propuesta_nivel where propxformID=" . $propxformID . " order by nombreNivel asc";

        $result = $database->getResult($sql);
        $result_conclusiones = $result;
        if ($result["num_rows"] > 0) {
            foreach ($result["data"] as $data) {

                $bgColor = ((int) $data["minimo"] <= (int) $puntajeFinalTotal && (int) $data["maximo"] >= (int) $puntajeFinalTotal) ? $GLOBALS["VERDE"] : $GLOBALS["BLANCO"];
                $table->addCell(1000 + ANCHO_PUNTUACION, array('borderSize' => 1, 'bgColor' => $bgColor))
                        ->addText("NIVEL " . $data["nombreNivel"], ['bold' => true, 'size' => 10, 'valign' => 'center', 'color' => $GLOBALS["NEGRO"]], ['align' => 'center']);
            }
            $section->addTextBreak();
        }

        $section->addTextBreak();
        $table = $section->addTable('tablacentrada');
        $table->addRow();
        $table->addCell(1000 + (5 * ANCHO_ALTERNATIVAS), ['name' => 'Arial', 'bgColor' => $GLOBALS["GRIS"], 'bold' => true, 'gridSpan' => 3, 'valign' => 'center'])
                ->addText("CUADRO DE PUNTAJES", ["bold" => true, 'size' => 10, 'color' => $GLOBALS["NEGRO"]], ["align" => "center"]);

        $table->addRow();
        $table->addCell(1000, ['name' => 'Arial', 'bgColor' => $GLOBALS["GRIS"], 'bold' => true, 'valign' => 'center'])
                ->addText("ITEM", ["bold" => true, 'size' => 10, 'color' => $GLOBALS["NEGRO"]], ["align" => "center"]);

        $table->addCell(6000, ['name' => 'Arial', 'bgColor' => $GLOBALS["GRIS"], 'bold' => true])
                ->addText("ÁREA EVALUADA", ["bold" => true, 'size' => 10, 'color' => $GLOBALS["NEGRO"]], ["align" => "left"]);

        $table->addCell(1000, ['name' => 'Arial', 'bgColor' => $GLOBALS["GRIS"], 'bold' => true])
                ->addText("CALIFICACIÓN %", ["bold" => true, 'size' => 10, 'color' => $GLOBALS["NEGRO"]], ["align" => "center"]);

        foreach ($prg_categorias as $i => $row_categoria):
            $table->addRow();
            $table->addCell(1000, ['valign' => 'center'])
                    ->addText(str_pad( ++$i, 2, "0", STR_PAD_LEFT), ["bold" => false, 'color' => $GLOBALS["NEGRO"]], ["align" => "center"]);

            $table->addCell(6000, ['valign' => 'center'])
                    ->addText(clearText($row_categoria["title"]), ["bold" => false, 'color' => $GLOBALS["NEGRO"]], ["align" => "left"]);

            $score = getScore($row_categoria["scoreRes"], $row_categoria["scoreAcu"]);
            $score = ($score == "") ? "INFORMATIVO" : $score;

            $table->addCell(1000, ['valign' => 'center'])
                    ->addText($score, ["bold" => false, 'color' => $GLOBALS["NEGRO"]], ["align" => "center"]);
        endforeach;
        $table->addRow();

        $table->addCell(7000, ['name' => 'Arial', 'size' => 20, 'bgColor' => $GLOBALS["GRIS"], 'bold' => true, 'valign' => 'center', 'gridSpan' => 2])
                ->addText("TOTAL", ["bold" => true, 'color' => $GLOBALS["NEGRO"]], ["align" => "center"]);
        $table->addCell(2000, ['name' => 'Arial', 'size' => 20, 'bgColor' => $GLOBALS["GRIS"], 'bold' => true, 'valign' => 'center'])
                ->addText(getScore($puntajeFinalTotal, 100), ["bold" => true, 'color' => $GLOBALS["NEGRO"]], ["align" => "center"]);
        $section->addTextBreak();



        $section->addText('7. CONCLUSIONES', array('name' => 'Arial', 'size' => 13, 'color' => 'CC0000', 'bold' => true, 'valign' => 'center', 'marginLeft' => 45));
        $section->addTextBreak();

        $conclusiones = "";
        if ($result_conclusiones["num_rows"] > 0) {
            $conclusiones = ((int) $result_conclusiones["data"][0]["minimo_aprobatorio"] < (int) $puntajeFinalTotal) ? "ha aprobado el proceso de homologación. Recomendándose emitir el Certificado de Proveedor." : "no ha aprobado el proceso de homologación.";
        } else {
            echo "Por favor corregir la base de datos. La tabla crm_form_propuesta_nivel no ha encontrado resultados para propxformID = " . $propxformID;
        }


        $section->addText("La empresa " . $businessName . " ha alcanzado el " . getScore($puntajeFinalTotal, 100) . " de cumplimiento total, correspondiéndole la calificación en el NIVEL " . $nivel . ". Por lo tanto BV considera que la empresa " . $businessName . " " . $conclusiones, [], ['align' => 'both']);
        $section->addTextBreak();
        $table = $section->addTable('tablacentrada');
        $table->addRow();
        $table->addCell(1000, ['valign' => 'center', 'bgColor' => $GLOBALS["GRIS"]])
                ->addText("NIVEL", ["bold" => true, "size" => 10, 'color' => $GLOBALS["NEGRO"]], ["align" => "center"]);
        $table->addCell(4000, ['valign' => 'center', 'bgColor' => $GLOBALS["GRIS"]])
                ->addText("RANGO %", ["bold" => true, "size" => 10, 'color' => $GLOBALS["NEGRO"]], ["align" => "center"]);

        if ($result_conclusiones["num_rows"] > 0) {
            foreach ($result_conclusiones["data"] as $data) {

                $bgColor = ((int) $data["minimo"] <= (int) $puntajeFinalTotal && (int) $data["maximo"] >= (int) $puntajeFinalTotal) ? $GLOBALS["VERDE"] : $GLOBALS["BLANCO"];

                $info = ($data["estado"] == 2) ? "(No certificada)" : "";
                $table->addRow();
                $table->addCell(1000, ['valign' => 'center', 'bgColor' => $bgColor])
                        ->addText("NIVEL " . $data["nombreNivel"], ["bold" => false, "size" => 10], ["align" => "center"]);
                $table->addCell(4000, ['valign' => 'center', 'bgColor' => $bgColor])
                        ->addText("[" . $data["minimo"] . " - " . $data["maximo"] . "] %" . $info, ["bold" => false, "size" => 10], ["align" => "center"]);
            }
            $section->addTextBreak();
        }



        $section->addText('8. OBSERVACIONES', array('name' => 'Arial', 'size' => 13, 'color' => 'CC0000', 'bold' => true, 'valign' => 'center', 'marginLeft' => 45));
        $section->addTextBreak();
        $section->addText($observaciones, ["size" => 11], ['align' => 'both']);
        $section->addTextBreak(2);


        $section->addText('9. RECOMENDACIONES', array('name' => 'Arial', 'size' => 13, 'color' => 'CC0000', 'bold' => true, 'valign' => 'center', 'marginLeft' => 45));
        $section->addTextBreak();

        $prg_recomendaciones = explode("- ", $recomendaciones);
        foreach ($prg_recomendaciones as $key => $recomendacion):
            $recomendacion = clearText($recomendacion);
            if ($recomendacion != ""):
                if (mb_strtolower($recomendacion) == "fortalezas"):
                    $section->addText("Fortalezas", ["bold" => true]);
                elseif (mb_strtolower(substr($recomendacion, -11)) == "debilidades"):
                    $section->addText("Debilidades", ["bold" => true]);
                else:
                    $section->addListItem($recomendacion, 0, [], $predefinedMultilevelStyle, $paragraphStyleName);
                endif;

            endif;
        endforeach;

        $section = $phpWord->addSection($margenes);
        $section->addText('10. FOTOGRAFÍAS', array('name' => 'Arial', 'size' => 13, 'color' => 'CC0000', 'bold' => true, 'valign' => 'center', 'marginLeft' => 45));
        $section->addTextBreak();

        $sql = "select * from crm_photo_homo where homologacionID=" . $homologacionID;
        $result = $database->getResult($sql);
        if ($result["num_rows"] > 0):
            $section->createHeader();
            $imageOptions = ['align' => 'center', 'width' => IMAGE_WIDTH, 'height' => IMAGE_HEIGHT];
            $table = $section->addTable('tablacentrada');
            $cellFotos = array('borderLeftSize' => 0, 'borderRightSize' => 0, 'borderTopSize' => 0, 'borderBottomSize' => 0, 'borderColor' => $GLOBALS["BLANCO"], 'valign' => 'center');
            foreach ($result["data"] as $data) {

                $table->addRow();
                if ($data["photo1"] != "" && file_exists(PUBLIC_RESOURCES_FOTOS . $data["photo1"]) && getimagesize(PUBLIC_RESOURCES_FOTOS . $data["photo1"]) > 0) {
                    $table->addCell(IMAGE_CELL_WIDTH, $cellFotos)->addImage(PUBLIC_RESOURCES_FOTOS . $data["photo1"], $imageOptions);
                }

                $table->addCell(ANCHO_PUNTUACION, $cellFotos)->addText(SEPARADOR); //separador entre columnas

                if ($data["photo2"] != "" && file_exists(PUBLIC_RESOURCES_FOTOS . $data["photo2"]) && getimagesize(PUBLIC_RESOURCES_FOTOS . $data["photo2"]) > 0) {
                    $table->addCell(IMAGE_CELL_WIDTH, $cellFotos)->addImage(PUBLIC_RESOURCES_FOTOS . $data["photo2"], $imageOptions);
                }
                $table->addRow();
                if ($data["photo1"] != "" && file_exists(PUBLIC_RESOURCES_FOTOS . $data["photo1"]) && getimagesize(PUBLIC_RESOURCES_FOTOS . $data["photo1"]) > 0) {
                    $table->addCell(IMAGE_CELL_WIDTH, $cellFotos)->addText($data["description1"], ['size' => 8, 'bold' => true, 'color' => $GLOBALS["NEGRO"]], ['align' => 'center']);
                }
                $table->addCell(ANCHO_PUNTUACION, $cellFotos)->addText(SEPARADOR); //separador entre columnas
                if ($data["photo2"] != "" && file_exists(PUBLIC_RESOURCES_FOTOS . $data["photo2"]) && getimagesize(PUBLIC_RESOURCES_FOTOS . $data["photo2"]) > 0) {
                    $table->addCell(IMAGE_CELL_WIDTH, $cellFotos)->addText($data["description2"], ['size' => 8, 'bold' => true, 'color' => $GLOBALS["NEGRO"]], ['align' => 'center']);
                }
                $table->addRow();
                if ($data["photo3"] != "" && file_exists(PUBLIC_RESOURCES_FOTOS . $data["photo3"]) && getimagesize(PUBLIC_RESOURCES_FOTOS . $data["photo3"]) > 0) {
                    $table->addCell(IMAGE_CELL_WIDTH, $cellFotos)->addImage(PUBLIC_RESOURCES_FOTOS . $data["photo3"], $imageOptions);
                }

                $table->addCell(ANCHO_PUNTUACION, $cellFotos)->addText(SEPARADOR); //separador entre columnas

                if ($data["photo4"] != "" && file_exists(PUBLIC_RESOURCES_FOTOS . $data["photo4"]) && getimagesize(PUBLIC_RESOURCES_FOTOS . $data["photo4"]) > 0) {
                    $table->addCell(IMAGE_CELL_WIDTH, $cellFotos)->addImage(PUBLIC_RESOURCES_FOTOS . $data["photo4"], $imageOptions);
                }
                $table->addRow();
                if ($data["photo3"] != "" && file_exists(PUBLIC_RESOURCES_FOTOS . $data["photo3"]) && getimagesize(PUBLIC_RESOURCES_FOTOS . $data["photo3"]) > 0) {
                    $table->addCell(IMAGE_CELL_WIDTH, $cellFotos)->addText($data["description3"], ['size' => 8, 'bold' => true, 'color' => $GLOBALS["NEGRO"]], ['align' => 'center']);
                }
                $table->addCell(ANCHO_PUNTUACION, $cellFotos)->addText(SEPARADOR); //separador entre columnas
                if ($data["photo4"] != "" && file_exists(PUBLIC_RESOURCES_FOTOS . $data["photo4"]) && getimagesize(PUBLIC_RESOURCES_FOTOS . $data["photo4"]) > 0) {
                    $table->addCell(IMAGE_CELL_WIDTH, $cellFotos)->addText($data["description4"], ['size' => 8, 'bold' => true, 'color' => $GLOBALS["NEGRO"]], ['align' => 'center']);
                }
                $table->addRow();
                if ($data["photo5"] != "" && file_exists(PUBLIC_RESOURCES_FOTOS . $data["photo5"]) && getimagesize(PUBLIC_RESOURCES_FOTOS . $data["photo5"]) > 0) {
                    $table->addCell(IMAGE_CELL_WIDTH, $cellFotos)->addImage(PUBLIC_RESOURCES_FOTOS . $data["photo5"], $imageOptions);
                }

                $table->addCell(ANCHO_PUNTUACION, $cellFotos)->addText(SEPARADOR); //separador entre columnas

                if ($data["photo6"] != "" && file_exists(PUBLIC_RESOURCES_FOTOS . $data["photo6"]) && getimagesize(PUBLIC_RESOURCES_FOTOS . $data["photo6"]) > 0) {
                    $table->addCell(IMAGE_CELL_WIDTH, $cellFotos)->addImage(PUBLIC_RESOURCES_FOTOS . $data["photo6"], $imageOptions);
                }
                $table->addRow();
                if ($data["photo5"] != "" && file_exists(PUBLIC_RESOURCES_FOTOS . $data["photo5"]) && getimagesize(PUBLIC_RESOURCES_FOTOS . $data["photo5"]) > 0) {
                    $table->addCell(IMAGE_CELL_WIDTH, $cellFotos)->addText($data["description5"], ['size' => 8, 'bold' => true, 'color' => $GLOBALS["NEGRO"]], ['align' => 'center']);
                }
                $table->addCell(ANCHO_PUNTUACION, $cellFotos)->addText(SEPARADOR); //separador entre columnas
                if ($data["photo6"] != "" && file_exists(PUBLIC_RESOURCES_FOTOS . $data["photo6"]) && getimagesize(PUBLIC_RESOURCES_FOTOS . $data["photo6"]) > 0) {
                    $table->addCell(IMAGE_CELL_WIDTH, $cellFotos)->addText($data["description6"], ['size' => 8, 'bold' => true, 'color' => $GLOBALS["NEGRO"]], ['align' => 'center']);
                }
                $table->addRow();
                if ($data["photo7"] != "" && file_exists(PUBLIC_RESOURCES_FOTOS . $data["photo7"]) && getimagesize(PUBLIC_RESOURCES_FOTOS . $data["photo7"]) > 0) {
                    $table->addCell(IMAGE_CELL_WIDTH, $cellFotos)->addImage(PUBLIC_RESOURCES_FOTOS . $data["photo7"], $imageOptions);
                }

                $table->addCell(ANCHO_PUNTUACION, $cellFotos)->addText(SEPARADOR); //separador entre columnas

                if ($data["photo8"] != "" && file_exists(PUBLIC_RESOURCES_FOTOS . $data["photo8"]) && getimagesize(PUBLIC_RESOURCES_FOTOS . $data["photo8"]) > 0) {
                    $table->addCell(IMAGE_CELL_WIDTH, $cellFotos)->addImage(PUBLIC_RESOURCES_FOTOS . $data["photo8"], $imageOptions);
                }
                $table->addRow();
                if ($data["photo7"] != "" && file_exists(PUBLIC_RESOURCES_FOTOS . $data["photo7"]) && getimagesize(PUBLIC_RESOURCES_FOTOS . $data["photo7"]) > 0) {
                    $table->addCell(IMAGE_CELL_WIDTH, $cellFotos)->addText($data["description5"], ['size' => 8, 'bold' => true, 'color' => $GLOBALS["NEGRO"]], ['align' => 'center']);
                }
                $table->addCell(ANCHO_PUNTUACION, $cellFotos)->addText(SEPARADOR); //separador entre columnas
                if ($data["photo8"] != "" && file_exists(PUBLIC_RESOURCES_FOTOS . $data["photo8"]) && getimagesize(PUBLIC_RESOURCES_FOTOS . $data["photo8"]) > 0) {
                    $table->addCell(IMAGE_CELL_WIDTH, $cellFotos)->addText($data["description6"], ['size' => 8, 'bold' => true, 'color' => $GLOBALS["NEGRO"]], ['align' => 'center']);
                }
            }
        endif;

        $section->addText($auditedBy, ['bold' => true, "size" => 11, 'color' => $GLOBALS["NEGRO"]], ['align' => 'right']);
        $section->addText("Homologador", ['bold' => true, "size" => 11, 'color' => $GLOBALS["NEGRO"]], ['align' => 'right']);
        $section->addText("Bureau Veritas del Perú S.A.", ['bold' => true, "size" => 11, 'color' => $GLOBALS["NEGRO"]], ['align' => 'right']);

        $objWriter = \PhpOffice\PhpWord\IOFactory::createWriter($phpWord, 'Word2007');
        $objWriter->save(PUBLIC_RESOURCES_INFORMES . $codigo_informe . INFORME_PARTE_DOS);


    endif;
endif;


$word_informe_general = $codigo_informe . "_" . $homologacionID . "___informe_general.docx";
$pdf_informe_general = $codigo_informe . "_" . $homologacionID . "___informe_general.pdf";
$docxMerge = \Jupitern\Docx\DocxMerge::instance();
$docxMerge->addFiles([PUBLIC_RESOURCES_INFORMES . $codigo_informe . INFORME_PARTE_UNO, PUBLIC_RESOURCES_INFORMES . $codigo_informe . INFORME_PARTE_DOS]);
$docxMerge->save(PUBLIC_RESOURCES_INFORMES . $word_informe_general, true);

echo "<br><script>window.location.href='./resources/" . $word_informe_general . "'</script>";

function clearText($text = "", $upperCase = false) {
	if(DEBUG){echo "<font color='red'><br>Limpiar texto (1):<br></font>";
	var_dump($text);}
    $text = strip_tags($text);
    if(DEBUG){echo "<font color='red'><br>Limpiar texto (2):<br></font>";
    var_dump($text);}
    $text = htmlspecialchars_decode($text);
    if(DEBUG){echo "<font color='red'><br>Limpiar texto (3):<br></font>";
    var_dump($text);}
    $text = htmlspecialchars($text, ENT_COMPAT, 'UTF-8');
	if(DEBUG){echo "<font color='red'><br>Limpiar texto (4):<br></font>";
    var_dump($text);}
    $text = ($upperCase) ? mb_strtoupper($text) : $text;
    if(DEBUG){echo "<font color='red'><br>Limpiar texto (5):<br></font>";
    var_dump($text);}
    $text = str_replace(["/", "undefined"], ["-", ""], trim(stripslashes($text)));
    if(DEBUG){echo "<font color='red'><br>Limpiar texto (6):<br></font>";
    var_dump($text);
    echo "<hr>";}
    return $text;
}

function getScore($notaObtenida = 0, $notaMaxima = 0) {
    if ($notaMaxima == 0):
        $score = INFORMATIVO; //artificio solicitado por victor por temas de definciión en su bd
    else:
        $score = ($notaObtenida + 0) / $notaMaxima * 100;
        $score = number_format($score, 2, ".", "0");
        $score = str_pad($score, 5, "0", STR_PAD_LEFT) . "%";
    endif;

    return $score;
}

function getHamburger($homologacionID, $checkID, $somelikeu, &$table, &$section, $rowspan = false, $score_rowspan = 0) {

    $cellRowSpan = array('vMerge' => 'restart', 'valign' => 'center', 'bgColor' => $GLOBALS["BLANCO"]);
    $cellRowContinue = array('vMerge' => 'continue', 'valign' => 'center', 'bgColor' => $GLOBALS["BLANCO"]);
    $cellHCentered = array('alignment' => \PhpOffice\PhpWord\SimpleType\Jc::CENTER);
    $cellVCentered = array('valign' => 'center');

    $sql = "select 
	crm_check_homo.checkHomoID,
	crm_check_homo.checkID,
	crm_check_homo.homologacionID,
	crm_check_homo.response1,
	crm_check_homo.response2,
	crm_check_homo.response3,
	crm_check_homo.response4,
	crm_check_homo.response5,
	crm_check_homo.registerDate,
	crm_check_homo.registerUpdate,
	crm_checklist.precheckID,
	crm_checklist.formID,
	crm_checklist.typeCheck,
	crm_checklist.title,
	crm_checklist.question1,
	crm_checklist.question2,
	crm_checklist.question3,
	crm_checklist.question4,
	crm_checklist.question5,
	crm_checklist.text1,
	crm_checklist.text2,
	crm_checklist.text3,
	crm_checklist.text4,
	crm_checklist.text5,
	crm_checklist.score,
	crm_checklist.numScore,
	crm_check_homo.score as final_score,
	crm_checklist.information,
	crm_checklist.state
	FROM crm_checklist
	INNER JOIN crm_check_homo ON crm_checklist.checkID = crm_check_homo.checkID
	WHERE 1=1
	AND crm_checklist.precheckID =" . $checkID . " 
	AND crm_check_homo.homologacionID = " . $homologacionID . "
	ORDER BY crm_checklist.checkID asc";
    $dataBase = new DataBase();
    $result = $dataBase->getResult($sql);
    if ($result["num_rows"] > 0):
        //sacar cabecera en caso tuviese

        foreach ($result["data"] as $y => $data):
            $alternativas = getAlternativas($data);
            if ($data["typeCheck"] == PREGUNTA_FINAL)://OK
                //-- DESGLOSAR 2 EN CATEGORÍA";
                //imprimir el título
                $isleyenda = false;
                if ($alternativas == 0 && $y > 0): //leyendas informativas
                    $isleyenda = true;
                    $align = 'left';
                    $italic = false;
                    $color = $GLOBALS["NEGRO"];

                elseif ($alternativas == 0 && $y == 0): //simular cabeceras de categoría
                    $align = 'left';
                    $italic = true;
                    $color = $GLOBALS["AZUL"];

                else:
                    $questions = array($data["question1"], $data["question2"], $data["question3"], $data["question4"], $data["question5"]);
                    if (in_array(PREGUNTA_CABECERA, $questions))://artificio para pintar las cabeceras
                        //imprimir espacio en blanco porque no tiene cabecera y se tiene categoría previa
                        if ($y == 0):
                        /* $table->addRow();
                          $table->addCell(8000, array('gridSpan' => 6, 'borderLeftSize' => 0,'borderRightSize' => 0,'borderTopSize' => 0, 'borderBottomSize' => 0, 'borderColor'=>$GLOBALS["BLANCO"],'valign'=>'center'))
                          ->addText("S5".SEPARADOR);//qwe "--separador de subpregunta--". */
                        endif;
                        $align = PREGUNTA_ALINEADA; //'right';
                        $italic = false;
                        $color = $GLOBALS["ROJO"];

                    else://otros titulos ya en la respuesta
                        $align = PREGUNTA_ALINEADA; //'right';
                        $italic = false;
                        $color = $GLOBALS["NEGRO"];
                    endif;
                endif;
                $title = clearText($data["title"]);
                $title = ($isleyenda) ? "**" . ucwords(mb_strtolower($title)) : $title; //leyenda

                $cellColSpan = array('gridSpan' => 6 - $alternativas, 'valign' => 'center');

                /* if($isleyenda):
                  $table->addRow();
                  $table->addCell(9000+ANCHO_PUNTUACION, array('gridSpan' => 7, 'borderLeftSize' => 0,'borderRightSize' => 0,'borderTopSize' => 0, 'borderBottomSize' => 0, 'borderColor'=>$GLOBALS["BLANCO"],'valign'=>'center'))->addText(SEPARADOR);
                  endif; */


                $table->addRow();
                $table->addCell(1000 + ((5 - $alternativas) * ANCHO_ALTERNATIVAS), $cellColSpan)
                        ->addText($title, ['name' => 'Arial', 'size' => 8, 'italic' => $italic, 'color' => $color, 'bold' => false, 'valign' => 'center'], ['align' => $align]
                );
                $iscabecera = true;
                if ($alternativas > 0):
                    //tiene alternativas
                    //for($i=1;$i<=$alternativas;$i++):// al menos para el nivel 3 las pregunas van en orden
                    //for($i=$alternativas;$i>0;$i--):
                    $mostrar_no_aplica = false;
                    for ($i = 1; $i <= $alternativas; $i++):
                        $value = "";
                        $color = $GLOBALS["NEGRO"];
                        if ($data["question" . $i] == PREGUNTA_CABECERA):
                            $color = $GLOBALS["ROJO"];
                            if ($data["response" . $i] == ""):
                                $value = $data["text" . $i];
                            else:
                                exit($data["checkID"] . ",inconsistencia en la base de datos, por favor revisar");
                            endif;
                        elseif ($data["question" . $i] == PREGUNTA_CERRADA_SIMPLE || $data["question" . $i] == PREGUNTA_CERRADA_COMPLEJA):
                            $iscabecera = false;
                            if (array_key_exists($data["response" . $i], $GLOBALS["RESPUESTAS"][$data["question" . $i]])):
                                $value = $GLOBALS["RESPUESTAS"][$data["question" . $i]][$data["response" . $i]];
                                if ($data["response" . $i] == 3) {
                                    $mostrar_no_aplica = true;
                                }
                            endif;
                        else:
                            $iscabecera = false;
                            $value = $data["response" . $i];
                        endif;
                        $value = clearText($value);
                        $value = ($value == "") ? REEMPLAZAR_VACIOS : $value;
                        $table->addCell($GLOBALS["WIDTH_FOR_QUESTION"][$data["question" . $i]], array('valign' => 'center'))
                                ->addText($value, ['color' => $color, 'size' => 8], ['align' => 'center']);
                    endfor;

                    if ($data["score"] == PUNTAJE_CALIFICADO && $isleyenda != true && $iscabecera != OCULTAR_SCORE_CABECERAS):
                        //aquí la celda se mueve..
                        $ss = ($mostrar_no_aplica) ? $GLOBALS["RESPUESTAS"][7][3] : getScore($data["final_score"], $data["numScore"]);
                        if ($rowspan):
                            $table->addCell(ANCHO_PUNTUACION, $cellRowContinue);
                        else:
                            $table->addCell(ANCHO_PUNTUACION, $cellRowSpan)
                                    ->addText($ss, ['name' => 'Arial', 'size' => 8, 'color' => $color, 'bold' => true, 'valign' => 'center'], ['align' => 'center']
                            );
                        endif;

                    elseif ($data["score"] == PUNTAJE_ACUMULADO && $isleyenda != true && $iscabecera != OCULTAR_SCORE_CABECERAS):
                        $ss = ($mostrar_no_aplica) ? $GLOBALS["RESPUESTAS"][7][3] : INFORMATIVO;
                        if ($rowspan):
                            $table->addCell(ANCHO_PUNTUACION, $cellRowContinue);
                        else:
                            $table->addCell(ANCHO_PUNTUACION, $cellRowSpan)
                                    ->addText($ss, ['name' => 'Arial', 'size' => 8, 'color' => $color, 'bold' => true, 'valign' => 'center'], ['align' => 'center']
                            );
                        endif;
                    else:
                        if ($rowspan):
                            $table->addCell(ANCHO_PUNTUACION, $cellRowContinue);
                        else:

                            if ($data["score"] == PUNTAJE_ACUMULADO):
                                $ss = ($mostrar_no_aplica) ? $GLOBALS["RESPUESTAS"][7][3] : INFORMATIVO;
                            else:
                                $ss = ($mostrar_no_aplica) ? $GLOBALS["RESPUESTAS"][7][3] : getScore($data["final_score"], $data["numScore"]);
                            endif;

                            $table->addCell(ANCHO_PUNTUACION, $cellRowSpan)
                                    ->addText($ss, ['name' => 'Arial', 'size' => 8, 'color' => $color, 'bold' => true, 'valign' => 'center'], ['align' => 'center']
                            );
                        //completar la celda par el caso en los que hay cabeceras 
                        endif;
                    //no mostrar puntaje del detalle xq en el contenido ya indica que es informativo, caso contrario, pedir a victor que escriba el contenedor
                    endif;

                else:
                    //ya se imprimió la leyenda, ver lineas arriba..
                    //no se imprime score aquí
                    if ($rowspan):
                        $table->addCell(ANCHO_PUNTUACION, $cellRowContinue);
                    else:
                        $table->addCell(ANCHO_PUNTUACION, $cellRowSpan)
                                ->addText("-", ['name' => 'Arial', 'size' => 8, 'color' => $color, 'bold' => false, 'valign' => 'center'], ['align' => 'center']
                        );
                    endif;
                endif;




            else:
                exit("estructura de preguntas no soportadas por el sistema. No se esperaban más preguntas múltiples");
            endif;
        endforeach;
    endif;
}

function getCaffe($homologacionID, $checkID, $somelikeu, &$table, &$section, $rowspan = false, $score_rowspan = 0) {

    $cellRowSpan = array('vMerge' => 'restart', 'valign' => 'center', 'bgColor' => $GLOBALS["BLANCO"]);
    $cellRowContinue = array('vMerge' => 'continue', 'valign' => 'center', 'bgColor' => $GLOBALS["BLANCO"]);
    $cellHCentered = array('alignment' => \PhpOffice\PhpWord\SimpleType\Jc::CENTER);
    $cellVCentered = array('valign' => 'center');

    $somelikeu++;
    $sql = "select 
	crm_check_homo.checkHomoID,
	crm_check_homo.checkID,
	crm_check_homo.homologacionID,
	crm_check_homo.response1,
	crm_check_homo.response2,
	crm_check_homo.response3,
	crm_check_homo.response4,
	crm_check_homo.response5,
	crm_check_homo.score as final_score,
	crm_check_homo.registerDate,
	crm_check_homo.registerUpdate,
	crm_checklist.precheckID,
	crm_checklist.formID,
	crm_checklist.typeCheck,
	crm_checklist.title,
	crm_checklist.question1,
	crm_checklist.question2,
	crm_checklist.question3,
	crm_checklist.question4,
	crm_checklist.question5,
	crm_checklist.text1,
	crm_checklist.text2,
	crm_checklist.text3,
	crm_checklist.text4,
	crm_checklist.text5,
	crm_checklist.score,
	crm_checklist.numScore,
	crm_checklist.information,
	crm_checklist.state
	FROM crm_checklist
	INNER JOIN crm_check_homo ON crm_checklist.checkID = crm_check_homo.checkID
	WHERE 1=1
	AND crm_checklist.precheckID =" . $checkID . " 
	AND crm_check_homo.homologacionID = " . $homologacionID . "
	ORDER BY crm_checklist.checkID asc";
    $dataBase = new DataBase();
    $result = $dataBase->getResult($sql);
    if ($result["num_rows"] > 0):

        foreach ($result["data"] as $y => $data):

            $alternativas = getAlternativas($data);
            $gmh = 0;
            $title = clearText($data["title"]);
            //$title = ($title=="")?"olvidaron poner un títutlo aquí, las alternativas son ".$alternativas:$title;
            //Gean: Victor, debe quitar de la base de datos aquellas preguntas 21 que no tiene titutlo de pregunta/alternativa
            $isleyenda = false;
            //typeCheck, toma valores entre 1 y 0
            if ($data["typeCheck"] == PREGUNTA_FINAL): // OK

                if ($alternativas == 0):
                    if ($y > 0): //leyendas informativas
                        $isleyenda = true;
                        $align = 'left';
                        $italic = false;
                        $color = $GLOBALS["NEGRO"];

                    elseif ($y == 0): //simular cabeceras de categoría
                        $align = 'left';
                        $italic = true;
                        $color = $GLOBALS["AZUL"]; //AZUL_MARINO
                    else:

                    endif;

                else:
                    $questions = array($data["question1"], $data["question2"], $data["question3"], $data["question4"], $data["question5"]);
                    if (in_array(PREGUNTA_CABECERA, $questions)):
                        //imprimir espacio en blanco porque no tiene cabecera y se tiene categoría previa
                        if ($y == 0):
                        /* $table->addRow();
                          $table->addCell(9000+ANCHO_PUNTUACION, array('gridSpan' => 7, 'borderLeftSize' => 0,'borderRightSize' => 0,'borderTopSize' => 0, 'borderBottomSize' => 0, 'borderColor'=>$GLOBALS["BLANCO"],'valign'=>'center'))
                          ->addText(SEPARADOR);//qwe "--separador de subpregunta--". */
                        endif;

                        $align = PREGUNTA_ALINEADA; //'right';
                        $italic = false;
                        if ($data["response1"] == "" && $data["response2"] == "" && $data["response3"] == "" && $data["response4"] == "" && $data["response5"] == ""):
                            $color = $GLOBALS["ROJO"];
                        else:
                            $color = $GLOBALS["NEGRO"];
                        endif;

                    else://otros titulos ya en la respuesta
                        $align = PREGUNTA_ALINEADA; //'right';
                        $italic = false;
                        $color = $GLOBALS["NEGRO"];
                    endif;
                endif;


                $title = ($isleyenda) ? "**" . ucwords(mb_strtolower($title)) : $title; //leyenda
                $cellColSpan = array('gridSpan' => 6 - $alternativas, 'valign' => 'center');

                $iscabecera = true;
                if ($alternativas > 0):
                    if ($isleyenda):
                    /* $table->addRow();
                      $table->addCell(9000+ANCHO_PUNTUACION, array('gridSpan' => 7, 'borderLeftSize' => 0,'borderRightSize' => 0,'borderTopSize' => 0, 'borderBottomSize' => 0, 'borderColor'=>$GLOBALS["BLANCO"],'valign'=>'center'))->addText(SEPARADOR); */
                    endif;
                    //tiene alternativas
                    $width = 1000 + ((5 - $alternativas) * ANCHO_ALTERNATIVAS);
                    $table->addRow();
                    $table->addCell($width, $cellColSpan)
                            ->addText($title, ['name' => 'Arial', 'size' => 8, 'italic' => $italic, 'color' => $color, 'bold' => false, 'valign' => 'center'], ['align' => $align]
                    );

                    $mostrar_no_aplica = false;
                    for ($i = 1; $i <= $alternativas; $i++):
                        $value = "";
                        $color = $GLOBALS["NEGRO"];
                        if ($data["question" . $i] == PREGUNTA_CABECERA):
                            if ($data["response1"] == "" && $data["response2"] == "" && $data["response3"] == "" && $data["response4"] == "" && $data["response5"] == ""):
                                $color = $GLOBALS["ROJO"];
                                $value = $data["text" . $i];
                            else:
                                $color = $GLOBALS["NEGRO"];
                                $value = $data["text" . $i];
                            endif;
                        elseif ($data["question" . $i] == PREGUNTA_CERRADA_SIMPLE || $data["question" . $i] == PREGUNTA_CERRADA_COMPLEJA):
                            $iscabecera = false;
                            $color = $GLOBALS["NEGRO"];
                            if (array_key_exists($data["response" . $i], $GLOBALS["RESPUESTAS"][$data["question" . $i]])):
                                $value = $GLOBALS["RESPUESTAS"][$data["question" . $i]][$data["response" . $i]];
                                if ($data["response" . $i] == 3) {
                                    $mostrar_no_aplica = true;
                                }
                            endif;
                        else:
                            $iscabecera = false;
                            $color = $GLOBALS["NEGRO"];
                            $value = $data["response" . $i];

                        endif;
                        $value = clearText($value);
                        $value = ($value == "") ? REEMPLAZAR_VACIOS : $value;
                        $width = $GLOBALS["WIDTH_FOR_QUESTION"][$data["question" . $i]];
                        $table->addCell($width, array('valign' => 'center'))
                                ->addText($value, ['size' => 8, 'color' => $color], ['align' => 'center']);
                    endfor;
                else:
                    if ($isleyenda):
                    /* $table->addRow();
                      $table->addCell(9000+ANCHO_PUNTUACION, array('gridSpan' => 7, 'borderLeftSize' => 0,'borderRightSize' => 0,'borderTopSize' => 0, 'borderBottomSize' => 0, 'borderColor'=>$GLOBALS["BLANCO"],'valign'=>'center'))->addText(SEPARADOR); */
                    endif;
                    //se mantiene $iscabecera = true para evitar que tenga el score en el titulo qye ocupa toda la fila;
                    $table->addRow();
                    $table->addCell(1000 + (5 * ANCHO_ALTERNATIVAS), $cellColSpan)
                            ->addText($title, ['name' => 'Arial', 'size' => 8, 'italic' => $italic, 'color' => $color, 'bold' => false, 'valign' => 'center'], ['align' => $align]
                    );
                //no tiene alternativas, eso quiere decir que pintó el título con un colspan de 6 (ver líneas arriba).
                endif;

                //calificaciones:

                if ($data["score"] == PUNTAJE_CALIFICADO && $isleyenda != true && $iscabecera != OCULTAR_SCORE_CABECERAS):
                    //aquí la celda se mueve..
                    if ($rowspan):
                        $table->addCell(ANCHO_PUNTUACION, $cellRowContinue);
                    /* ->addText($score_rowspan,['name' => 'Arial', 'size' => 8, 'bold' => true],
                      ['align'=>'center']); */

                    else:
                        $ss = ($mostrar_no_aplica) ? $GLOBALS["RESPUESTAS"][7][3] : getScore($data["final_score"], $data["numScore"]);
                        $table->addCell(ANCHO_PUNTUACION, $cellRowSpan)
                                ->addText($ss, ['name' => 'Arial', 'size' => 8, 'color' => $color, 'bold' => true, 'valign' => 'center'], ['align' => 'center']
                        );
                    endif;
                elseif ($data["score"] == PUNTAJE_ACUMULADO && $isleyenda != true && $iscabecera != OCULTAR_SCORE_CABECERAS):
                    if ($rowspan):
                        $table->addCell(ANCHO_PUNTUACION, $cellRowContinue);
                    /* ->addText($score_rowspan,['name' => 'Arial', 'size' => 8, 'bold' => true],
                      ['align'=>'center']); */
                    else:
                        $ss = ($mostrar_no_aplica) ? $GLOBALS["RESPUESTAS"][7][3] : INFORMATIVO;
                        $table->addCell(ANCHO_PUNTUACION, $cellRowSpan)
                                ->addText($ss, ['name' => 'Arial', 'size' => 8, 'color' => $color, 'bold' => true, 'valign' => 'center'], ['align' => 'center']
                        );
                    endif;
                else:
                    if ($rowspan):
                        $table->addCell(ANCHO_PUNTUACION, $cellRowContinue);
                    /* ->addText($score_rowspan,['name' => 'Arial', 'size' => 8, 'bold' => true],
                      ['align'=>'center']); */
                    else:
                        $table->addCell(ANCHO_PUNTUACION, $cellRowSpan)->addText(SEPARADOR);
                    endif;
                //no mostrar puntaje del detalle xq en el contenido ya indica que es informativo, caso contrario, pedir a victor que escriba el contenedor
                endif;
            elseif ($data["typeCheck"] == PREGUNTA_MULTIPLE)://OK

                $somelikeu = 2;
                if ($alternativas == 0):
                    if ($y == 0) {
                        $isleyenda = true;
                    }
                    //son cabeceras
                    /* $table->addRow();
                      $table->addCell(9000+ANCHO_PUNTUACION, array('gridSpan' => 7, 'borderLeftSize' => 0,'borderRightSize' => 0,'borderTopSize' => 0, 'borderBottomSize' => 0, 'borderColor'=>$GLOBALS["BLANCO"],'valign'=>'center'))
                      ->addText("soyseparador".SEPARADOR); */

                    $table->addRow();
                    $table->addCell(1000 + (5 * ANCHO_ALTERNATIVAS), array('gridSpan' => 6, 'valign' => 'center'))
                            ->addText($title, array('name' => 'Arial', 'size' => 8, 'italic' => true, 'color' => $GLOBALS["AZUL"], 'bold' => true), ['align' => 'left']
                    );

                    if ($data["score"] == PUNTAJE_CALIFICADO && $isleyenda != true):
                        //difernt de title debido a que aquí se estan dando contenedores dentro de contenedores, por tanto no tendría sentido poner porcentaje cuando no hay una descripciíon del titulo
                        $score_rowspan = getScore($data["final_score"], $data["numScore"]);
                        if ($rowspan):
                            $table->addCell(ANCHO_PUNTUACION, $cellRowContinue);
                        else:
                            $table->addCell(ANCHO_PUNTUACION, $cellRowSpan)
                                    ->addText($score_rowspan, ['size' => 8, 'color' => $GLOBALS["NEGRO"], 'bold' => true], ['align' => 'center']);
                        endif;
                        $rowspan_next = true;


                    elseif ($data["score"] == PUNTAJE_ACUMULADO && $isleyenda != true):

                        if ($rowspan):
                            $table->addCell(ANCHO_PUNTUACION, $cellRowContinue);
                        else:
                            $table->addCell(ANCHO_PUNTUACION, $cellRowSpan)
                                    ->addText($score_rowspan, ['size' => 8, 'color' => $GLOBALS["NEGRO"], 'bold' => true], ['align' => 'center']);
                        endif;
                        $rowspan_next = false;
                        $score_rowspan = INFORMATIVO;

                    else://no debemos poner nada debido a que son subpreguntas de subpreguntas

                        $score_rowspan = ($data["score"] == PUNTAJE_ACUMULADO) ? INFORMATIVO : getScore($data["final_score"], $data["numScore"]);

                        if ($rowspan):
                            $score_rowspan = getScore($data["final_score"], $data["numScore"]);
                            $table->addCell(ANCHO_PUNTUACION, $cellRowContinue);
                        else:
                            $score_rowspan = getScore($data["final_score"], $data["numScore"]);
                            $table->addCell(ANCHO_PUNTUACION, $cellRowSpan)
                                    ->addText($score_rowspan, ['size' => 8, 'color' => $GLOBALS["NEGRO"], 'bold' => true], ['align' => 'center']);
                        endif;
                        $rowspan_next = ($data["score"] == PUNTAJE_ACUMULADO) ? false : true;

                    endif;
                else:

                    //exit("no esperaba que tuvieses más preguntas");
                    //imprimir las preguntas
                    if ($isleyenda):
                    // $table->addRow();
                    // $table->addCell(9000+ANCHO_PUNTUACION, array('gridSpan' => 7, 'borderLeftSize' => 0,'borderRightSize' => 0,'borderTopSize' => 0, 'borderBottomSize' => 0, 'borderColor'=>$GLOBALS["BLANCO"],'valign'=>'center'))->addText(SEPARADOR);
                    endif;
                    //tiene alternativas
                    $width = 1000 + ((5 - $alternativas) * ANCHO_ALTERNATIVAS);
                    $table->addRow();
                    $table->addCell($width, $cellColSpan)
                            ->addText($title, ['name' => 'Arial', 'size' => 8, 'italic' => $italic, 'color' => $color, 'bold' => false, 'valign' => 'center'], ['align' => $align]
                    );

                    $mostrar_no_aplica = false;
                    for ($i = 1; $i <= $alternativas; $i++):
                        $value = "";
                        $color = $GLOBALS["NEGRO"];
                        if ($data["question" . $i] == PREGUNTA_CABECERA):
                            if ($data["response1"] == "" && $data["response2"] == "" && $data["response3"] == "" && $data["response4"] == "" && $data["response5"] == ""):
                                $color = $GLOBALS["ROJO"];
                                $value = $data["text" . $i];
                            else:
                                $color = $GLOBALS["NEGRO"];
                                $value = $data["text" . $i];
                            endif;
                        elseif ($data["question" . $i] == PREGUNTA_CERRADA_SIMPLE || $data["question" . $i] == PREGUNTA_CERRADA_COMPLEJA):
                            $iscabecera = false;
                            $color = $GLOBALS["NEGRO"];
                            if (array_key_exists($data["response" . $i], $GLOBALS["RESPUESTAS"][$data["question" . $i]])):
                                $value = $GLOBALS["RESPUESTAS"][$data["question" . $i]][$data["response" . $i]];
                                if ($data["response" . $i] == 3) {
                                    $mostrar_no_aplica = true;
                                }
                            endif;
                        else:
                            $iscabecera = false;
                            $color = $GLOBALS["NEGRO"];
                            $value = $data["response" . $i];

                        endif;
                        $value = clearText($value);
                        $value = ($value == "") ? REEMPLAZAR_VACIOS : $value;
                        $table->addCell($GLOBALS["WIDTH_FOR_QUESTION"][$data["question" . $i]], array('valign' => 'center'))
                                ->addText($value, ['size' => 8, 'color' => $color], ['align' => 'center']);
                    endfor;

                    if ($data["score"] == PUNTAJE_CALIFICADO):
                        //difernt de title debido a que aquí se estan dando contenedores dentro de contenedores, por tanto no tendría sentido poner porcentaje cuando no hay una descripciíon del titulo
                        $score_rowspan = getScore($data["final_score"], $data["numScore"]);
                        if ($rowspan):
                            $table->addCell(ANCHO_PUNTUACION, $cellRowContinue);
                        else:
                            $table->addCell(ANCHO_PUNTUACION, $cellRowSpan)
                                    ->addText($score_rowspan, ['size' => 8, 'color' => $GLOBALS["NEGRO"], 'bold' => true], ['align' => 'center']);
                        endif;
                        $rowspan_next = true;


                    else://PUNTAJE_ACUMULADO

                        if ($rowspan):
                            $table->addCell(ANCHO_PUNTUACION, $cellRowContinue);
                        else:
                            $table->addCell(ANCHO_PUNTUACION, $cellRowSpan)
                                    ->addText($score_rowspan, ['size' => 8, 'color' => $GLOBALS["NEGRO"], 'bold' => true], ['align' => 'center']);
                        endif;
                        $rowspan_next = false;
                        $score_rowspan = INFORMATIVO;
                    endif;
                endif;
                getHamburger($homologacionID, $data["checkID"], $somelikeu, $table, $section, $rowspan_next, $score_rowspan);

            else:
                exit("Advertencia, la columna typeCheck con valor " . $data["typeCheck"] . " no existe el valor en la base de datos, mayor detalle aquí");
            endif;

        endforeach;
    endif;
}

function getAlternativas($data) {
    if (($data["question5"] + 0) > 0) {
        $alternativas = 5;
    } elseif (($data["question4"] + 0) > 0) {
        $alternativas = 4;
    } elseif (($data["question3"] + 0) > 0) {
        $alternativas = 3;
    } elseif (($data["question2"] + 0) > 0) {
        $alternativas = 2;
    } elseif (($data["question1"] + 0) > 0) {
        $alternativas = 1;
    } else {
        $alternativas = 0;
    }
    return $alternativas;
}


/*function getNap($checkID = 0, $precheckID = 0, $whereIn) {
    $sql = "SELECT precheckID,checkID
	FROM crm_checklist 
	WHERE state=1 
	AND precheckID IN (" . $checkID . ") 
	ORDER BY checkID ASC";
    $database = new DataBase();
    $result = $database->getResult($sql);
    if ($result["num_rows"] > 0):
        $checkID = "";
        $precheckID = "";
        foreach ($result["data"] as $data):
            $checkID.= "," . $data["checkID"];
            $precheckID.= "," . $data["precheckID"];
        endforeach;
        $checkID = substr($checkID, 1);
        $precheckID = substr($precheckID, 1);
        array_push($whereIn, ["checkID" => $checkID, "precheckID" => $precheckID]);
        $whereIn = getNap($checkID, $precheckID, $whereIn);
    endif;
    return $whereIn;
}*/