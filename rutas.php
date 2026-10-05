<?php

require 'vendor/autoload.php';

use PhpOffice\PhpSpreadsheet\IOFactory;
use PhpOffice\PhpSpreadsheet\Spreadsheet;
use PhpOffice\PhpSpreadsheet\Writer\Xlsx;
use PhpOffice\PhpSpreadsheet\Style\Fill;
use PhpOffice\PhpSpreadsheet\Style\Alignment;

/*
==========================================================
        SISTEMA DE RUTAS - BACKTRACKING + GOOGLE MAPS
==========================================================
*/

// ========================================================
// 1. CONFIGURACIÓN
// ========================================================

$GOOGLE_API_KEY = "AIzaSyBJNrjA3hDdMvPZvrdSzppQda6Z30nnrok";
$horaInicioJornada = 8;
$K = 4;


// ========================================================
// 2. VARIABLES
// ========================================================

$clientes = [];
$distancias = [];
$tiemposViaje = [];
$rutasValidas = [];
$error = "";
$archivoExcelGenerado = null;


// ========================================================
// 3. FUNCIONES
// ========================================================

function formatoHora($hora)
{
    $horas = floor($hora);
    $minutos = round(($hora - $horas) * 60);

    if ($minutos >= 60) {
        $horas++;
        $minutos -= 60;
    }

    return sprintf("%02d:%02d", $horas, $minutos);
}

function duracionGoogleAHoras($duracion)
{
    if (empty($duracion)) {
        return 0;
    }

    $segundos = floatval(str_replace("s", "", $duracion));
    return $segundos / 3600;
}

function generarEnlaceGoogleMaps($ruta, $clientes)
{
    $origen = $clientes[0]["direccion"];
    $destino = $clientes[0]["direccion"];
    $waypoints = [];

    foreach ($ruta as $cliente) {
        $waypoints[] = $clientes[$cliente]["direccion"];
    }

    $url = "https://www.google.com/maps/dir/?api=1";
    $url .= "&origin=" . urlencode($origen);
    $url .= "&destination=" . urlencode($destino);

    if (count($waypoints) > 0) {
        $url .= "&waypoints=" . urlencode(implode("|", $waypoints));
    }

    $url .= "&travelmode=driving";

    return $url;
}


// ========================================================
// 4. CREAR Y GUARDAR EXCEL EN SERVIDOR TEMPORALMENTE
// ========================================================

function generarArchivoExcel($mejoresRutas, $clientes)
{
    $spreadsheet = new Spreadsheet();
    $spreadsheet->removeSheetByIndex(0);

    $secciones = [
        "rapida" => ["titulo" => "Ruta Mas Rapida", "color" => "FF9800"],
        "efectiva" => ["titulo" => "Ruta Mas Efectiva", "color" => "2196F3"],
        "viable" => ["titulo" => "Ruta Mas Viable", "color" => "4CAF50"]
    ];

    foreach ($secciones as $clave => $meta) {
        if (!isset($mejoresRutas[$clave])) continue;

        $item = $mejoresRutas[$clave];
        $ruta = $item["ruta"];
        $res = $item["resultado"];
        $link = generarEnlaceGoogleMaps($ruta, $clientes);

        $hoja = $spreadsheet->createSheet();
        $hoja->setTitle($meta["titulo"]);

        // Título
        $hoja->setCellValue('A1', strtoupper($meta["titulo"]));
        $hoja->mergeCells('A1:D1');
        $hoja->getStyle('A1')->getFont()->setBold(true)->setSize(14)->getColor()->setARGB('FFFFFF');
        $hoja->getStyle('A1')->getFill()->setFillType(Fill::FILL_SOLID)->getStartColor()->setARGB($meta["color"]);
        $hoja->getStyle('A1')->getAlignment()->setHorizontal(Alignment::HORIZONTAL_CENTER);

        // Métricas
        $hoja->setCellValue('A3', 'Distancia Total:');
        $hoja->setCellValue('B3', number_format($res["distancia"], 2) . ' km');
        
        $hoja->setCellValue('A4', 'Fin de Jornada:');
        $hoja->setCellValue('B4', formatoHora($res["tiempoTotal"]));
        
        $hoja->setCellValue('A5', 'Tiempo de Espera:');
        $hoja->setCellValue('B5', formatoHora($res["tiempoEspera"]));

        $hoja->setCellValue('A6', 'Enlace Google Maps:');
        $hoja->setCellValue('B6', $link);
        $hoja->getCell('B6')->getHyperlink()->setUrl($link);

        $hoja->getStyle('A3:A6')->getFont()->setBold(true);

        // Tabla
        $hoja->setCellValue('A8', 'Orden');
        $hoja->setCellValue('B8', 'Cliente / Ubicación');
        $hoja->setCellValue('C8', 'Hora Llegada');
        $hoja->setCellValue('D8', 'Fin Servicio');

        $hoja->getStyle('A8:D8')->getFont()->setBold(true);
        $hoja->getStyle('A8:D8')->getFill()->setFillType(Fill::FILL_SOLID)->getStartColor()->setARGB('EEEEEE');

        // Filas
        $filaIndex = 9;
        $hoja->setCellValue('A' . $filaIndex, 1);
        $hoja->setCellValue('B' . $filaIndex, $clientes[0]["nombre"] . " (Bodega)");
        $hoja->setCellValue('C' . $filaIndex, formatoHora(8));
        $hoja->setCellValue('D' . $filaIndex, formatoHora(8));
        $filaIndex++;

        $orden = 2;
        foreach ($res["detalle"] as $d) {
            $hoja->setCellValue('A' . $filaIndex, $orden);
            $hoja->setCellValue('B' . $filaIndex, $d["nombre"]);
            $hoja->setCellValue('C' . $filaIndex, formatoHora($d["llegada"]));
            $hoja->setCellValue('D' . $filaIndex, formatoHora($d["finServicio"]));
            $filaIndex++;
            $orden++;
        }

        $hoja->setCellValue('A' . $filaIndex, $orden);
        $hoja->setCellValue('B' . $filaIndex, $clientes[0]["nombre"] . " (Retorno Bodega)");
        $hoja->setCellValue('C' . $filaIndex, formatoHora($res["tiempoTotal"]));
        $hoja->setCellValue('D' . $filaIndex, formatoHora($res["tiempoTotal"]));

        foreach (range('A', 'D') as $col) {
            $hoja->getColumnDimension($col)->setAutoSize(true);
        }
    }

    $nombreArchivo = 'Reporte_Rutas_ZonaLogistik_' . time() . '.xlsx';
    $rutaGuardado = __DIR__ . '/' . $nombreArchivo;

    $writer = new Xlsx($spreadsheet);
    $writer->save($rutaGuardado);

    return $nombreArchivo;
}


// ========================================================
// 5. ATENDER DESCARGA DE ARCHIVO TEMPORAL
// ========================================================

if (isset($_GET['descargar'])) {
    $archivo = basename($_GET['descargar']);
    $rutaArchivo = __DIR__ . '/' . $archivo;

    if (file_exists($rutaArchivo)) {
        header('Content-Type: application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
        header('Content-Disposition: attachment; filename="' . $archivo . '"');
        header('Content-Length: ' . filesize($rutaArchivo));
        readfile($rutaArchivo);
        @unlink($rutaArchivo); // Elimina el archivo después de descargarlo
        exit;
    }
}


// ========================================================
// 6. LEER EXCEL ENTRADA
// ========================================================

function leerExcel($archivo)
{
    $clientes = [];
    $spreadsheet = IOFactory::load($archivo);
    $hoja = $spreadsheet->getActiveSheet();
    $filas = $hoja->toArray(null, true, true, true);

    foreach ($filas as $numero => $fila) {
        if ($numero == 1) continue;

        if (empty($fila["A"]) && $fila["A"] !== "0" && $fila["A"] !== 0) {
            continue;
        }

        $id = intval($fila["A"]);
        $cliente = trim($fila["B"]);
        $direccion = trim($fila["C"]);
        $horaInicio = floatval($fila["D"]);
        $horaFin = floatval($fila["E"]);
        $tiempoServicio = floatval($fila["F"]);

        $clientes[$id] = [
            "id" => $id,
            "nombre" => $cliente,
            "direccion" => $direccion,
            "horaInicio" => $horaInicio,
            "horaFin" => $horaFin,
            "tiempoServicio" => $tiempoServicio
        ];
    }

    ksort($clientes);
    return array_values($clientes);
}


// ========================================================
// 7. GOOGLE ROUTES API
// ========================================================

function obtenerMatrizGoogle($clientes, $apiKey)
{
    $cantidad = count($clientes);

    $distancias = array_fill(0, $cantidad, array_fill(0, $cantidad, 0));
    $tiempos = array_fill(0, $cantidad, array_fill(0, $cantidad, 0));

    $origins = [];
    $destinations = [];

    foreach ($clientes as $cliente) {
        $origins[] = ["waypoint" => ["address" => $cliente["direccion"]]];
        $destinations[] = ["waypoint" => ["address" => $cliente["direccion"]]];
    }

    $data = [
        "origins" => $origins,
        "destinations" => $destinations,
        "travelMode" => "DRIVE",
        "routingPreference" => "TRAFFIC_AWARE"
    ];

    $json = json_encode($data);
    $url = "https://routes.googleapis.com/distanceMatrix/v2:computeRouteMatrix";

    $headers = [
        "Content-Type: application/json",
        "X-Goog-Api-Key: " . $apiKey,
        "X-Goog-FieldMask: originIndex,destinationIndex,status,condition,distanceMeters,duration"
    ];

    $ch = curl_init($url);
    curl_setopt($ch, CURLOPT_POST, true);
    curl_setopt($ch, CURLOPT_POSTFIELDS, $json);
    curl_setopt($ch, CURLOPT_HTTPHEADER, $headers);
    curl_setopt($ch, CURLOPT_RETURNTRANSFER, true);

    $respuesta = curl_exec($ch);
    $httpCode = curl_getinfo($ch, CURLINFO_HTTP_CODE);
    $curlError = curl_error($ch);
    curl_close($ch);

    if ($curlError) {
        throw new Exception("Error CURL: " . $curlError);
    }

    if ($httpCode != 200) {
        throw new Exception("Google Routes API respondió HTTP " . $httpCode . ": " . $respuesta);
    }

    $resultado = json_decode($respuesta, true);

    if (!is_array($resultado)) {
        throw new Exception("Respuesta inválida de Google.");
    }

    foreach ($resultado as $elemento) {
        if (!isset($elemento["originIndex"], $elemento["destinationIndex"])) continue;

        $origen = intval($elemento["originIndex"]);
        $destino = intval($elemento["destinationIndex"]);

        if ($origen === $destino) {
            $distancias[$origen][$destino] = 0;
            $tiempos[$origen][$destino] = 0;
            continue;
        }

        if (isset($elemento["condition"]) && $elemento["condition"] != "ROUTE_EXISTS") {
            $distancias[$origen][$destino] = INF;
            $tiempos[$origen][$destino] = INF;
            continue;
        }

        $metros = $elemento["distanceMeters"] ?? 0;
        $duracion = $elemento["duration"] ?? "0s";

        $distancias[$origen][$destino] = $metros / 1000;
        $tiempos[$origen][$destino] = duracionGoogleAHoras($duracion);
    }

    return [
        "distancias" => $distancias,
        "tiempos" => $tiempos
    ];
}


// ========================================================
// 8. EVALUAR RUTA
// ========================================================

function evaluarRuta($ruta, $distancias, $tiemposViaje, $clientes, $horaInicioJornada)
{
    $tiempoActual = $horaInicioJornada;
    $distanciaTotal = 0;
    $tiempoServicioTotal = 0;
    $tiempoEsperaTotal = 0;
    $puntoActual = 0;
    $detalle = [];

    foreach ($ruta as $cliente) {
        $distancia = $distancias[$puntoActual][$cliente];
        if ($distancia === INF) return ["valida" => false];
        $distanciaTotal += $distancia;

        $tiempoViaje = $tiemposViaje[$puntoActual][$cliente];
        if ($tiempoViaje === INF) return ["valida" => false];
        $tiempoActual += $tiempoViaje;

        if ($tiempoActual < $clientes[$cliente]["horaInicio"]) {
            $espera = $clientes[$cliente]["horaInicio"] - $tiempoActual;
            $tiempoEsperaTotal += $espera;
            $tiempoActual = $clientes[$cliente]["horaInicio"];
        }

        if ($tiempoActual > $clientes[$cliente]["horaFin"]) {
            return ["valida" => false];
        }

        $horaLlegada = $tiempoActual;
        $tiempoActual += $clientes[$cliente]["tiempoServicio"];
        $tiempoServicioTotal += $clientes[$cliente]["tiempoServicio"];

        $detalle[] = [
            "cliente" => $cliente,
            "nombre" => $clientes[$cliente]["nombre"],
            "llegada" => $horaLlegada,
            "finServicio" => $tiempoActual
        ];

        $puntoActual = $cliente;
    }

    $distanciaRegreso = $distancias[$puntoActual][0];
    $tiempoRegreso = $tiemposViaje[$puntoActual][0];

    if ($distanciaRegreso === INF || $tiempoRegreso === INF) {
        return ["valida" => false];
    }

    $distanciaTotal += $distanciaRegreso;
    $tiempoActual += $tiempoRegreso;

    return [
        "valida" => true,
        "distancia" => $distanciaTotal,
        "tiempoTotal" => $tiempoActual,
        "tiempoServicio" => $tiempoServicioTotal,
        "tiempoEspera" => $tiempoEsperaTotal,
        "detalle" => $detalle
    ];
}


// ========================================================
// 9. BACKTRACKING
// ========================================================

function backtracking($ruta, $visitados, $N, $K, $distancias, $tiemposViaje, $clientes, $horaInicioJornada, &$rutasValidas)
{
    if (count($ruta) == $N) {
        $resultado = evaluarRuta($ruta, $distancias, $tiemposViaje, $clientes, $horaInicioJornada);

        if ($resultado["valida"]) {
            $rutasValidas[] = [
                "ruta" => $ruta,
                "resultado" => $resultado
            ];
        }
        return;
    }

    if (count($ruta) >= $K) return;

    for ($cliente = 1; $cliente <= $N; $cliente++) {
        if ($visitados[$cliente]) continue;

        $ruta[] = $cliente;
        $visitados[$cliente] = true;

        backtracking($ruta, $visitados, $N, $K, $distancias, $tiemposViaje, $clientes, $horaInicioJornada, $rutasValidas);

        array_pop($ruta);
        $visitados[$cliente] = false;
    }
}


// ========================================================
// 10. OBTENER LAS 3 MEJORES RUTAS
// ========================================================

function obtenerTresMejoresRutas($rutasValidas)
{
    if (empty($rutasValidas)) return [];

    $masRapida = $rutasValidas[0];
    foreach ($rutasValidas as $r) {
        if ($r["resultado"]["tiempoTotal"] < $masRapida["resultado"]["tiempoTotal"]) {
            $masRapida = $r;
        }
    }

    $masEfectiva = $rutasValidas[0];
    foreach ($rutasValidas as $r) {
        if ($r["resultado"]["distancia"] < $masEfectiva["resultado"]["distancia"]) {
            $masEfectiva = $r;
        }
    }

    $masViable = $rutasValidas[0];
    foreach ($rutasValidas as $r) {
        if ($r["resultado"]["tiempoEspera"] < $masViable["resultado"]["tiempoEspera"]) {
            $masViable = $r;
        }
    }

    return [
        "rapida" => $masRapida,
        "efectiva" => $masEfectiva,
        "viable" => $masViable
    ];
}


// ========================================================
// 11. PROCESAR SOLICITUD
// ========================================================

if ($_SERVER["REQUEST_METHOD"] === "POST") {
    try {
        if (!isset($_FILES["excel"]) || $_FILES["excel"]["error"] !== UPLOAD_ERR_OK) {
            throw new Exception("Error al cargar el archivo Excel.");
        }

        $archivo = $_FILES["excel"]["tmp_name"];
        $clientes = leerExcel($archivo);

        if (count($clientes) < 2) {
            throw new Exception("El Excel debe tener al menos una bodega y un cliente.");
        }

        $N = count($clientes) - 1;
        if ($K > $N) $K = $N;

        $matrizGoogle = obtenerMatrizGoogle($clientes, $GOOGLE_API_KEY);
        $distancias = $matrizGoogle["distancias"];
        $tiemposViaje = $matrizGoogle["tiempos"];

        $visitados = array_fill(0, $N + 1, false);
        $visitados[0] = true;

        backtracking([], $visitados, $N, $K, $distancias, $tiemposViaje, $clientes, $horaInicioJornada, $rutasValidas);

        $mejoresRutas = obtenerTresMejoresRutas($rutasValidas);

        if (!empty($mejoresRutas)) {
            $archivoExcelGenerado = generarArchivoExcel($mejoresRutas, $clientes);
        }

    } catch (Exception $e) {
        $error = $e->getMessage();
    }
}

?>
<!DOCTYPE html>
<html lang="es">
<head>
    <meta charset="UTF-8">
    <title>Optimización de Rutas ZONA LOGISTIK</title>
    <style>
        body { font-family: Arial, sans-serif; background: #f4f6f8; margin: 40px; }
        .contenedor { max-width: 1000px; margin: auto; background: white; padding: 30px; border-radius: 10px; box-shadow: 0 2px 10px rgba(0,0,0,0.1); }
        h1, h2, h3 { color: #333; }
        .error { background: #ffd6d6; padding: 15px; color: #900; margin-top: 20px; border-radius: 5px; }
        
        .tarjeta-ruta { background: #ffffff; border-radius: 8px; padding: 20px; margin-bottom: 25px; border-left: 6px solid #ccc; box-shadow: 0 2px 8px rgba(0,0,0,0.05); }
        .rapida { border-left-color: #ff9800; background: #fffdf9; }
        .efectiva { border-left-color: #2196f3; background: #f8fcff; }
        .viable { border-left-color: #4caf50; background: #f9fff9; }

        .etiqueta { display: inline-block; padding: 4px 12px; border-radius: 20px; color: white; font-weight: bold; font-size: 14px; margin-bottom: 10px; }
        .bg-rapida { background: #ff9800; }
        .bg-efectiva { background: #2196f3; }
        .bg-viable { background: #4caf50; }

        table { width: 100%; border-collapse: collapse; margin-top: 15px; background: white; }
        th, td { border: 1px solid #ccc; padding: 8px; text-align: left; }
        th { background: #eee; }

        .boton { display: inline-block; padding: 10px 15px; background: #1a73e8; color: white; text-decoration: none; border-radius: 5px; font-weight: bold; margin-top: 15px; }
        .boton:hover { background: #1557b0; }
        .boton-excel { background: #1e7e34; }
        .boton-excel:hover { background: #145a23; }
        button[type="submit"] { padding: 10px 20px; background: #28a745; color: white; border: none; border-radius: 5px; cursor: pointer; font-size: 16px; font-weight: bold; }
        button[type="submit"]:hover { background: #218838; }
    </style>
</head>
<body>

<div class="contenedor">

    <h1>🚚 Optimización de rutas <b><span style="color: red;">ZONA</span>LOGISTIK</b></h1>

    <form method="POST" enctype="multipart/form-data">
        <p>Selecciona el archivo Excel con tus direcciones:</p>
        <input type="file" name="excel" accept=".xlsx,.xls" required>
        <br><br>
        <button type="submit">Calcular Rutas y Generar Reporte</button>
    </form>

    <?php if ($error): ?>
        <div class="error">
            <strong>Error:</strong> <?= htmlspecialchars($error) ?>
        </div>
    <?php endif; ?>

    <?php if (!empty($mejoresRutas)): ?>

        <div style="display: flex; justify-content: space-between; align-items: center; margin-top: 35px;">
            <h2>⭐ Las 3 Mejores Rutas</h2>
            
            <?php if ($archivoExcelGenerado): ?>
                <a href="?descargar=<?= urlencode($archivoExcelGenerado) ?>" class="boton boton-excel">
                    📊 Volver a Descargar Excel
                </a>
            <?php endif; ?>
        </div>

        <?php
        $secciones = [
            "rapida" => ["titulo" => "⚡ RUTA MÁS RÁPIDA", "sub" => "Termina la jornada lo antes posible.", "clase" => "rapida", "bg" => "bg-rapida"],
            "efectiva" => ["titulo" => "📏 RUTA MÁS EFECTIVA", "sub" => "Recorre la menor cantidad de kilómetros.", "clase" => "efectiva", "bg" => "bg-efectiva"],
            "viable" => ["titulo" => "⏱ RUTA MÁS VIABLE", "sub" => "Minimiza los tiempos muertos de espera.", "clase" => "viable", "bg" => "bg-viable"]
        ];

        foreach ($secciones as $clave => $meta):
            $item = $mejoresRutas[$clave];
            $ruta = $item["ruta"];
            $res = $item["resultado"];
            $link = generarEnlaceGoogleMaps($ruta, $clientes);
        ?>

            <div class="tarjeta-ruta <?= $meta["clase"] ?>">
                <span class="etiqueta <?= $meta["bg"] ?>"><?= $meta["titulo"] ?></span>
                <p style="color: #666; font-size: 13px; margin-top: 0;"><?= $meta["sub"] ?></p>

                <p><strong>Recorrido:</strong> Bodega 
                    <?php foreach ($ruta as $c): ?>
                        ➔ <?= htmlspecialchars($clientes[$c]["nombre"]) ?>
                    <?php endforeach; ?>
                    ➔ Bodega
                </p>

                <p>
                    <strong>Distancia:</strong> <?= number_format($res["distancia"], 2) ?> km | 
                    <strong>Fin de Jornada:</strong> <?= formatoHora($res["tiempoTotal"]) ?> | 
                    <strong>Tiempo de Espera:</strong> <?= formatoHora($res["tiempoEspera"]) ?>
                </p>

                <table>
                    <tr>
                        <th>Cliente</th>
                        <th>Llegada</th>
                        <th>Fin servicio</th>
                    </tr>
                    <?php foreach ($res["detalle"] as $d): ?>
                        <tr>
                            <td><?= htmlspecialchars($d["nombre"]) ?></td>
                            <td><?= formatoHora($d["llegada"]) ?></td>
                            <td><?= formatoHora($d["finServicio"]) ?></td>
                        </tr>
                    <?php endforeach; ?>
                </table>

                <a href="<?= $link ?>" target="_blank" class="boton">🗺️ Abrir en Google Maps</a>
            </div>

        <?php endforeach; ?>

        <!-- Script para iniciar la descarga del Excel automáticamente -->
        <?php if ($archivoExcelGenerado): ?>
            <script>
                window.addEventListener('DOMContentLoaded', function() {
                    window.location.href = '?descargar=<?= urlencode($archivoExcelGenerado) ?>';
                });
            </script>
        <?php endif; ?>

    <?php endif; ?>

</div>

</body>
</html>