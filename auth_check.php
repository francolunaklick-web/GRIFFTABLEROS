<?php
// Verificacion de sesion para todas las paginas protegidas.
// Si no esta logueado, redirige a login.php (calcula la URL relativa al root).
if (session_status() === PHP_SESSION_NONE) {
    session_start();
}
if (empty($_SESSION['logueado'])) {
    // Calcular la URL del login segun cuantos niveles abajo del root estamos.
    // El root es donde vive este archivo.
    $rootDir = realpath(__DIR__);
    $self    = realpath($_SERVER['SCRIPT_FILENAME'] ?? __FILE__);
    $depth   = 0;
    if ($self && $rootDir) {
        $rel = str_replace($rootDir, '', $self);
        $depth = substr_count(trim($rel, '/\\'), DIRECTORY_SEPARATOR);
        if ($depth === 0) {
            $depth = substr_count(trim($rel, '/\\'), '/');
        }
    }
    $up = str_repeat('../', max(0, $depth));
    header('Location: ' . $up . 'login.php');
    exit;
}
