<?php
session_start();
if (!empty($_SESSION['logueado'])) {
    header('Location: INICIO.php');
} else {
    header('Location: login.php');
}
exit;
