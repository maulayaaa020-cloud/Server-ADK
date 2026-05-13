<?php
session_start();
session_unset();
session_destroy();
header('Location: /adk/index.html');
exit;
