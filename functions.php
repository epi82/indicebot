<?
include ("config.inc.php");

function data(){
$db_host = "localhost";
$db_user = "curtiptr_bota";
$db_password = "bt;:Amint";
$db_name = "curtiptr_botanica";
$db = mysql_connect($db_host, $db_user, $db_password);

if ($db == FALSE)
	die ("Errore nella connessione. Verificare i parametri nel file config.inc.php");
mysql_select_db($db_name, $db)
or die ("Errore nella selezione del database. Verificare i parametri nel file config.inc.php");

$query = "SELECT scheda, genere, specie, autore, ca, famiglia, nomecomune, topic, datains FROM archivio ORDER BY datains DESC";

$result = mysql_query($query, $db);

$row = mysql_fetch_array($result);

return $row[datains];

}

function verifica($tabella,$campo,$valore){
$db_host = "localhost";
$db_user = "curtiptr_bota";
$db_password = "bt;:Amint";
$db_name = "curtiptr_botanica";

$presente=0;

 $db = mysql_connect($db_host, $db_user, $db_password);
  	if ($db == FALSE)
    	die ("Errore nella connessione. Verificare i parametri nel file config.inc.php");
  	mysql_select_db($db_name, $db)
    	or die ("Errore nella selezione del database. Verificare i parametri nel file config.inc.php");

	$controlla = "SELECT " . $campo . " FROM " . $tabella;
    $resultfami = mysql_query($controlla, $db);
    while ($fam = mysql_fetch_array($resultfami)){
  	if ($valore == $fam[$campo]) $presente=1;
  	}
	mysql_close($db);
	return $presente;
}

function verifica_genspe($tabella,$campo1,$campo2,$genere,$specie){
 
$db_host = "localhost";
$db_user = "curtiptr_bota";
$db_password = "bt;:Amint";
$db_name = "curtiptr_botanica";

$presente=0;

 $db = mysql_connect($db_host, $db_user, $db_password);
  	if ($db == FALSE)
    	die ("Errore nella connessione. Verificare i parametri nel file config.inc.php");
  	mysql_select_db($db_name, $db)
    	or die ("Errore nella selezione del database. Verificare i parametri nel file config.inc.php");

	$controlla = "SELECT " . $campo1 . " FROM " . $tabella;
    $resultfami = mysql_query($controlla, $db);
    while ($fam = mysql_fetch_array($resultfami)){
  	if ($genere == $fam[$campo1]) $gen=$genere;
	}
	
	if($gen == $genere){
	$controlla2 = "SELECT " . $campo2 . " FROM " . $tabella . " WHERE " . $campo1 . " LIKE '" . $gen . "'";
    $resultfami2 = mysql_query($controlla2, $db);
    while ($fam2 = mysql_fetch_array($resultfami2)){
  	if ($specie == $fam2[$campo2]) $presente=1;
  	}
	}
	
	mysql_close($db);
	
	return $presente;
}

?>