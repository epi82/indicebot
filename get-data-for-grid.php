<?
include ("config.inc.php");

$genere=($_POST['genere']!="")? $_POST['genere'] :"";
$specie=($_POST['specie']!="")? $_POST['specie'] :"";
$autore=($_POST['autore']!="")? $_POST['autore'] :"";
$scheda=($_POST['scheda']!="")? $_POST['scheda'] :"";
$comme=($_POST['comme']!="")? $_POST['comme'] :"";
$famiglia=($_POST['famiglia']!="")? $_POST['famiglia'] :"";
$nomecomune=($_POST['nomecomune']!="")? $_POST['nomecomune'] :"";
$datains=($_POST['datains']!="")? $_POST['datains'] :"";
$start=($_POST['start']!="")? $_POST['start'] :"";
$limit=($_POST['limit']!="")? $_POST['limit'] :"";
$dir=($_POST['dir']!="")? $_POST['dir'] :"";
$sort=($_POST['sort']!="")? $_POST['sort'] :"";

$db = mysql_connect($db_host, $db_user, $db_password);

if ($db == FALSE)
  die ("Errore nella connessione. Verificare i parametri nel file config.inc.php");
mysql_select_db($db_name, $db)
or die ("Errore nella selezione del database. Verificare i parametri nel file config.inc.php");


$query_where .= " WHERE genere LIKE '".$genere."%'";

if ($specie!="") 
  {
    $query_where .= " AND specie LIKE '".$specie."%'";
  }

if ($autore!="") 
  {
    $query_where .= " AND autore LIKE '".$autore."%'";
  }

if ($comme!="" AND $comme!="-") 
  {
    $query_where .= " AND ca='".$comme."'";
  }  
  
if ($scheda!="" AND $scheda!="-") 
  {
    $query_where .= " AND scheda='".$scheda."'";
  }    
  
if ($famiglia!="") 
  {
    $query_where .= " AND famiglia LIKE '".$famiglia."%'";
  }  

if ($nomecomune!="") 
  {
    $query_where .= " AND nomecomune LIKE '%".$nomecomune."%'";
  }  

$query_count="SELECT count(*) FROM archivio";
$query_count .= $query_where;
$sql = mysql_query($query_count) or die(mysql_error());
$count = mysql_result($sql, "0"); 


$query = "SELECT id,scheda,topic,genere,specie,autore,ca,famiglia,nomecomune,datains FROM archivio";
$query .= $query_where;

if ($sort!="")
  {
   $query .= " ORDER BY ".$sort." ".$dir;  
  }

$query .= " LIMIT ".$start.",".$limit;
 
$json = array();
$i = 0;

$result = mysql_query($query, $db);
while ($row = mysql_fetch_array($result))
    {
      $id_db=trim($row[id]);  
      $sch_db=trim($row[scheda]);
      $top_db=trim($row[topic]);      
      $gen_db=htmlentities(trim($row[genere]));
      $spe_db=htmlentities(trim($row[specie]));
      $aut_db=htmlentities(trim($row[autore]));
      $ca_db=trim($row[ca]);
      $fam_db=htmlentities(trim($row[famiglia]));
      $nom_db=htmlentities(trim($row[nomecomune]));
      $dat_db=trim($row[datains]);            
                                                                                     
        $json[]= array(
          "id" => $id_db,
          "scheda" => $sch_db,
          "topic" => $top_db,
          "genere" => $gen_db,
          "specie" => $spe_db,
          "autore" => $aut_db,
          "comme" => $ca_db,
          "famiglia" => $fam_db,
          "nomecomune" => $nom_db,
          "datains" => $dat_db                    
        );        
    }

echo json_encode(array("totalCount" => $count,"rows" => $json));    
    
mysql_close($db); 

?>