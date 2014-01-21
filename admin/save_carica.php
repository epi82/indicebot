<?
include("../top_foot.inc.php");
include("../config.inc.php");
include("../functions.php");
top_admin();
link_admin();

$fam=$_POST['fam'];
$mod_fam=$_POST['mod_fam'];
$mod_gen=$_POST['mod_gen'];
$mod_spe=$_POST['mod_spe'];
$tipo=$_POST['myhiddenbutton'];
$id=$_POST['id'];
$id_scheda=$_POST['id_scheda'];
$gen=$_POST['gen'];
$spe=$_POST['spe'];
$gen_old=$_POST['gen_old'];
$spe_old=$_POST['spe_old'];
$fam_old=$_POST['fam_old'];
$archivio=$_POST['archivio'];
$autore=$_POST['autore'];
$topic=$_POST['topic'];
$data=$_POST['data'];
$nom_comune=$_POST['nom_comune'];
$scheda=$_POST['scheda'];
$comme=$_POST['comme'];

switch ($tipo)
  {
  case 'new_fam':
      $tabella="famiglia";
      $campo="famiglia";
      $valore=$fam;
      $presente=verifica($tabella,$campo,$valore);
        
        if (trim($fam) == "")
          $stampa2="Il campo Famiglia deve essere riempito!";
        elseif ($presente == '1') $stampa2="Famiglia presente in archivio per la tabella [Famiglia]";
        else 
        {
          $fam = addslashes(stripslashes($fam));
          $fam = str_replace("<", "&lt;", $fam);
          $fam = str_replace(">", "&gt;", $fam);
          
          $db = mysql_connect($db_host, $db_user, $db_password);
        
          if ($db == FALSE)
            die ("Errore nella connessione. Verificare i parametri nel file config.inc.php");
      
          mysql_select_db($db_name, $db)
            or die ("Errore nella selezione del database. Verificare i parametri nel file config.inc.php");
      
          $query = "INSERT INTO famiglia (famiglia) VALUES ('$fam')";
            if (mysql_query($query, $db))
            $stampa='Famiglia inserita correttamente';
            else
            $stampa='Errore inserimento dati!!!';
        mysql_close($db);
        }
  break;
  case 'mod_fam':       
        if (trim($mod_fam) == "")
          $stampa2="Il campo Famiglia deve essere riempito!";
        else 
        {
          $mod_fam = addslashes(stripslashes($mod_fam));
          $mod_fam = str_replace("<", "&lt;", $mod_fam);
          $mod_fam = str_replace(">", "&gt;", $mod_fam);
          
          $db = mysql_connect($db_host, $db_user, $db_password);
        
          if ($db == FALSE)
            die ("Errore nella connessione. Verificare i parametri nel file config.inc.php");
      
          mysql_select_db($db_name, $db)
            or die ("Errore nella selezione del database. Verificare i parametri nel file config.inc.php");
      
          $query = "UPDATE famiglia SET famiglia='$mod_fam' WHERE id='$id'";
            if (mysql_query($query, $db))
            $stampa='Famiglia modificata correttamente';
            else
            $stampa='Errore modifica dati!!!';
            
          if ($archivio=='on')
          {
            $query_archivio="UPDATE archivio SET famiglia='$mod_fam' WHERE famiglia='$fam_old'";
              if (mysql_query($query_archivio, $db))
              $stampa2='Modifica tabella Archivio avvenuta con successo.';
              else
              $stampa2='Errore aggiornamento tabella Archivio!!!';              
          }
                      
        mysql_close($db);
        }  
  break;
  case 'new_genspe':
      $tabella="genspe";
      $campo1="genere";
      $campo2="specie";
      $presente=verifica_genspe($tabella,$campo1,$campo2,$gen,$spe);
      
      if (trim($gen) == "" OR trim($spe) == "" OR trim($gen)=="Inserisci Genere..." OR trim($spe) == "Inserisci specie..."){$stampa2="I campi Genere, Specie devono essere riempiti!";}
      elseif ($presente == '1') $stampa2="Genere-specie presenti in archivio";
        else {
        $gen = addslashes(stripslashes($gen));
        $spe = addslashes(stripslashes($spe));
        $gen = str_replace("<", "&lt;", $gen);
        $gen = str_replace(">", "&gt;", $gen);
        $spe = str_replace("<", "&lt;", $spe);
        $spe = str_replace(">", "&gt;", $spe);
        $db = mysql_connect($db_host, $db_user, $db_password);
        if ($db == FALSE)
          die ("Errore nella connessione. Verificare i parametri nel file config.inc.php");
    
        mysql_select_db($db_name, $db)
          or die ("Errore nella selezione del database. Verificare i parametri nel file config.inc.php");
    
      $query = "INSERT INTO genspe (genere, specie) VALUES ('$gen', '$spe')";
          if (mysql_query($query, $db))
          $stampa='Genere-specie inseriti correttamente';
          else
        $stampa='Errore inserimento dati!!!';
      mysql_close($db);
        }            
  break;
  case 'mod_genspe':
          $mod_gen = addslashes(stripslashes($mod_gen));
          $mod_gen = str_replace("<", "&lt;", $mod_gen);
          $mod_gen = str_replace(">", "&gt;", $mod_gen);
          $mod_spe = addslashes(stripslashes($mod_spe));
          $mod_spe = str_replace("<", "&lt;", $mod_spe);
          $mod_spe = str_replace(">", "&gt;", $mod_spe);
          
          $db = mysql_connect($db_host, $db_user, $db_password);
        
          if ($db == FALSE)
            die ("Errore nella connessione. Verificare i parametri nel file config.inc.php");
      
          mysql_select_db($db_name, $db)
            or die ("Errore nella selezione del database. Verificare i parametri nel file config.inc.php");
      
          $query = "UPDATE genspe SET genere='$mod_gen',specie='$mod_spe' WHERE id='$id'";
            if (mysql_query($query, $db))
            $stampa='Genere-specie modificato correttamente';
            else
            $stampa='Errore modifica dati!!!';
            
          if ($archivio=='on')
          {
            $query_archivio="UPDATE archivio SET genere='$mod_gen', specie='$mod_spe' WHERE (genere='$gen_old' AND specie='$spe_old')";
              if (mysql_query($query_archivio, $db))
              $stampa2='Modifica tabella Archivio avvenuta con successo.';
              else
              $stampa2='Errore aggiornamento tabella Archivio!!!';              
          }
        mysql_close($db);
  break;
  case 'new_scheda':                  
          $anno=substr($data, -4);
          $mese=substr($data, -7, 2);
          $giorno=substr($data, -10, 2);
          $datains="$anno$mese$giorno";
          
          $db = mysql_connect($db_host, $db_user, $db_password);
        
          if ($db == FALSE)
            die ("Errore nella connessione. Verificare i parametri nel file config.inc.php");
      
          mysql_select_db($db_name, $db)
            or die ("Errore nella selezione del database. Verificare i parametri nel file config.inc.php");
      
          $query = "INSERT INTO archivio (topic,scheda,genere,specie,autore,ca,famiglia,nomecomune,datains) VALUES ('$topic','$scheda','$gen','$spe','$autore','$comme','$fam','$nom_comune','$datains')";
            if (mysql_query($query, $db))
            $stampa='Nuova scheda inserita correttamente';
            else
            $stampa='Errore inserimento dati dati!!!';
            
        mysql_close($db);
  break;
  case 'new_scheda_sino':                  
          $anno=substr($data, -4);
          $mese=substr($data, -7, 2);
          $giorno=substr($data, -10, 2);
          $datains="$anno$mese$giorno";
          
          $db = mysql_connect($db_host, $db_user, $db_password);
        
          if ($db == FALSE)
            die ("Errore nella connessione. Verificare i parametri nel file config.inc.php");
      
          mysql_select_db($db_name, $db)
            or die ("Errore nella selezione del database. Verificare i parametri nel file config.inc.php");
      
          $query = "INSERT INTO archivio (topic,scheda,genere,specie,autore,datains) VALUES ('$topic','$scheda','$gen','$spe','$autore','$datains')";
            if (mysql_query($query, $db))
            $stampa='Nuova scheda inserita correttamente';
            else
            $stampa='Errore inserimento dati dati!!! Query:'.$query;
            
        mysql_close($db);
  break;
  case 'mod_scheda':         
          $anno=substr($data, -4);
          $mese=substr($data, -7, 2);
          $giorno=substr($data, -10, 2);
          $datains="$anno$mese$giorno";    
          $autore = addslashes(stripslashes($autore));
          $topic = addslashes(stripslashes($topic));
          $nom_comune = addslashes(stripslashes($nom_comune));
          
          $db = mysql_connect($db_host, $db_user, $db_password);
        
          if ($db == FALSE)
            die ("Errore nella connessione. Verificare i parametri nel file config.inc.php");
      
          mysql_select_db($db_name, $db)
            or die ("Errore nella selezione del database. Verificare i parametri nel file config.inc.php");
      
          $query = "UPDATE archivio SET autore='$autore', topic='$topic', scheda='$scheda', famiglia='$fam', ca='$comme', nomecomune='$nom_comune', datains='$datains' WHERE id='$id_scheda'";
            if (mysql_query($query, $db))
            $stampa='Scheda modificata correttamente';
            else
            $stampa='Errore modifica dati!!! Query='.$query;
            
        mysql_close($db);
  break;
  case 'mod_scheda_sino':         
          $anno=substr($data, -4);
          $mese=substr($data, -7, 2);
          $giorno=substr($data, -10, 2);
          $datains="$anno$mese$giorno";    
          $autore = addslashes(stripslashes($autore));
          $topic = addslashes(stripslashes($topic));
          
          $db = mysql_connect($db_host, $db_user, $db_password);
        
          if ($db == FALSE)
            die ("Errore nella connessione. Verificare i parametri nel file config.inc.php");
      
          mysql_select_db($db_name, $db)
            or die ("Errore nella selezione del database. Verificare i parametri nel file config.inc.php");
      
          $query = "UPDATE archivio SET autore='$autore', topic='$topic', scheda='$scheda', datains='$datains' WHERE id='$id_scheda'";
            if (mysql_query($query, $db))
            $stampa='Scheda modificata correttamente';
            else
            $stampa='Errore modifica dati!!! Query='.$query;
            
        mysql_close($db);
  break;
  case 'del_scheda':         
          
          $db = mysql_connect($db_host, $db_user, $db_password);
        
          if ($db == FALSE)
            die ("Errore nella connessione. Verificare i parametri nel file config.inc.php");
      
          mysql_select_db($db_name, $db)
            or die ("Errore nella selezione del database. Verificare i parametri nel file config.inc.php");
      
          $query = "DELETE FROM archivio WHERE id='$id_scheda'";
            if (mysql_query($query, $db))
            $stampa='<i>Id=</i>'.$id_scheda.' <i>Genere=</i>'.$gen.' <i>specie=</i>'.$spe.' <i>Autore=</i>'.$autore.'<br />Scheda cancellata correttamente';
            else
            $stampa='Errore cancellazione dati!!! Query='.$query;
            
        mysql_close($db);
  break;                                                                                                                   
  }    

?>

<script type="text/javascript">var stampa1="<? echo $stampa;?>"</script>
<script type="text/javascript">var stampa2="<? echo $stampa2;?>"</script>
<link rel="stylesheet" type="text/css" href="../resources/css/ext-all.css" />
<script type="text/javascript" src="../adapter/ext/ext-base.js"></script>
<script type="text/javascript" src="../ext-all.js"></script>
<script type="text/javascript" src="../js/msg-box.js"></script>
<script type="text/javascript" src="../js/FormShowResults.js"></script>
<style type="text/css">
    .x-window-dlg .ext-mb-download {
        background:transparent url(../images/download.gif) no-repeat top left;
        height:46px;
</style>
<script type="text/javascript">
//setTimeout("('#TableResults').show()",3);
setTimeout("ShowResults()", 3000);

function ShowResults() {
document.getElementById("TableResults").style.display = 'block';
}
</script>
<table class="Table_Corpo" width="980" cellpadding="10" height="351">
    <tr class="testo1" height="327">
       <td valign="top" bgcolor="#e6e6fa" width="954" height="327">
           <div id="TableResults" style= "display:none;" align="center">
                   <table width="940" border="0" cellspacing="0" cellpadding="3" align="center" height="50">
                      <tbody>                    
                         <tr class="testo1">
                             <td valign="top" width="954">
                             <br />  
                                <div align="center">
                                    <div id="formShowResults"></div>                
                                </div>
                            </td>
                         </tr>
                      </tbody>
                  </table>
           </div>
       </td>
    </tr>                   
</table>

<?
// chiude la verifica della presenza dei dati

foot();
?>