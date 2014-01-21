<?
include ("../config.inc.php");
include ("../top_foot.inc.php");

//intestazione
$db = mysql_connect($db_host, $db_user, $db_password);

if ($db == FALSE)
  die ("Errore nella connessione. Verificare i parametri nel file config.inc.php");
mysql_select_db($db_name, $db)
or die ("Errore nella selezione del database. Verificare i parametri nel file config.inc.php");

top_admin();
link_admin();

$gen=$_POST['gen'];
$id_scheda=$_POST['SchedaId'];
 
?>
<script type="text/javascript">var gensel="<? echo $gen;?>"</script>
    
<script src="../js/jquery.min.js" type="text/javascript"></script>
<!-- Include Ext and app-specific scripts: -->
<script type="text/javascript" src="../adapter/ext/ext-base.js"></script>
<script type="text/javascript" src="../ext-all.js"></script>
<script type="text/javascript" src="../ext-all-debug.js"></script>
 
<!-- Include your own Javascript file here - adapt the filename to your filename-->
<script type="text/javascript" src="../js/FormModSchedaGenSpe.js"></script>
<script type="text/javascript" src="../js/FormModSchedaGenSpe_Step2.js"></script>    
 
<!-- Include Ext stylesheets here: -->
<link rel="stylesheet" type="text/css" href="../resources/css/ext-all.css" />
<link rel="stylesheet" type="text/css" href="../resources/css/xtheme-blue.css" />

<!-- ComboBox -->
<link rel="stylesheet" type="text/css" href="../css/combos.css" />

<style type="text/css">
        p { width:650px; }
</style>

<table class="Table_Corpo_Admin" cellpadding="10" height="351" width="980">
 <?
 if ($gen=="" AND $id_scheda=="")
    {
 ?>    
     <tr class="testo1">
         <td valign="top" width="954">
         <br />
            <div align="center">
                <div id="formModSchedaGenSpe"></div>                
            </div>
        </td>
     </tr>
 <?
    }
 elseif ($gen!="" AND $id_scheda=="")
    {
 ?> 
     <tr class="testo1">
         <td valign="top" width="954">
         <br />  
            <div align="center">
                <div id="formModSchedaGenSpe_Step2"></div>               
            </div>
        </td>
     </tr>
 <?
    }
 elseif ($id_scheda!="")
    {
      
        $query = "SELECT * FROM archivio WHERE id=$id_scheda";
        $result = mysql_query($query, $db);
        while ($row = mysql_fetch_array($result))
            {
              $topic=$row[topic];
              $scheda=$row[scheda];
              $genere=$row[genere];
              $specie=$row[specie];
              $autore=$row[autore];
              $comme=$row[ca];
              $famiglia=$row[famiglia];
              $nom_comune=$row[nomecomune];
              $datains=$row[datains];
            }
        mysql_close($db);
        
        $anno=substr($datains, -8, 4);
        $mese=substr($datains, -4, 2);
        $giorno=substr($datains, -2, 2);
        $data=$giorno.'-'.$mese.'-'.$anno;
    
        if ($scheda!='S')
        {            
 ?>            
      <script type="text/javascript">var topic="<? echo $topic;?>"</script>
      <script type="text/javascript">var scheda="<? echo $scheda;?>"</script>
      <script type="text/javascript">var genere="<? echo $genere;?>"</script>
      <script type="text/javascript">var specie="<? echo $specie;?>"</script>
      <script type="text/javascript">var autore="<? echo $autore;?>"</script>
      <script type="text/javascript">var comme="<? echo $comme;?>"</script>
      <script type="text/javascript">var famiglia="<? echo $famiglia;?>"</script>
      <script type="text/javascript">var nom_comune="<? echo $nom_comune;?>"</script>
      <script type="text/javascript">var data="<? echo $data;?>"</script>            
      <script type="text/javascript">var id_scheda="<? echo $id_scheda;?>"</script>
      <script type="text/javascript" src="../js/FormModScheda.js"></script>
     
     <tr class="testo1">
         <td valign="top" width="954">
         <br />  
            <div align="center">
                <div id="formModScheda"></div>            
            </div>
        </td>
     </tr>
 <?
        } 
        else
        {
 ?>       
   
      <script type="text/javascript">var topic="<? echo $topic;?>"</script>
      <script type="text/javascript">var scheda="<? echo $scheda;?>"</script>
      <script type="text/javascript">var genere="<? echo $genere;?>"</script>
      <script type="text/javascript">var specie="<? echo $specie;?>"</script>
      <script type="text/javascript">var autore="<? echo $autore;?>"</script>
      <script type="text/javascript">var data="<? echo $data;?>"</script>            
      <script type="text/javascript">var id_scheda="<? echo $id_scheda;?>"</script>
      <script type="text/javascript" src="../js/FormModSchedaSino.js"></script>
     
     <tr class="testo1">
         <td valign="top" width="954">
         <br />  
            <div align="center">
                <div id="formModSchedaSino"></div>               
            </div>
        </td>
     </tr>   
   
 <?         
        }  
    }    
 ?>     
</table>
<?
foot();
?>
<br />