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
$spe=$_POST['spe'];
$step=$_POST['step'];
$autore=$_POST['autore'];
$scheda=$_POST['scheda'];
$descr_scheda=$_POST['descr_scheda'];
$topic=$_POST['topic'];
$data=$_POST['data'];
$fam=$_POST['fam'];
$comme=$_POST['comme'];
$nom_comune=$_POST['nom_comune'];
 
?>
<script type="text/javascript">var gensel="<? echo $gen;?>"</script>
<script type="text/javascript">var spesel="<? echo $spe;?>"</script>
<script type="text/javascript">var autore="<? echo $autore;?>"</script>
<script type="text/javascript">var scheda="<? echo $scheda;?>"</script>
<script type="text/javascript">var topic="<? echo $topic;?>"</script>
<script type="text/javascript">var data="<? echo $data;?>"</script>
<script type="text/javascript">var fam="<? echo $fam;?>"</script>
<script type="text/javascript">var nom_comune="<? echo $nom_comune;?>"</script>
<script type="text/javascript">var comme="<? echo $comme;?>"</script>
    
<script src="../js/jquery.min.js" type="text/javascript"></script>
<!-- Include Ext and app-specific scripts: -->
<script type="text/javascript" src="../adapter/ext/ext-base.js"></script>
<script type="text/javascript" src="../ext-all.js"></script>
<script type="text/javascript" src="../ext-all-debug.js"></script>
 
<!-- Include your own Javascript file here - adapt the filename to your filename-->
<script type="text/javascript" src="../js/FormSchedaGenSpe.js"></script>
<script type="text/javascript" src="../js/FormSchedaGenSpe_Step2.js"></script>    
 
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
 if ($gen=="")
    {
 ?>    
     <tr class="testo1">
         <td valign="top" width="954">
         <br />
            <div align="center">
                <div id="formSchedaGenSpe"></div>                
            </div>
        </td>
     </tr>
 <?
    }
 elseif ($gen!="" AND $spe=="")
    {
 ?> 
     <tr class="testo1">
         <td valign="top" width="954">
         <br />  
            <div align="center">
                <div id="formSchedaGenSpe_Step2"></div>                
            </div>
        </td>
     </tr>
 <?
    }
 elseif ($gen!="" AND $spe!="" AND $step=="")
    {
 ?>            
     <script type="text/javascript" src="../js/FormSchedaStep1.js"></script>
     
     <tr class="testo1">
         <td valign="top" width="954">
         <br />  
            <div align="center">
                <div id="formSchedaStep1"></div>               
            </div>
        </td>
     </tr>
 <?
    }
 elseif ($gen!="" AND $spe!="" AND $step=="1" AND $scheda!='S')
    {
     switch ($scheda)
        {
        case 'S':
          $descr_scheda='Scheda Sinonimo';
        break;
        case 'A':
          $descr_scheda='Scheda AMINT';
        break;
        case 'B':
          $descr_scheda='Scheda Breve';
        break;        
        }
 ?>     
     <script type="text/javascript">var descr_scheda="<? echo $descr_scheda;?>"</script>
     <script type="text/javascript" src="../js/FormSchedaStep2.js"></script>
     
     <tr class="testo1">
         <td valign="top" width="954">
         <br />  
            <div align="center">
                <div id="formSchedaStep2"></div>               
            </div>
        </td>
     </tr> 
 <?
    }
 elseif ($gen!="" AND $spe!="" AND $step=="2" AND $scheda!='S')
    {
    switch ($comme)
        {
        case 'PO':
          $descr_comme='Piante officinali - Aromatiche e cosmetiche';
        break;
        case 'PC':
          $descr_comme='Piante commestibili';
        break;
        case 'PCO':
          $descr_comme='Piante commestibili e officinali';
        break;
        case 'PV':
          $descr_comme='Piante velenose';
        break;
        case 'PVO':
          $descr_comme='Piante velenose e officinali';
        break;
        case 'P':
          $descr_comme='Pianta senza indicazioni di proprieta';
        break;                 
        }
 ?>          
     <script type="text/javascript">var descr_scheda="<? echo $descr_scheda;?>"</script>
     <script type="text/javascript">var descr_comme="<? echo $descr_comme;?>"</script>
     <script type="text/javascript" src="../js/FormSchedaEnd.js"></script>
 
     <tr class="testo1">
         <td valign="top" width="954">
         <br />  
            <div align="center">
                <div id="formSchedaEnd"></div>                                
            </div>
        </td>
     </tr> 
 
 <?
    }
 elseif ($gen!="" AND $spe!="" AND $step=="1" AND $scheda=='S')
    {
     switch ($scheda)
        {
        case 'S':
          $descr_scheda='Scheda Sinonimo';
        break;
        case 'A':
          $descr_scheda='Scheda AMINT';
        break;
        case 'B':
          $descr_scheda='Scheda Breve';
        break;        
        }          
 ?>
     <script type="text/javascript">var descr_scheda="<? echo $descr_scheda;?>"</script>
     <script type="text/javascript" src="../js/FormSchedaSinoEnd.js"></script>
 
     <tr class="testo1">
         <td valign="top" width="954">
         <br />  
            <div align="center">
                <div id="formSchedaSinoEnd"></div>                             
            </div>
        </td>
     </tr>
 
 <?
    }
 ?>         
</table>
<?
foot();
?>
<br />