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

$id_fam=$_POST['FamId'];

$query = "SELECT famiglia FROM famiglia WHERE id=$id_fam";

if ($id_fam!='')
    {
        $result = mysql_query($query, $db);
        while ($row = mysql_fetch_array($result))
            {
              $famiglia=$row[famiglia];
            }
        mysql_close($db); 
    }
 
?>
<script type="text/javascript">var nome_fam="<? echo $famiglia;?>"</script>
<script type="text/javascript">var id="<? echo $id_fam;?>"</script>  
<script src="../js/jquery.min.js" type="text/javascript"></script>
<!-- Include Ext and app-specific scripts: -->
<script type="text/javascript" src="../adapter/ext/ext-base.js"></script>
<script type="text/javascript" src="../ext-all.js"></script>
<!--<script type="text/javascript" src="../ext-all-debug.js"></script>-->
 
<!-- Include your own Javascript file here - adapt the filename to your filename-->
<script type="text/javascript" src="../js/FormEditFamiglia.js"></script>
<script type="text/javascript" src="../js/FormModFamiglia.js"></script>
 
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
 if ($id_fam=="")
    {
 ?>    
     <tr class="testo1">
         <td valign="top" width="954">
         <br />
            <div align="center">
                <div id="formEditFamiglia"></div>                
            </div>
        </td>
     </tr>
 <?
    }
else
    {
 ?> 
     <tr class="testo1">
         <td valign="top" width="954">
         <br />  
            <div align="center">
                <div id="formModFamiglia"></div>                
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