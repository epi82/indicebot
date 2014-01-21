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
?>
<script type="text/javascript" src="../adapter/ext/ext-base.js"></script>
<script type="text/javascript" src="../ext-all.js"></script>
<script type="text/javascript" src="../js/panels.js"></script>
<script type="text/javascript" src="../js/genspe.js"></script>
<script type="text/javascript" src="../js/menu_fam.js"></script>
<link rel="stylesheet" type="text/css" href="../css/buttons.css"/>
<link rel="stylesheet" type="text/css" href="../resources/css/ext-all.css" />
<style type="text/css">
    .x-panel-body p {
        margin:10px;
    }
    .container {
        padding:3px;
    }
  em.cfg { font-style: italic; font-weight: bold;}
</style>

<table class="Table_Corpo_Admin" cellpadding="10" height="351" width="980">
	<tbody>
	   <tr class="testo1" height="327">
	       <td height="327" valign="top" width="954">
	       <br />
		        <div align="center">
                    <div id="panel-basic" class="container">                    
                    </div>
                    <div id="panel-genspe" class="container">                    
                    </div>
                    <div id="panel-fam" class="container">                    
                    </div>
            </div>
        </td>
     </tr>
  </tbody>
</table>
<?
foot();
?>
<br />
