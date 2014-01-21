<?
include ("config.inc.php");
include ("top_foot.inc.php");

//intestazione
top();

$forma=($_GET['forma']!="")? $_GET['forma'] :"";
$descrizione=($_GET['descrizione']!="")? $_GET['descrizione'] :"";
$antesi=($_GET['antesi']!="")? $_GET['antesi'] :"";
$tipo_coro=($_GET['tipo_coro']!="")? $_GET['tipo_coro'] :"";
$distribuzione=($_GET['distribuzione']!="")? $_GET['distribuzione'] :"";
$habitat=($_GET['habitat']!="")? $_GET['habitat'] :"";
$note_siste=($_GET['note_siste']!="")? $_GET['note_siste'] :"";
$etimologia=($_GET['etimologia']!="")? $_GET['etimologia'] :"";
$proprieta=($_GET['proprieta']!="")? $_GET['proprieta'] :"";
$curiosita=($_GET['curiosita']!="")? $_GET['curiosita'] :"";
$attenzione=($_GET['attenzione']!="")? $_GET['attenzione'] :"";
$fonti=($_GET['fonti']!="")? $_GET['fonti'] :"";
$autore=($_GET['autore']!="")? $_GET['autore'] :"";

?>

<br />
<br />

<table class="Table_Corpo" width="980">
<tr>
<td>
[b]Forma Biologica: [/b] <? echo "$forma"; ?><br /><br />
[b]Descrizione: [/b] <? echo "$descrizione"; ?><br /><br />
[b]Antesi: [/b] <? echo "$antesi"; ?><br /><br />
[b]Tipo corologico: [/b] <? echo "$tipo_coro"; ?><br /><br />
[b]Distribuzione in Italia: [/b] <? echo "$distribuzione"; ?><br /><br />
[b]Habitat: [/b] <? echo "$habitat"; ?><br /><br />
[b]Note di Sistematica: [/b] <? echo "$note_siste"; ?><br />
____________________________________________________________________<br />

</td>  
</tr>
</table>