Ext.onReady(function(){ 
  
  Ext.QuickTips.init();
  Ext.form.Field.prototype.msgTarget = 'side';
 
  var rd_random_employee_data = new Ext.data.JsonReader({}, ['famiglia']);
    
  var ds_random_employee_data_active = new Ext.data.Store({  
  proxy: new Ext.data.HttpProxy({  
      url: 'get-famiglia-from-db-archivio.php'  
  }),  
  reader: rd_random_employee_data,  
  remoteSort: false  
  });
  
  var ModFam = new Ext.form.FormPanel
  ({
 
        standardSubmit: true,
        frame:true,
        title: 'Modifica Scheda',
        bodyStyle:'padding:15px 5px 0',
        width: 500,
        defaults: {width: 300},
        defaultType: 'textfield',
        
        items: [
                {
                    name: 'gen',
                    fieldLabel: 'Genere',  
                    mode: 'remote',
                    minChars: 0,
                    typeAhead: true,
                    triggerAction: 'all',
                    value: genere,                   
                    style:'background-color:transparent; background-image: none;',
                    readOnly: true     
                },
                {
                    name: 'spe',
                    fieldLabel: 'specie',   
                    mode: 'remote',
                    minChars: 0,
                    typeAhead: true,
                    triggerAction: 'all',
                    value: specie,                                      
                    style:'background-color:transparent; background-image: none;',
                    readOnly: true                 
                },
                {
                    name: 'autore',
                    fieldLabel: 'Autore',  
                    mode: 'remote',
                    minChars: 0,
                    typeAhead: true,
                    triggerAction: 'all',
                    value: autore,
                    selectOnFocus:true,
                    forceSelection:true,
                    allowBlank: false,
                    blankText: 'Campo obbligatorio'                    
                },
                {
                    id:             'id_tipo_scheda',
                    xtype:          'combo',
                    mode:           'local',                    
                    triggerAction:  'all',
                    forceSelection: true,
                    editable:       false,
                    allowBlank:     false,
                    blankText:      'Campo obbligatorio',
                    fieldLabel:     'Scheda',
                    name:           'combo_scheda',
                    value:          scheda,
                    displayField:   'name',
                    valueField:     'value',
                    hiddenName :    'scheda',
                    emptyText:'Seleziona tipo scheda...',
                    store:          new Ext.data.JsonStore({
                        fields : ['name', 'value'],
                        data   : [
                            {name : 'Scheda AMINT',   value: 'A'},
                            {name : 'Scheda Breve',  value: 'B'},
                            {name : 'Scheda Sinonimo', value: 'S'}
                        ]
                    })                  
                },
                {
                    name: 'topic',
                    fieldLabel: 'Topic',  
                    mode: 'remote',
                    minChars: 0,
                    typeAhead: true,
                    triggerAction: 'all',
                    value: topic,
                    selectOnFocus:true,
                    forceSelection:true,
                    allowBlank: false,
                    blankText: 'Campo obbligatorio'                    
                },
                {
                    name: 'fam',
                    xtype: 'combo',
                    store: ds_random_employee_data_active,
                    fieldLabel: 'Famiglia',  
                    displayField: 'famiglia',
                    value: famiglia,
                    emptyText:'Seleziona famiglia...',  
                    mode: 'remote',
                    minChars: 0,
                    typeAhead: true,
                    triggerAction: 'all',
                    selectOnFocus:true,
                    forceSelection:true,
                    allowBlank: false,
                    blankText: 'Campo obbligatorio'
                },
                {
                    id:             'id_comme',
                    xtype:          'combo',
                    mode:           'local',                    
                    triggerAction:  'all',
                    forceSelection: true,
                    editable:       false,
                    allowBlank:     false,
                    blankText:      'Campo obbligatorio',
                    fieldLabel:     'Commestibilit&agrave;',
                    name:           'combo_comme',
                    displayField:   'name',
                    valueField:     'value',
                    value:          comme,
                    hiddenName :    'comme',
                    emptyText:'Seleziona commestibilita...',
                    store:          new Ext.data.JsonStore({
                        fields : ['name', 'value'],
                        data   : [
                            {name : 'Piante officinali - Aromatiche e cosmetiche',   value: 'PO'},
                            {name : 'Piante commestibili',  value: 'PC'},
                            {name : 'Piante commestibili e officinali', value: 'PCO'},
                            {name : 'Piante velenose', value: 'PV'},
                            {name : 'Piante velenose e officinali', value: 'PVO'},
                            {name : 'Pianta senza indicazioni di proprieta', value: 'P'}
                        ]
                    })                  
                },
                {
                    name: 'nom_comune',
                    fieldLabel: 'Nome Comune',  
                    mode: 'remote',
                    minChars: 0,
                    typeAhead: true,
                    triggerAction: 'all',
                    value: nom_comune,
                    selectOnFocus:true,
                    forceSelection:true,
                    allowBlank: false,
                    blankText: 'Campo obbligatorio'                    
                },
                {
                    xtype     : 'datefield',
                    name      : 'data',
                    fieldLabel: 'Data',
                    value: data,
                    format: 'd-m-Y',
                    altFormats: 'd-m-Y', 
                    allowBlank: false,
                    blankText: 'Campo obbligatorio',
                    invalidText: 'Data non valida - Formato richiesto: (gg-mm-aaaa)',
                    emptyText:'Seleziona data aggiornamento...'                      
                },                                
                {
                    inputType: 'hidden',
                    id: 'submitbutton',
                    name: 'myhiddenbutton',
                    value: 'mod_scheda'
                },
                {
                    inputType: 'hidden',
                    id: 'submitbutton2',
                    name: 'id_scheda',
                    value: id_scheda
                }                
              ],
        buttons: [{
                      text: 'Modifica',
                      handler: function() {
                      ModFam.getForm().getEl().dom.action = 'save_carica.php';
                      ModFam.getForm().getEl().dom.method = 'POST';
                      ModFam.getForm().submit();
                  }
                }]
 
  });

if ($("#formModScheda").length>0)
  {
    ModFam.render('formModScheda');
  }
    
});