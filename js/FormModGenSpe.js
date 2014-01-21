Ext.onReady(function(){ 
 
  var simple = new Ext.form.FormPanel
  ({
 
        standardSubmit: true,
        frame:true,
        title: 'Modifica Genere-specie',
        bodyStyle:'padding:15px 5px 0',
        width: 500,
        defaults: {width: 300},
        defaultType: 'textfield',
        
        items: [
                {
                    name: 'mod_gen',
                    fieldLabel: 'Genere',  
                    selectOnFocus:true,
                    forceSelection:true,
                    value: gensel,
                    allowBlank: false,
                    blankText: 'Campo obbligatorio'                    
                },
                {
                    name: 'mod_spe',
                    fieldLabel: 'specie',  
                    selectOnFocus:true,
                    forceSelection:true,
                    value: spesel,
                    allowBlank: false,
                    blankText: 'Campo obbligatorio'                    
                },
                {
                    xtype: 'checkbox',
                    fieldLabel: '',
                    boxLabel: 'Aggiorna anche la tabella ARCHIVIO',
                    name: 'archivio',
                    value: '1',
                    checked: true
                },
                {
                    inputType: 'hidden',
                    id: 'submitbutton',
                    name: 'myhiddenbutton',
                    value: 'mod_genspe'
                },
                {
                    inputType: 'hidden',
                    id: 'submitbutton2',
                    name: 'id',
                    value: id_genspe
                },
                {
                    inputType: 'hidden',
                    id: 'submitbutton3',
                    name: 'gen_old',
                    value: gensel
                },
                {
                    inputType: 'hidden',
                    id: 'submitbutton4',
                    name: 'spe_old',
                    value: spesel
                }
              ],
        buttons: [{
                      text: 'Modifica',
                      handler: function() {
                      simple.getForm().getEl().dom.action = 'save_carica.php';
                      simple.getForm().getEl().dom.method = 'POST';
                      simple.getForm().submit();
                  }
                }]
 
  });

if ($("#formModGenSpe").length>0)
  {
    simple.render('formModGenSpe');
  }
    
});