Ext.onReady(function(){            
    
    var store = new Ext.data.JsonStore({
      url: "get-data-for-grid.php",
      root: "rows",
      id:'id',   
      totalProperty:'totalCount',      
      //baseParams: {limit: 15},
      remoteSort: true,
      sortInfo: {
            field: 'genere,specie',
            direction: 'ASC'
      },      
      autoDestroy: true,
      fields: [
         {name: 'scheda'},
         {name: 'topic'},
         {name: 'genere'},
         {name: 'specie'},
         {name: 'autore'},
         {name: 'comme'},
         {name: 'famiglia'},
         {name: 'nomecomune'},
         {name: 'datains'}
      ]
    });
      
    var filterRow = new Ext.ux.grid.FilterRow({
      autoFilter: false,
      listeners: {
        change: function(data) {
          store.baseParams = data;
          store.load({
            params: {start: 0, limit: 15}
          });
        }
      }
    }); 
    
    var cellTips = new Ext.ux.plugins.grid.CellToolTips([
      { field: 'genere', tpl: '<b>Genere: <i>{genere}</i></b><br />Specie: <i>{specie}</i><br />Famiglia: <i>{famiglia}</i><br />' },
      { field: 'comme', tpl: '<tpl if="comme==\'P\'"><b>P = Pianta senza indicazioni di propriet&agrave;</b></tpl><tpl if="comme==\'PO\'"><b>PO = Piante officinali - Aromatiche e cosmetiche</b></tpl><tpl if="comme==\'PC\'"><b>PC = Piante commestibili</b></tpl><tpl if="comme==\'PCO\'"><b>PCO = Piante commestibili e officinali</b></tpl><tpl if="comme==\'PV\'"><b>PV = Piante velenose</b></tpl><tpl if="comme==\'PVO\'"><b>PVO = Piante velenose e officinali</b></tpl>'}
    ]);    
    
    var panelResizer = new Ext.ux.PanelResizer({ minHeight: 200 });
    
    var highlightSort = new Ext.ux.plugins.grid.highlightSort({});
    
    var tbar = new Ext.Toolbar({
        items: 
        [
          {xtype: 'tbseparator'},
          {
            text:'<div align="center" class="testoicone">Il Meraviglioso Mondo dei Funghi e dei Fiori Spontanei</div>',
            iconCls:'forum',
            listeners: { 
                render: function(button){ 
                    button.getEl().on('click', function(){ 
                        //alert('works in IE6, IE8, FF');
                        location.href='http://www.funghiitaliani.it/';                
                    }); 
                } 
            }       
          },
          {xtype: 'tbseparator'},          
          {
            text:'<div align="center" class="testoicone">Micologia</div>',
            iconCls:'mico',
            listeners: { 
                render: function(button){ 
                    button.getEl().on('click', function(){ 
                        //alert('works in IE6, IE8, FF');
                        location.href='http://www.funghiitaliani.it/micologia/indice.html';                
                    }); 
                } 
            }       
          },                    
          {xtype: 'tbseparator'},          
          {
            text:'<div align="center" class="testoicone">Botanica</div>',
            iconCls:'bota',
            listeners: { 
                render: function(button){ 
                    button.getEl().on('click', function(){ 
                        //alert('works in IE6, IE8, FF');
                        location.href='http://www.funghiitaliani.it/botanica/indice.html';                
                    }); 
                } 
            }       
          },          
          {xtype: 'tbseparator'},
          {
            text:'<div align="center" class="testoicone">Iscriviti AMINT</div>',      
            iconCls:'shop',
            listeners: { 
                render: function(button){ 
                    button.getEl().on('click', function(){ 
                        //alert('works in IE6, IE8, FF');
                        location.href='http://www.funghiitaliani.it/index.php?app=nexus';                
                    }); 
                } 
            }       
          },          
          {xtype: 'tbseparator'},
          {
            text:'<div align="center" class="testoicone">Chat AMINT</div>',            
            iconCls:'chat',
            listeners: { 
                render: function(button){ 
                    button.getEl().on('click', function(){ 
                        //alert('works in IE6, IE8, FF');
                        location.href='http://www.funghiitaliani.it/index.php?app=chat';                
                    }); 
                } 
            }       
          }                                                                    
        ]        
    });               

      var grid = new Ext.grid.GridPanel({
        store: store,
        loadMask: new Ext.LoadMask(Ext.getBody(), {msg:"Caricamento..."}),
        columns: [
            {id:'scheda', header: "Scheda", width: 73, align: 'center', sortable: false, renderer: scheda, dataIndex: 'scheda',
              filter: 
              {
                field: 
                {
                  xtype: "combo",
                  mode: 'local',
                  //store: new Ext.data.ArrayStore
                  store: new Ext.data.JsonStore
                  ({
                    id: 0,
                    fields: ['name','value'],
                    //data: [['-'], ['A'], ['B'], ['S']]
                    data: 
                    [
                        {name : '-',   value: '-'},
                        {name : 'AMINT',   value: 'A'},
                        {name : 'Breve',  value: 'B'},
                        {name : 'Sinonimo', value: 'S'}
                    ]
                  }),
                  valueField: 'value',
                  displayField: 'name',
                  triggerAction: 'all',
                  value: "-"
                },
                fieldEvents: ["select"],
                test: function(filterValue, value) 
                      {
                        return filterValue === "-" || filterValue === value;
                      }
               }            
            },            
            {id:'genere', header: "Genere", width: 130, renderer: gen, sortable: true, dataIndex: 'genere', filter:{ }},
            {id:'specie', header: "Specie", width: 150, sortable: true, renderer: spe, dataIndex: 'specie', filter:{ }},
            {id:'autore', header: "Autore", width: 150, sortable: true, renderer: all, dataIndex: 'autore', filter:{ }},
            {id:'comme', header: "Com.", width: 53, sortable: false, renderer: comme, dataIndex: 'comme',              
              filter: 
              {
                field: 
                {
                  xtype: "combo",
                  mode: 'local',
                  //store: new Ext.data.ArrayStore
                  store: new Ext.data.JsonStore
                  ({
                    id: 0,
                    /*fields : ['value'],
                    data: [['-'], ['PO'], ['PC'], ['PCO'],['PV'], ['PVO'], ['P']]*/                    
                    fields : ['name', 'value'], 
                    data: 
                    [
                        {name : '-',   value: '-'},
                        {name : 'PO = Piante officinali - Aromatiche e cosmetiche', value: 'PO'},
                        {name : 'PC = Piante commestibili', value: 'PC'},
                        {name : 'PCO = Piante commestibili e officinali', value: 'PCO'},
                        {name : 'PV = Piante velenose', value: 'PV'},
                        {name : 'PVO = Piante velenose e officinali', value: 'PVO'},
                        {name : 'P = Pianta senza indicazioni di propriet&agrave', value: 'P'}
                    ]
                  }),
                  valueField: 'value',
                  displayField: 'name',
                  listWidth: 250,
                  //displayField: 'value',
                  triggerAction: 'all',
                  value: "-"
                },
                fieldEvents: ["select"],
                test: function(filterValue, value) 
                      {
                        return filterValue === "-" || filterValue === value;
                      }
               }                        
            },
            {id:'famiglia', header: "Famiglia", width: 150, sortable: true, renderer: fam, dataIndex: 'famiglia', filter:{ }},
            {id:'nomecomune', header: "Nome Comune", width: 190, sortable: true, renderer: all, dataIndex: 'nomecomune', filter:{ }},
            {id:'datains', header: "Aggior.", width: 75, sortable: true, renderer: data, dataIndex: 'datains'}   
        ],
        highlightClasses: {
          ASC:  'x-custom-sort-asc'
          // DESC stays at default = x-ux-col-sort-desc
        },        
        bbar: new Ext.PagingToolbar({
            pageSize: 15,
            store: store,
            displayInfo: true,
            plugins: [new Ext.ux.ProgressBarPager()]
        }),    
        plugins: [filterRow, panelResizer, highlightSort, cellTips],
        
        tbar: tbar,
        /*footer: true,
        
        footerCfg: {
            cls: 'testo_foot',            
            html: '<br /><center>a cura di Elia Curti WebMaster A.M.I.N.T. - Grafica A.M.I.N.T. Tomaso Lezzi</center>'
        },*/         
        height:700,
        width:1010,
        frame:true,                
        title:
        '<div align="center"><img src="images/testata_micologia.jpg" alt="" height="68" width="974" border="0"></div>',
        renderTo: "grid-example"
      });
      
    grid.render('grid-example');    

    store.load({params:{start:0, limit:15}});      
           
        
    function all(val){
            return '<span style="color:#000088;">' + val + '</span>';
    }
    
    function gen(val,meta,record){                   
            var topic = record.data.topic;
            var sino = record.data.scheda;
            if (sino == 'S'){
            return ('<a class="link20" style="color:#000088;" href="http://www.funghiitaliani.it/index.php?showtopic='+topic+'" target="_blank">' + val + '</a>');
            }
            return ('<a class="link20" href="http://www.funghiitaliani.it/index.php?showtopic='+topic+'" target="_blank">' + val + '</a>');
    }
    
    function spe(val,meta,record){
            var sino = record.data.scheda;
            if (sino == 'S'){
            return '<i><span style="color:#000088;">' + val + '</span></i>';
            }             
            return '<i><span style="color:#007883;">' + val + '</span></i>';
    }
    
    function fam(val){
            return '<i><span style="color:#000088;">' + val + '</span></i>';
    }
    
    function data(val){
        anno=val.substring(0,4);
        mese=val.substring(4,6);
        giorno=val.substring(6,8);
        date=giorno+"/"+mese+"/"+anno;
        if(val == '0'){            
            return '<span style="color:#000088;">' + "" + '</span>';
        }
        return '<span style="color:#000088;">' + date + '</span>';        
    }
    
    function scheda(val){
      if (val=='A')
        {
          return '<img align="center" src="images/scheda1.gif" alt="" height="21" width="25" border="0"/>';
        }
      else if (val=='B')
        {
          return '<img src="images/scheda2.gif" alt="" height="21" width="25" border="0"/>';
        }
      return '<img src="images/scheda3.gif" alt="" height="21" width="25" border="0"/>';       
    }      
    
    function comme(val){
      switch (val)
      {
        case 'P':
        return '<img align="center" src="images/P.gif" border="0"/>';
        break;
        case 'PC':
        return '<img align="center" src="images/PC.gif" border="0"/>';
        break;
        case 'PV':
        return '<img align="center" src="images/PV.gif" border="0"/>';
        break;
        case 'PO':
        return '<img align="center" src="images/PO.gif" border="0"/>';
        break;
        case 'PCO':
        return '<img align="center" src="images/PCO.gif" border="0"/>';
        break;
        case 'PVO':
        return '<img align="center" src="images/PVO.gif" border="0"/>';
        break;                                        
      }      
    }        
        
});