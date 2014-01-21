/*
Text[...] = [titolo,testo]

Style[...] = [TitleColor,TextColor,TitleBgColor,TextBgColor,TitleBgImag,TextBgImag,TitleTextAlign,TextTextAlign, TitleFontFace, TextFontFace, TipPosition, StickyStyle, TitleFontSize, TextFontSize, Width, Height, BorderSize, PadTextArea, CoordinateX , CoordinateY, TransitionNumber, TransitionDuration, TransparencyLevel ,ShadowType, ShadowColor]
*/

var FiltersEnabled = 1; // se non vuoi usare transizioni o filtri in nessuno dei tips impostala a 0

Text[0]=["Informazioni","Le parole che nel testo sono sottolineate aprono una maschera del Glossario al passaggio del mouse"];
Text[1]=["Carl von Linné","Carl von Linné (1707-1778), noto anche come Carolus Linnaeus o, in italiano Carlo Linneo, fu un biologo e grande sistematico del Settecento; fu il creatore della moderna classificazione scientifica"];
Text[2]=["Sistema naturae","In sistema naturae, Pubblicato nel 1758, Linneo suddivide gli esseri viventi in due grandi raggruppamenti: il regno delle piante e quello degli animali e propone l’utilizzazione di un sistema binomiale di nomenclatura (Genere e specie)"];
Text[3]=["I.C.B.N.","Il Codice di Nomenclatura Botanica (I.C.B.N.) è costituito da un insieme di articoli che regolano come debbano venire dati i nomi dei vegetali e dei funghi.<br>Il codice sancisce quali sono i requisiti per i quali un nome sia validamente pubblicato, cosa è un tipo nomenclaturale e come debba venire designato, ecc.."];
Text[4]=[" "," "];


Style[0]=["white","black","#000099","white","","","","","","","","","","",400,"",2,1,10,10,51,1,0,"",""];
Style[1]=["white","black","#000099","#E8E8FF","","","","","","","center","","","",200,"",2,2,10,10,"","","","",""];
Style[2]=["white","black","#000099","#E8E8FF","","","","","","","left","","","",200,"",2,2,10,10,"","","","",""];
Style[3]=["white","black","#000099","#E8E8FF","","","","","","","float","","","",200,"",2,2,10,10,"","","","",""];
Style[4]=["white","black","#000099","#E8E8FF","","","","","","","fixed","","","",200,"",2,2,1,1,"","","","",""];
Style[5]=["white","black","#000099","#E8E8FF","","","","","","","","sticky","","",200,"",2,2,10,10,"","","","",""];
Style[6]=["white","black","#000099","#E8E8FF","","","","","","","","keep","","",200,"",2,2,10,10,"","","","",""];
Style[7]=["white","black","#000099","#E8E8FF","","","","","","","","","","",200,"",2,2,40,10,"","","","",""];
Style[8]=["white","black","#000099","#E8E8FF","","","","","","","","","","",200,"",2,2,10,50,"","","","",""];
Style[9]=["white","black","#000099","#E8E8FF","","","","","","","","","","",200,"",2,2,10,10,51,0.5,75,"simple","gray"];
Style[10]=["white","black","black","white","","","right","","Impact","cursive","center","",3,5,200,150,5,20,10,0,50,1,80,"complex","gray"];
Style[11]=["white","black","#000099","#E8E8FF","","","","","","","","","","",200,"",2,2,10,10,51,0.5,45,"simple","gray"];
Style[12]=["white","black","#000099","#E8E8FF","","","","","","","","","","",200,"",2,2,10,10,"","","","",""];

applyCssFilter();