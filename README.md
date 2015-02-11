## **ABCLogico**

Primi passi di programmazione per
bambini 

[Installazione]()

#### **Presentazione**

 

Progetto propedeutico per
l'apprendimento del programmazione attraverso analisi di un problema logica e
la realizzazione della soluzione.

 

_ABCLogica_ nasce come mini progetto educativo
finalizzato principalmente all'apprendimento del processo di programmazione ed
il funzionamento basale del microprocessore.

 

L'obiettivo di ABCLogica e di creare un semplice 

[ambiente
di sviluppo](http://it.wikipedia.org/wiki/Integrated_development_environment) (un visual studio) che
permetteal principiantedi
sperimentare i principi del design di software:

1. Analisi logica del problema
2. Implementazione di un programma che risolve il
     problema
3. Verifica tramite esecuzione del programma

 

Il studente ha al suo
disposizione un semplice set di istruzioni minimali ([reduced
instruction set](http://it.wikipedia.org/wiki/Reduced_instruction_set_computer))

      
I.    
Movimentazione dati (Rappresentato da un gatto)

    
II.    
Test di condizioni (presenza cibo, cane)

  
III.    
Esecuzioni o salto condizionale (if-then, for-next loop)

#### **A chi e diretto**

"ABC Logico - il gatto di
turing" è rivolto a giovani di 6-10 anni di età e ai loro genitori. Si
propone come obiettivo di far apprendere i meccanismi e limiti dei onnipresenti
microprocessori per la generazione che cresce nell’era informatica. Tutti sanno
utilizzare dei App e videogiochi ma forse pochi comprendono che si basano su
semplice algoritmi che da solo apparenza di [essere
processi intelligenti o magici](http://www.repubblica.it/scienze/2014/06/10/news/computer_si_finge_umano-88594405/). 

Il processo di programmazione
rappresenta invece una sfida all’intelletto ed alla creatività con delle misure
ben precise sul raggiungimento del obiettivo. Quando si aggiunge variazioni nel
problema da risolvere bisogna sviluppare la capacità di pensare astratto per
implementare algoritmi che si adattano per risolvere diversi problemi.

 

#### **Come è fatto?**

L’interfaccia del ABCLogico è
composto di 3 parte

- Un editor di codice sorgente 
    - Un elenco di istruzioni a disposizione secondo il
      livello del utente
    - il programma creato (compilatore)

- Un campo che dimostra un scenario del problema da
     risolvere e l’andamento del programma
- Una finestra di feedback con esiti del programma e
     suggerimenti
 

Lo studente deve programmare il
gatto di muoversi in maniera di arrivare a casa. Ci varie scenari richiedono
ovviamente diversi percorsi. Ci sono dei ostacoli da evitare e compiti da
eseguire. I istruzioni devono essere organizzate in modo che il gatto mangia il
cibo ed evita i cani sul percorso verso casa. Dal programmazione di un percorso
specifico il studente dovrà cominciare a generare pezzi generici utilizzando
comandi condizionati dai test, tipo se c’è cibo mangia, altrimenti
prosegui.  

Quindi in essenza e
l’implementazione della [macchina di
Turing](http://it.wikipedia.org/wiki/Macchina_di_Turing_universale)

Il gatto – rappresenta dello
stato attuale del programma – in un campo – [spazio di memoria](http://it.wikipedia.org/wiki/Architettura_di_von_Neumann) von Neumann – in cui dati
possono essere modificati.
#### **Contributo del adulto o più esperto**

Il percorso prevede una graduale
aumento del livello di complessità e apprendimento in cui sarebbe opportuno
guidare il/la bambino/a leggendo l’istruzioni e cambiare impostazioni insieme.
Innanzitutto bisogna [scaricare
ABCLogico](https://github.com/ThorsenRune/ABCLogico/blob/master/ABCLogico.zip), un piccolo programma che non richiede installazione. 

Al primo livello lo studente deve
semplicemente calcolare il percorso verso casa - deve contare i passi da fare e
dove girare. 

Il programma va creato tirando
appositi istruzioni dalla finestra ‘commandi’ al finestra di programma a
fianco, usando il mouse. Istruzioni possono essere rimossi dal programma
utilizzando il cestino.

A
qualsiasi momento si possa provare il programma avviandolo a singoli passi
(step) o in esecuzione automatica (run).
Quando si ha acquisito
familiarità con l’ambiente si potrebbe salvare il programma e problema per poi
scegliere un problema più difficile. 

 

Successivo passo è di aggiungere
la possibilità di ripetere comandi con istruzione [ripeti]. Adesso l’obiettivo
sarebbe di rendere più efficace la programmazione minimizzare i numeri di righe
di codice (quindi ottimizzare l’uso della memoria). 

Quando il concetto di ripetizione
del codice e chiaro si passa al apprezzamento del test logici, ovvero
diramazione tra istruzioni da eseguire o saltare. A questo livello viene
aggiunto istruzioni di test tipo se c’è un cane/cibo/ostacolo avanti, quindi fa
la prossima istruzione (tipo cambia direzione o mangia) oppure saltala. Ora il
programmatore ‘in spe’ comincia a sviluppare codice più intelligente potrebbe
salire un altra livello con aggiunto di possibilità di raggruppare sezioni di
istruzioni in chiamate  (il principio di
__subroutine__,  __funzioni – functions o metodi)__. A questo punto il programmatore potrebbe
passare ad sviluppare un programma generico che porta a casa il gatto in
scenari diverse ponendo come l’obiettivo il minimo dispensabile di codice. 

 

Se poi vuole approfondire l’arte
di programmazione l’alluno potrebbe passare a programmi più avanzati come il  [PythonTurtle](http://pythonturtle.org/), ****[Scratch](http://scratch.mit.edu/),
oppure lanciarsi al programmazione in ****[C++,
JAVA, JavaScript o scaricare il grande Visual Studio](http://www.visualstudio.com/it-it/visual-studio-homepage-vs.aspx). C’è solo da studiare.

 

Se invece vuole programmare il
Visual Basic, che permette un facilita di programmazione seria, potrebbe
cominciare modificare il [codice
sorgente](https://github.com/ThorsenRune/ABCLogico/archive/master.zip) del ABCLogico stesso, per contribuire al progetto.
#### **Installazione**

ABCLogico è testato __su Windows__ XP e Windows 7. Non richiede installazione, basta
[scaricare
ABCLogico.zip](https://github.com/ThorsenRune/ABCLogico/blob/master/ABCLogico.zip).
Unzip e avvia ABCLogico.bat.

Note su Win7: 

Problemi
di avvio? - potrebbe essere necessario di avere dei privilegi di amministratore
per avviare programmi. Vai alla cartella “BIN” , __cliccando destro su __ABCLogico.exe scelgendo 'avvio come amministratore'/'run as
administrator'
