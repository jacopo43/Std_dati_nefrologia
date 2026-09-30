# CONVERSIONE E STANDARDIZZAZIONE DATI NEFROLOGIA

## Indice

1. [Obiettivo](#1-obiettivo)
2. [Scelta strutturale](#2-scelta-strutturale)
3. [File coinvolti](#3-file-coinvolti)
4. [Requisiti e installazione](#4-requisiti-e-installazione)
   - [4.1 Pacchetti necessari](#41-pacchetti-necessari)
   - [4.2 Installazione](#42-installazione)
5. [Esecuzione](#5-esecuzione)
   - [5.1 Esecuzione base](#51-esecuzione-base)
   - [5.2 Nomi personalizzati](#52-nomi-personalizzati)
6. [Logica delle variabili di output](#6-logica-delle-variabili-di-output)
7. [Pulizia generale](#7-pulizia-generale)
8. [Output](#8-output)
9. [Costruzione dei dizionari](#9-costruzione-dei-dizionari)
   - [9.1 Dizionario dei farmaci](#91-dizionario-dei-farmaci)
   - [9.2 Dizionario delle comorbidità](#92-dizionario-delle-comorbidità)
   - [9.3 Dizionario delle conseguenze delle interazioni](#93-dizionario-delle-conseguenze-delle-interazioni)

## 1. Obiettivo

Il programma trasforma il file originale, nel quale ogni paziente occupa un foglio separato, in due file Excel: 
un output pulito con il foglio "dati_standardizzati" e un output di controllo con il foglio "dati_con_fonti". 
I dati anagrafici identificativi non vengono riportati. Ogni paziente riceve un identificativo progressivo anonimo (id_paziente).

## 2. Scelta strutturale

I domini presenti nel file originale hanno cardinalità diverse: un paziente può avere una sola età, molte comorbidità, molti farmaci e molte interazioni. 
Unirli orizzontalmente produrrebbe duplicazioni o prodotti cartesiani. Per evitare questo problema l'output è in formato lungo e contiene quattro tipi di record:
- riepilogo: una riga sempre presente per ciascun paziente;
- comorbidita: una riga per ciascuna comorbidità standardizzata;
- terapia: una riga per ciascuna terapia pre revisione farmacologica;
- interazione: una riga per ciascuna interazione farmacologica.

## 3. File coinvolti

- converti_dati_nefrologia.py: programma principale.
- dizionario_farmaci.xlsx: associa sinonimi e denominazioni commerciali al principio attivo e al codice ATC.
- dizionario_comorbidita.xlsx: categorie standard, sinonimi, radici di ricerca e regole di uso automatico.
- dizionario_interazioni.xlsx: categorie standard delle conseguenze delle interazioni e termini associati.
- mappa_variabili.xlsx: documentazione sintetica del passaggio sorgente -> trasformazione -> variabile finale.
- costruzione_dizionari.txt: descrizione dettagliata della costruzione e delle fonti dei dizionari.
- requisiti.txt: librerie Python necessarie.

## 4. Requisiti e installazione

### 4.1 Pacchetti necessari

```text
pandas>=2.2
numpy>=1.26
openpyxl>=3.1
xlrd>=2.0.1
```

### 4.2 Installazione

Da terminale, nella cartella del pacchetto:

```bash
pip install -r requisiti.txt
```

Per leggere il file originale .xls è necessario xlrd, già incluso in
requisiti.txt.

## 5. Esecuzione

### 5.1 Esecuzione base

```bash
python converti_dati_nefrologia.py pz_nefro_originale.xls
```

Vengono creati automaticamente:
- dati_nefrologia_standardizzati.xlsx
- dati_nefrologia_standardizzati_controllo_fonti.xlsx

### 5.2 Nomi personalizzati

```bash
python converti_dati_nefrologia.py pz_nefro_originale.xls \
    --output dati_puliti.xlsx \
    --output-controllo dati_controllo_fonti.xlsx
```

Se i dizionari sono in una cartella diversa:

```bash
python converti_dati_nefrologia.py pz_nefro_originale.xls \
    --risorse "C:/percorso/dizionari"
```

## 6. Logica delle variabili di output

### `id_paziente`
**Origine:** ordine dei fogli-paziente nel file sorgente, escluso il foglio RIASSUNTO.
**Trasformazione:** numerazione progressiva 1, 2, 3, ...
**Motivo:** elimina la necessità di esportare codice fiscale, nome o cognome.

### `tipo_record`
**Origine:** non presente nel file originale.
**Trasformazione:** valore derivato dal dominio del record: riepilogo, comorbidita, terapia o interazione.
**Motivo:** consente di mantenere in un unico foglio informazioni con cardinalità differenti senza creare combinazioni artificiali.

### `indice_record`
**Origine:** ordine del record all'interno del singolo dominio e del singolo paziente.
**Trasformazione:** numerazione progressiva a partire da 1.

### `eta_anni`
**Origine:** DDN.
**Data di riferimento:** DATA VALUTAZIONE; se assente, DATA VISITA AMB.
**Trasformazione:** differenza esatta in anni compiuti. Se una delle date non è disponibile il valore resta mancante.
**Miglioramento rispetto al notebook:** non viene più usata una data fissa arbitraria, quindi l'età è riferita al momento clinicamente pertinente.

### `fumo`
**Origine:** MEDICAL HISTORY.
**Dizionario:** dizionario_comorbidita.xlsx.
**Trasformazione:** 1 se il testo produce la categoria FUMO, 0 altrimenti. Le negazioni vengono escluse.

### `alcool`
**Origine:** MEDICAL HISTORY.
**Dizionario:** dizionario_comorbidita.xlsx.
**Trasformazione:** 1 se il testo produce la categoria ALCOL/ALCOOL, 0 altrimenti. Le negazioni vengono escluse.

### `data_raccolta_terapia`
**Origine:** DATA VISITA AMB.
**Trasformazione:** data standard ISO AAAA-MM-GG.

### `data_revisione_farmacologica`
**Origine:** DATA VALUTAZIONE.
**Trasformazione:** data standard ISO AAAA-MM-GG.

### `punteggio_acb_pre`
**Origine:** ACB Score.
**Trasformazione:** conversione numerica; valori non interpretabili restano mancanti.

### `numero_comorbidita_pre`
**Origine:** tutte le righe di MEDICAL HISTORY.
**Dizionario:** dizionario_comorbidita.xlsx.
**Trasformazione:** conteggio delle categorie standard uniche, escludendo FUMO e ALCOL/ALCOOL perché trattate come variabili dedicate.

### `numero_farmaci_pre`
**Origine:** TERAPIA IN CORSO DI RICOVERO.
**Trasformazione:** numero di righe non vuote di terapia. La misura rappresenta quindi le terapie registrate, indipendentemente dal successo del riconoscimento del 
principio attivo.

### `numero_interazioni_cd_pre`
**Origine:** GRAVITA' INTERAZIONI.
**Trasformazione:** conteggio delle righe con classe C o D.

### `numero_farmaci_inappropriati_beers_pre`
**Origine:** CRITERI DI BEERS.
**Dizionario:** dizionario_farmaci.xlsx.
**Trasformazione:** conteggio dei principi attivi unici riconosciuti.

### `numero_farmaci_inappropriati_start_pre`
**Origine:** CRITERI START.
**Dizionario:** dizionario_farmaci.xlsx.
**Trasformazione:** conteggio dei principi attivi unici riconosciuti.

### `numero_farmaci_inappropriati_stopp_pre`
**Origine:** CRITERI STOPP.
**Dizionario:** dizionario_farmaci.xlsx.
**Trasformazione:** conteggio dei principi attivi unici riconosciuti.

### `comorbidita_pre`
**Origine:** MEDICAL HISTORY.
**Dizionario:** dizionario_comorbidita.xlsx.
**Trasformazione:** normalizzazione del testo, ricerca di sinonimi/radici, esclusione di negazioni, familiarità e formulazioni incerte, rimozione dei duplicati. 
Le infezioni vengono ricondotte alla categoria generale prevista dal dizionario. Le categorie specifiche prevalgono su quelle generiche quando previsto.

### `testo_comorbidita_originale`
**Origine:** MEDICAL HISTORY.
**Trasformazione:** nessuna trasformazione sostanziale; è conservato il testo che ha generato la categoria standard. Serve per controllo e validazione manuale.

### `data_ricognizione`
**Origine:** DATA VALUTAZIONE.
**Trasformazione:** data ISO AAAA-MM-GG e ripetizione sulle righe di terapia del paziente.

### `farmaco_pre`
**Origine:** TERAPIA IN CORSO DI RICOVERO.
**Dizionario:** dizionario_farmaci.xlsx.
**Trasformazione:** ricerca del principio attivo tramite nome canonico, sinonimi, denominazioni commerciali e radici di ricerca. Se non riconosciuto resta mancante; 
il testo originale è comunque conservato.

### `codice_atc_pre`
**Origine:** derivato dal farmaco riconosciuto.
**Dizionario:** dizionario_farmaci.xlsx.
**Trasformazione:** codice ATC associato al principio attivo; può essere mancante se non disponibile nel dizionario.

### `dosaggio_pre`
**Origine:** TERAPIA IN CORSO DI RICOVERO.
**Trasformazione:** prima quantità riconosciuta prima di un'unità farmacologica (kg, mg, mcg/ug/µg, g, UI/IU); la virgola decimale viene convertita in punto numerico.

### `unita_dosaggio_pre`
**Origine:** TERAPIA IN CORSO DI RICOVERO.
**Trasformazione:** unità riconosciuta e normalizzata in minuscolo.

### `somministrazioni_giornaliere_pre`
**Origine:** TERAPIA IN CORSO DI RICOVERO.
**Trasformazione:** stima della frequenza giornaliera da espressioni come "ore 8 e ore 20", "2 volte al giorno", "ogni 8 ore", "bid", "tid", "1 cp". Non rappresenta 
i milligrammi totali al giorno: rappresenta il numero di somministrazioni giornaliere.

### `via_somministrazione_pre`
**Origine:** TERAPIA IN CORSO DI RICOVERO.
**Trasformazione:** riconoscimento di espressioni standard come orale/per os, EV/IV, IM, SC/sottocute, inalatoria, topica, transdermica.

### `testo_terapia_originale`
**Origine:** TERAPIA IN CORSO DI RICOVERO.
**Trasformazione:** conservato integralmente per permettere verifica dei campi derivati.

### `farmaco_interagente_1`
**Origine:** FARMACO interagente "A".
**Trasformazione:** forward-fill per gestire celle unite nel file originale; il valore non viene forzato nel dizionario per non perdere la denominazione clinicamente 
inserita.

### `farmaco_interagente_2`
**Origine:** FARMACO interagente "B".
**Trasformazione:** forward-fill per gestire celle unite nel file originale.

### `tipo_interazione`
**Origine:** GRAVITA' INTERAZIONI.
**Trasformazione:** mantenimento del valore sorgente (per esempio C o D).

### `motivo_interazione`
**Origine:** POSSIBILI EFFETTI INTERAZIONE.
**Trasformazione:** testo originale conservato; costituisce anche l'input per la codifica delle conseguenze.

### `conseguenza_interazione_1 / 2 / 3`
**Origine:** POSSIBILI EFFETTI INTERAZIONE.
**Dizionario:** dizionario_interazioni.xlsx.
**Trasformazione:** ricerca multi-etichetta delle conseguenze standard; al massimo tre categorie ordinate per specificità/priorità. Una tossicità specifica prevale 
sulle categorie generiche; il rischio di sanguinamento viene distinto dal sanguinamento già manifestato.

### `consiglio_clinico`
**Origine:** COMPORTAMENTO CLINICO.
**Trasformazione:** testo originale, senza codifica automatica.

## 7. Pulizia generale

- gli spazi iniziali/finali e i caratteri nulli vengono rimossi durante il matching;
- il matching non dipende da maiuscole/minuscole o accenti;
- i dati mancanti sono ammessi e non vengono imputati;
- le righe vuote non generano record;
- le righe di interazione create esclusivamente dal forward-fill vengono eliminate se per la stessa coppia esiste già una riga con dettagli;
- i duplicati identici delle interazioni vengono rimossi;
- nome, cognome e codice fiscale non vengono esportati.

## 8. Output

Lo script genera ora due file:
1. il file standardizzato pulito, con le sole variabili finali;
2. un file di controllo, in cui ogni variabile è immediatamente affiancata da
   "fonte_<nome_variabile>". La fonte riporta la colonna e il valore originale
   da cui la variabile è stata ricavata e, quando il valore dipende da un
   dizionario, indica anche il dizionario utilizzato.

Il file di controllo è progettato per validare il comportamento dello script e
confrontare rapidamente il valore standardizzato con il testo sorgente.

## 9. Costruzione dei dizionari

### 9.1 Dizionario dei farmaci

#### 9.1.1 Fonti originarie
Il dizionario dei farmaci è stato costruito a partire principalmente da:

1. classificazione ATC (Anatomical Therapeutic Chemical Classification System);
2. dizionario dei farmaci utilizzato nel progetto DiAna, fornito come input
   originario;
3. denominazioni e varianti di farmaci presenti nei dati clinici forniti;
4. eventuali nomi commerciali, sinonimi e varianti ortografiche presenti negli
   input originari.

La classificazione ATC costituisce la fonte per l'identificazione del principio
attivo e, quando disponibile, per il relativo codice e livello/classificazione
farmacologica.

Il dizionario DiAna e le denominazioni presenti nei dati originari ampliano la
possibilità di riconoscere lo stesso principio attivo quando viene scritto con
forme diverse.

#### 9.1.2 Costruzione del nome standard
Per ogni farmaco viene definita una denominazione standard, corrispondente al
principio attivo o alla combinazione di principi attivi che deve essere
riportata nell'output.

Quando negli input originari sono presenti più denominazioni riferibili allo
stesso farmaco, queste vengono ricondotte alla stessa denominazione standard.

Esempio concettuale:

    denominazione standard
        <- nome del principio attivo ATC
        <- sinonimi
        <- denominazioni alternative
        <- nomi commerciali presenti negli input
        <- varianti ortografiche documentate

Le diverse denominazioni non diventano quindi farmaci differenti, ma termini
alternativi di riconoscimento dello stesso farmaco.

#### 9.1.3 Generazione dei termini multipli
I termini alternativi possono derivare da:

- denominazione del principio attivo;
- denominazioni alternative presenti nel dizionario farmacologico originario;
- nome commerciale quando disponibile nell'input;
- forme con o senza sali o specificazioni farmaceutiche, quando già presenti
  nelle fonti originarie;
- varianti ortografiche effettivamente osservate negli input;
- combinazioni di principi attivi già definite come tali nelle fonti.

Non vengono inventati nuovi principi attivi.

Quando una fonte contiene più termini equivalenti, questi vengono separati e
memorizzati come sinonimi dello stesso elemento standard.

#### 9.1.4 Radici di ricerca
Per alcuni termini può essere utilizzata una radice esclusivamente per
riconoscere flessioni o varianti ortografiche dello stesso termine.

La radice non rappresenta una nuova fonte terminologica.

Nel matching automatico una radice deve iniziare all'inizio di una parola e
non può essere cercata come sequenza arbitraria di caratteri all'interno di
un'altra parola.

Questa regola evita, per esempio, che una sequenza come "asm" venga
riconosciuta all'interno di "spasmo".

#### 9.1.5 Normalizzazione tecnica
Prima del confronto vengono uniformati:

- maiuscole/minuscole;
- accenti;
- apostrofi;
- punteggiatura non informativa;
- spazi multipli.

Queste trasformazioni non generano sinonimi clinici nuovi: rendono soltanto
comparabili forme graficamente diverse dello stesso termine.


### 9.2 Dizionario delle comorbidità

#### 9.2.1 Fonti originarie
Il dizionario delle comorbidità è stato costruito a partire da:

1. elenco delle comorbidità e delle categorie cliniche definite negli input
   originari del progetto;
2. terminologia MedDRA fornita come input originario, utilizzata come
   riferimento terminologico per diagnosi, condizioni ed eventi clinici;
3. termini clinici e varianti presenti nei dati di cartella clinica forniti;
4. esempi, eccezioni e indicazioni cliniche fornite durante la definizione
   delle variabili;
5. acronimi e denominazioni mediche equivalenti presenti negli input originali.

I file tecnici successivamente creati per organizzare queste informazioni non
sono fonti indipendenti: sono la traduzione operativa di questi input.

#### 9.2.2 Costruzione delle categorie standard
Le categorie finali non corrispondono necessariamente a una singola stringa
testuale.

Per ogni concetto clinico di interesse è stata definita una categoria standard
da utilizzare nell'output. Alla categoria vengono associati i termini
originari che possono esprimere lo stesso concetto.

Schema generale:

    categoria standard
        <- termine clinico principale
        <- termine MedDRA pertinente, quando utilizzato negli input
        <- sinonimi clinici
        <- acronimi
        <- denominazioni italiane/inglesi presenti negli input
        <- varianti ortografiche osservate
        <- forme morfologiche riconoscibili tramite radice

#### 9.2.3 Generazione di sinonimi e varianti
I termini multipli associati a una categoria sono stati ottenuti combinando le
forme presenti nelle fonti originarie.

In particolare possono essere incluse:

- denominazione estesa;
- abbreviazione o acronimo;
- sinonimo clinico;
- termine italiano e corrispondente forma inglese, se presenti negli input;
- grafie alternative;
- varianti morfologiche dello stesso concetto.

La presenza di più termini serve esclusivamente a riconoscere modi diversi di
scrivere la stessa informazione clinica.

#### 9.2.4 Uso di MedDRA
MedDRA viene utilizzato come riferimento terminologico clinico quando i termini
originari provengono da tale nomenclatura o sono stati ricondotti a essa negli
input forniti.

Lo scopo non è trasformare automaticamente tutto il testo libero in un codice
MedDRA, ma utilizzare la terminologia disponibile per rendere più stabile e
riproducibile il riconoscimento dei concetti clinici di interesse.

Le categorie finali del progetto possono pertanto essere più aggregate rispetto
a un singolo Preferred Term MedDRA.

#### 9.2.5 Radici di ricerca
Per alcune comorbidità sono state utilizzate radici per riconoscere varianti
della stessa parola.

Esempio concettuale:

    ipertens-
        -> ipertensione
        -> ipertensivo
        -> ipertensiva

Una radice viene applicata soltanto dall'inizio di una parola.

Non è consentito il matching della radice nel mezzo di una parola.

Esempio:

    radice: asm

    "asma"       -> compatibile
    "asmatico"   -> compatibile
    "spasmo"     -> NON compatibile

Questa regola è fondamentale per ridurre i falsi positivi.

#### 9.2.6 Termini ambigui
Un termine breve, generico o potenzialmente ambiguo non viene necessariamente
utilizzato come match automatico.

Quando dagli input originari emerge che un termine può avere più significati,
esso può essere:

- escluso dal matching automatico;
- utilizzato soltanto con un vincolo di contesto;
- contrassegnato come termine da revisionare.

Questa distinzione deriva dalla valutazione clinica dei termini originali e non
da una generazione automatica di nuove diagnosi.

#### 9.2.7 Negazione, familiarità e incertezza
La presenza del termine non è sufficiente da sola per assegnare una
comorbidità.

Il contesto viene controllato per evitare di codificare come diagnosi presente
espressioni quali:

- "non presenta...";
- "assenza di...";
- "nessuna storia di...";
- "familiarità per...";
- "madre/padre con...";
- "sospetto...";
- "possibile...";
- "da escludere...".

Le espressioni di negazione, familiarità e incertezza sono state definite a
partire dalle regole e dagli esempi forniti negli input originari e
dall'osservazione del tipo di testo libero contenuto nelle cartelle.

#### 9.2.8 Accorpamento delle categorie
Quando più termini originari rappresentano varianti dello stesso concetto
richiesto nell'output, essi vengono ricondotti alla stessa categoria.

Esempi di questa logica possono comprendere:

- differenti descrizioni di una stessa patologia;
- sedi o forme cliniche che il protocollo richiede di aggregare;
- termini generici sostituiti da una categoria più specifica quando entrambe
  vengono riconosciute.

Queste regole sono definite sulla base della struttura clinica richiesta per il
dataset finale, non mediante una classificazione statistica automatica.


### 9.3 Dizionario delle conseguenze delle interazioni

#### 9.3.1 Fonti originarie
Il dizionario delle interazioni è stato costruito a partire da:

1. categorie di conseguenza delle interazioni farmacologiche definite negli
   input originari del progetto;
2. descrizioni testuali delle interazioni presenti nei dati clinici e nei
   materiali originari forniti;
3. terminologia clinica MedDRA utilizzata negli input quando la conseguenza
   dell'interazione corrisponde a un evento o a una condizione clinica;
4. esempi clinici e regole di interpretazione fornite per distinguere
   conseguenze specifiche da descrizioni farmacocinetiche o generiche;
5. denominazioni, sinonimi e varianti linguistiche presenti negli input
   originari.

Anche in questo caso i successivi file di dizionario o guida costituiscono
soltanto una formalizzazione tecnica di tali fonti.

#### 9.3.2 Costruzione delle categorie di conseguenza
Le frasi che descrivono una possibile conseguenza di un'interazione possono
essere molto diverse pur indicando lo stesso fenomeno.

Per questo motivo vengono ricondotte a una categoria standard.

#### 9.3.3 Schema generale:

    conseguenza standard
        <- descrizione clinica originaria
        <- termine clinico equivalente
        <- eventuale termine MedDRA pertinente
        <- sinonimo
        <- variante linguistica
        <- formulazione farmacologica equivalente


#### 9.3.4 Tipi di termini inclusi
Tra i termini associabili a una conseguenza possono rientrare:

- evento clinico manifesto;
- aumento del rischio di un evento;
- alterazione farmacocinetica;
- aumento o riduzione dell'esposizione;
- aumento della tossicità;
- riduzione dell'efficacia;
- alterazione di parametri clinici o laboratoristici;
- denominazioni equivalenti italiane e inglesi presenti negli input.

Un evento manifesto e un aumento del rischio dello stesso evento non vengono
necessariamente considerati equivalenti.

Per esempio, una formulazione che indica "rischio di sanguinamento" può essere
distinta da una formulazione che documenta un "sanguinamento" già avvenuto.

#### 9.3.5 Uso della terminologia MedDRA
Quando l'effetto dell'interazione è una manifestazione clinica, la terminologia
MedDRA fornita negli input può contribuire alla standardizzazione del concetto.

Le categorie finali, tuttavia, seguono il livello di aggregazione richiesto dal
progetto e non devono essere interpretate automaticamente come singoli codici
MedDRA.

#### 9.3.6 Priorità tra conseguenze
Quando una stessa frase permette di riconoscere sia una conseguenza generica
sia una conseguenza clinica specifica, viene privilegiata quella più
informativa.

Esempio concettuale:

    "aumento dei livelli del farmaco con rabdomiolisi"

può contenere contemporaneamente:

- aumento dell'esposizione;
- danno muscolare/rabdomiolisi.

La conseguenza clinica specifica viene preferita rispetto alla sola
descrizione farmacocinetica generica.

La regola deriva dalla necessità, definita negli input del progetto, di
conservare l'informazione clinicamente più specifica.

#### 9.3.7 Termini da revisionare
Formulazioni eccessivamente generiche o non sufficienti a identificare una
conseguenza univoca possono essere conservate come elementi da revisionare,
anziché essere convertite automaticamente.

Questo evita che la semplice presenza di parole come "tossicità", "aumento",
"riduzione" o altre espressioni generiche produca una classificazione non
supportata dal contesto.
