# Generator Sign Outlook with Active Directory

Script Python per generare automaticamente firme email per Outlook utilizzando i dati provenienti da Active Directory.

## Descrizione

Questo progetto permette di:
- Connettersi ad Active Directory tramite LDAP
- Recuperare informazioni degli utenti (nome, cognome, email, telefono, ecc.)
- Generare firme HTML personalizzate per Outlook
- Salvare le firme nella cartella appropriata di Outlook

## Requisiti

- Python 3.x
- Accesso ad Active Directory
- Microsoft Outlook installato

## Installazione

1. Clona il repository:
```bash
git clone https://github.com/Sandro-01/Generator_Sign_Outlook_with_ActiveDirectory.git
cd Generator_Sign_Outlook_with_ActiveDirectory
```

2. Installa le dipendenze:
```bash
pip install -r requirements.txt
```

3. Configura i parametri di connessione ad Active Directory nel file di configurazione

## Configurazione

Prima di eseguire lo script, configura:
- Server LDAP
- Credenziali di accesso
- Base DN per la ricerca utenti
- Template della firma

## Utilizzo

```bash
python Generator_Sign_Outlook_with_ActiveDirectory.py
```

## Sicurezza

⚠️ **IMPORTANTE**: Non committare mai credenziali o informazioni sensibili nel repository.

- Usa file di configurazione separati (`.env` o `config.ini`)
- Aggiungi file sensibili al `.gitignore`
- Considera l'uso di variabili d'ambiente per credenziali

## Struttura del Progetto

```
.
├── Generator_Sign_Outlook_with_ActiveDirectory.py  # Script principale
├── requirements.txt                                 # Dipendenze Python
├── .gitignore                                       # File da escludere
└── README.md                                        # Documentazione
```

## Contribuire

Le pull request sono benvenute. Per modifiche importanti, apri prima un issue per discutere cosa vorresti cambiare.

## Licenza

[Specificare la licenza]

## Contatti

Sandro-01 - GitHub
