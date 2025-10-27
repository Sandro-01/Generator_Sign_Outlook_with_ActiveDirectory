"""
File di configurazione di esempio per Generator_Sign_Outlook_with_ActiveDirectory
Copia questo file in config.py e modifica con i tuoi valori
NON committare config.py su GitHub!
"""

# Configurazione Active Directory
AD_SERVER = "dc.tuaazienda.local"  # Modifica con il tuo server AD
DOMAIN = "TUODOMINIO"              # Modifica con il tuo dominio

# Configurazione sedi (modifica con le tue OU)
SEDI = {
    '1': {
        'nome': 'Sede Principale',
        'ou': 'OU=Principale,OU=Client,DC=tuaazienda,DC=local'
    },
    '2': {
        'nome': 'Sede Secondaria',
        'ou': 'OU=Secondaria,OU=Client,DC=tuaazienda,DC=local'
    },
    # Aggiungi altre sedi...
}

# Informazioni aziendali per la firma
COMPANY_INFO = {
    'nome': 'La Tua Azienda S.r.l.',
    'indirizzo': 'Via Roma 123, 20100 Milano (MI)',
    'website': 'www.tuaazienda.it',
    'logo_url': 'https://www.tuaazienda.it/logo.png',
    'colore_primario': '#0066cc',
    'colore_secondario': '#666666',
    'disclaimer': 'Questo messaggio è confidenziale. Se non siete il destinatario, vi preghiamo di eliminarlo.'
}
