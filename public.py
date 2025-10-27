"""
Gestore Firme Email Active Directory -> Outlook
Permette di selezionare utenti specifici e aggiornare le loro firme Outlook
"""

import os
import sys
from pathlib import Path
from datetime import datetime

try:
    from ldap3 import Server, Connection, ALL, SUBTREE
except ImportError:
    print("ERRORE: Modulo ldap3 non trovato. Installa con: pip install ldap3")
    sys.exit(1)


class ADSignatureManager:
    def __init__(self, ad_server, domain, base_dn, username, password):
        """
        Inizializza la connessione ad Active Directory
        
        Args:
            ad_server: Server AD (es. 'dc.azienda.local')
            domain: Dominio (es. 'AZIENDA')
            base_dn: Base DN (es. 'DC=azienda,DC=local')
            username: Username AD con permessi di lettura
            password: Password
        """
        self.ad_server = ad_server
        self.domain = domain
        self.base_dn = base_dn
        self.username = username
        self.password = password
        self.connection = None
        
        # Configurazione aziendale
        self.company_info = {
            'nome': 'La Tua Azienda S.r.l.',
            'indirizzo': 'Via Roma 123, 20100 Milano (MI)',
            'website': 'www.tuaazienda.it',
            'logo_url': 'https://www.tuaazienda.it/logo.png',
            'colore_primario': '#0066cc',
            'colore_secondario': '#666666',
            'disclaimer': 'Questo messaggio è confidenziale. Se non siete il destinatario, vi preghiamo di eliminarlo.'
        }
    
    def connect_to_ad(self):
        """Connette ad Active Directory"""
        print(f"\n→ Tentativo connessione a: {self.ad_server}")
        print(f"  Dominio: {self.domain}")
        print(f"  Username: {self.username}")
        
        try:
            # Prova prima con LDAP standard (porta 389)
            server = Server(self.ad_server, port=389, get_info=ALL)
            
            # Prova diverse varianti del formato username
            user_formats = [
                f"{self.domain}\\{self.username}",  # DOMAIN\username
                f"{self.username}@{self.ad_server}",  # username@server
                f"{self.username}@{self.domain}",  # username@DOMAIN
                self.username  # username semplice
            ]
            
            connected = False
            for user_format in user_formats:
                try:
                    print(f"  Provo con formato: {user_format}")
                    self.connection = Connection(
                        server, 
                        user_format, 
                        self.password, 
                        auto_bind=True,
                        raise_exceptions=True
                    )
                    connected = True
                    print(f"✓ Connesso ad Active Directory: {self.ad_server}")
                    print(f"  Formato username utilizzato: {user_format}")
                    break
                except Exception as e:
                    print(f"  ✗ Formato {user_format} fallito: {str(e)[:100]}")
                    continue
            
            if not connected:
                # Prova LDAPS (porta 636)
                print("\n  Provo connessione sicura LDAPS (porta 636)...")
                server = Server(self.ad_server, port=636, use_ssl=True, get_info=ALL)
                user_dn = f"{self.domain}\\{self.username}"
                self.connection = Connection(server, user_dn, self.password, auto_bind=True)
                print(f"✓ Connesso via LDAPS")
                connected = True
            
            return connected
            
        except Exception as e:
            print(f"\n✗ Errore connessione AD: {e}")
            print("\n🔍 Suggerimenti:")
            print("  1. Verifica che il server AD sia raggiungibile (ping adds.cartongrp.com)")
            print("  2. Controlla username e password")
            print("  3. Verifica che il firewall permetta connessioni LDAP (porta 389/636)")
            print("  4. Prova a usare l'IP del server invece del nome DNS")
            print("  5. Assicurati che l'account abbia permessi di lettura su AD")
            return False
    
    def search_users(self, search_filter="(objectClass=user)"):
        """
        Cerca utenti in Active Directory
        
        Args:
            search_filter: Filtro LDAP (default: tutti gli utenti)
        
        Returns:
            Lista di dizionari con i dati degli utenti
        """
        if not self.connection:
            print("✗ Non connesso ad AD. Esegui connect_to_ad() prima.")
            return []
        
        attributes = [
            'sAMAccountName', 'displayName', 'givenName', 'sn',
            'title', 'mail', 'telephoneNumber', 'mobile',
            'department', 'company', 'physicalDeliveryOfficeName'
        ]
        
        print(f"\n→ Ricerca in corso su Base DN: {self.base_dn}")
        print(f"  Filtro: {search_filter}")
        
        try:
            result = self.connection.search(
                search_base=self.base_dn,
                search_filter=search_filter,
                search_scope=SUBTREE,
                attributes=attributes
            )
            
            if not result:
                print(f"✗ Ricerca fallita. Codice: {self.connection.result}")
                print(f"  Messaggio: {self.connection.result.get('description', 'Nessuna descrizione')}")
                
                # Prova con OU=Users se il Base DN non funziona
                alt_base_dn = f"OU=Users,{self.base_dn}"
                print(f"\n  Provo con Base DN alternativo: {alt_base_dn}")
                result = self.connection.search(
                    search_base=alt_base_dn,
                    search_filter=search_filter,
                    search_scope=SUBTREE,
                    attributes=attributes
                )
            
            print(f"  Elaborazione risultati... ({len(self.connection.entries)} entries trovate)")
            
            users = []
            for idx, entry in enumerate(self.connection.entries, 1):
                try:
                    # Gestione sicura degli attributi
                    def get_attr(entry, attr_name):
                        try:
                            if hasattr(entry, attr_name):
                                val = getattr(entry, attr_name)
                                if val and str(val) != '[]':
                                    return str(val).strip('[]')
                            return ''
                        except:
                            return ''
                    
                    user = {
                        'username': get_attr(entry, 'sAMAccountName'),
                        'display_name': get_attr(entry, 'displayName'),
                        'first_name': get_attr(entry, 'givenName'),
                        'last_name': get_attr(entry, 'sn'),
                        'title': get_attr(entry, 'title'),
                        'email': get_attr(entry, 'mail'),
                        'phone': get_attr(entry, 'telephoneNumber'),
                        'mobile': get_attr(entry, 'mobile'),
                        'department': get_attr(entry, 'department'),
                        'company': get_attr(entry, 'company') or self.company_info['nome'],
                        'office': get_attr(entry, 'physicalDeliveryOfficeName')
                    }
                    
                    # Aggiungi solo se ha email valida
                    if user['email'] and '@' in user['email']:
                        users.append(user)
                        if idx <= 3:  # Mostra primi 3 per debug
                            print(f"    [{idx}] {user['display_name']} - {user['email']}")
                    
                except Exception as e:
                    print(f"  ⚠ Errore elaborazione entry {idx}: {e}")
                    continue
            
            print(f"\n✓ Trovati {len(users)} utenti con email valida")
            
            if len(users) == 0:
                print("\n⚠ Nessun utente con email trovato. Possibili cause:")
                print("  1. Il Base DN non contiene utenti con email")
                print("  2. L'account non ha permessi per leggere gli attributi")
                print("  3. Gli utenti sono in una OU diversa")
                print(f"\n  Vuoi provare a cercare in tutto l'albero? (s/n): ", end='')
            
            return users
            
        except Exception as e:
            print(f"\n✗ Errore ricerca utenti: {e}")
            print(f"  Tipo errore: {type(e).__name__}")
            import traceback
            print("\n  Dettagli tecnici:")
            traceback.print_exc()
            return []
    
    def generate_signature_html(self, user):
        """
        Genera firma HTML per un utente
        
        Args:
            user: Dizionario con dati utente
        
        Returns:
            Stringa HTML della firma
        """
        mobile_html = ""
        if user['mobile']:
            mobile_html = f"Mobile: {user['mobile']}<br>"
        
        html = f"""<!DOCTYPE html>
<html>
<head>
    <meta charset="UTF-8">
    <meta name="Generator" content="AD-Outlook Signature Manager">
</head>
<body style="margin: 0; padding: 0; font-family: Arial, Helvetica, sans-serif;">
    <table cellpadding="0" cellspacing="0" border="0" style="font-family: Arial, Helvetica, sans-serif; font-size: 13px; line-height: 1.6; color: #333; max-width: 600px;">
        <tr>
            <td style="padding: 0; vertical-align: top;">
                <!-- Nome e Ruolo -->
                <div style="margin-bottom: 10px;">
                    <strong style="font-size: 15px; color: #000; font-weight: bold;">
                        {user['display_name'].replace('(..........)', '').replace('(Carton Group)', '').strip()}
                    </strong><br>
                    <span style="font-size: 13px; color: #333;">
                        {user['title']}
                    </span>
                </div>
                
                <!-- Informazioni Azienda -->
                <div style="margin-bottom: 10px; padding-bottom: 10px;">
                    <strong style="font-size: 13px; color: #000;">
                        Your Company
                    </strong><br>
                    <span style="font-size: 12px; color: #333;">
                        Address
                    </span>
                </div>
                
                <!-- Contatti -->
                <div style="margin-bottom: 15px; font-size: 12px; color: #333;">
                    Tel: {user['phone']}<br>
                    {mobile_html}Mail: <a href="mailto:{user['email']}" style="color: #0066cc; text-decoration: none;">{user['email']}</a>
                </div>
                
                <!-- Logo -->
                <div style="margin-bottom: 15px; max-width: 10px; height:10px; display:flex">
                    <img src="...." alt="Gruppo">
                </div>
                
                <!-- Loghi Vari
                <div style="margin-bottom: 15px;">
                    <table cellpadding="0" cellspacing="0" border="0">
                        <tr>
                            <td style="padding-right: 20px;">
                                <img src="https://via.placeholder.com/150x50/1a1a1a/ffffff?text=Carton+Group" alt="..." style="max-width: 150px; height: auto; display: block;">
                            </td>
                            <td>
                                <img src="https://via.placeholder.com/150x80/1a1a1a/ffffff?text=Beyond+Packaging" alt="..." style="max-width: 150px; height: auto; display: block;">
                            </td>
                        </tr>
                    </table>
                </div>
		-->
                
                <!-- Link Siti Web -->
                <div style="margin-bottom: 10px; font-size: 12px;">
                    <a href="https://www.google.it/" style="color: #0066cc; text-decoration: none;">www.google.it/</a> | 
                    <a href="https://www.google.com/" style="color: #0066cc; text-decoration: none;">www.google.com/</a>
                </div>
                
                <!-- Social Links -->
                <div style="margin-bottom: 15px; font-size: 12px;">
                    Follow us: 
                    <a href="https://www.linkedin.com/company/google" style="color: #0066cc; text-decoration: none;">google</a> | 
                    <a href="https://www.linkedin.com/company/google" style="color: #0066cc; text-decoration: none;">LinkedIn</a>
                </div>
                
                <!-- Disclaimer -->
                <div style="margin-top: 15px; padding-top: 10px; border-top: 1px solid #cccccc; font-size: 9px; color: #666; line-height: 1.3;">
                    This e-mail may contain confidential and/or privileged information. If you are not the intended recipient (or have received this e-mail in error) please notify the sender immediately and destroy this e-mail. Any unauthorized copying, disclosure or distribution of the material in this e-mail is strictly forbidden.
                </div>
            </td>
        </tr>
    </table>
</body>
</html>"""
        
        return html
    
    def generate_signature_txt(self, user):
        """Genera firma in formato testo"""
        mobile_txt = f"\nMobile: {user['mobile']}" if user['mobile'] else ""
        
        txt = f"""{user['display_name'].replace('(Europoligrafico)', '').replace('(Carton Group)', '').strip()}
{user['title']}

Your Company
Your Address

Tel: {user['phone']}{mobile_txt}
Mail: {user['email']}

www.google.it/ | www.google.com/
Follow us: google | LinkedIn

This e-mail may contain confidential and/or privileged information. If you are not the intended recipient (or have received this e-mail in error) please notify the sender immediately and destroy this e-mail. Any unauthorized copying, disclosure or distribution of the material in this e-mail is strictly forbidden."""
        
        return txt
    
    def deploy_signature_to_user(self, user, target_username=None):
        """
        Distribuisce la firma per un utente specifico
        
        Args:
            user: Dizionario con dati utente
            target_username: Username Windows (default: username AD dell'utente)
        
        Returns:
            True se successo, False altrimenti
        """
        if not target_username:
            target_username = user['username']
        
        # Percorso firma Outlook
        signature_folder = Path(f"C:\\Users\\{target_username}\\AppData\\Roaming\\Microsoft\\Signatures")
        
        # Crea cartella se non esiste
        try:
            signature_folder.mkdir(parents=True, exist_ok=True)
        except Exception as e:
            print(f"✗ Errore creazione cartella firma per {user['display_name']}: {e}")
            return False
        
        # Nome file firma
        signature_name = f"Firma-{self.company_info['nome'].replace(' ', '-')}"
        
        # Genera e salva firma HTML
        try:
            html_content = self.generate_signature_html(user)
            html_file = signature_folder / f"{signature_name}.htm"
            html_file.write_text(html_content, encoding='utf-8')
            print(f"  ✓ Salvata firma HTML: {html_file}")
        except Exception as e:
            print(f"  ✗ Errore salvataggio HTML: {e}")
            return False
        
        # Genera e salva firma TXT
        try:
            txt_content = self.generate_signature_txt(user)
            txt_file = signature_folder / f"{signature_name}.txt"
            txt_file.write_text(txt_content, encoding='utf-8')
            print(f"  ✓ Salvata firma TXT: {txt_file}")
        except Exception as e:
            print(f"  ✗ Errore salvataggio TXT: {e}")
            return False
        
        # Imposta firma come predefinita nel registro (opzionale)
        try:
            import winreg
            reg_path = r"Software\Microsoft\Office\16.0\Common\MailSettings"
            key = winreg.CreateKey(winreg.HKEY_CURRENT_USER, reg_path)
            winreg.SetValueEx(key, "NewSignature", 0, winreg.REG_SZ, signature_name)
            winreg.SetValueEx(key, "ReplySignature", 0, winreg.REG_SZ, signature_name)
            winreg.CloseKey(key)
            print(f"  ✓ Firma impostata come predefinita in Outlook")
        except Exception as e:
            print(f"  ⚠ Avviso: Non è stato possibile impostare la firma come predefinita: {e}")
        
        return True
    
    def save_signature_to_file(self, user, output_folder="./firme"):
        """
        Salva la firma in una cartella locale (per distribuzione manuale)
        
        Args:
            user: Dizionario con dati utente
            output_folder: Cartella dove salvare le firme
        """
        output_path = Path(output_folder)
        output_path.mkdir(exist_ok=True)
        
        user_folder = output_path / user['username']
        user_folder.mkdir(exist_ok=True)
        
        # Salva HTML
        html_content = self.generate_signature_html(user)
        html_file = user_folder / "firma.htm"
        html_file.write_text(html_content, encoding='utf-8')
        
        # Salva TXT
        txt_content = self.generate_signature_txt(user)
        txt_file = user_folder / "firma.txt"
        txt_file.write_text(txt_content, encoding='utf-8')
        
        print(f"✓ Firma salvata in: {user_folder}")
        return str(user_folder)


def display_users(users):
    """Mostra lista utenti in formato tabella"""
    if not users:
        print("Nessun utente trovato.")
        return
    
    print("\n" + "="*100)
    print(f"{'#':<4} {'Username':<15} {'Nome Completo':<30} {'Email':<35}")
    print("="*100)
    
    for idx, user in enumerate(users, 1):
        print(f"{idx:<4} {user['username']:<15} {user['display_name']:<30} {user['email']:<35}")
    
    print("="*100 + "\n")


def main():
    print("""
╔══════════════════════════════════════════════════════════════╗
║   Gestore Firme Email: Active Directory -> Outlook          ║
║   Selezione utenti specifica per aggiornamento firme        ║
╚══════════════════════════════════════════════════════════════╝
    """)
    
    # Configurazione AD - 
    print("Configurazione Active Directory")
    print("=" * 60)
    
    AD_SERVER = "adds.google.com"
    DOMAIN = "GOOGLE"
    BASE_DN = "DC=adds,DC=GOOGLE,DC=com"  # BASE DN CORRETTO!
    
    print(f"Server: {AD_SERVER}")
    print(f"Dominio: {DOMAIN}")
    print(f"Base DN: {BASE_DN}")
    print("=" * 60)
    
    # Chiedi sede per filtrare utenti
    print("\nSede da gestire:")
    print("1. Italy")
    print("2. Roma")
    print("3. Spain")
    print("4. Germania")
    print("5. Tutte le sedi")
    
    sede_choice = input("\nScegli sede (1-5): ").strip()
    
    sede_filters = {
        '1': ('Italy', 'OU=Treviso,OU=rIT,OU=Client,DC=adds,DC=google,DC=com'),
        '2': ('Roma', 'OU=Perugia,OU=rIT,OU=Client,DC=adds,DC=google,DC=com'),
        '3': ('Spain', 'OU=Verona,OU=rIT,OU=Client,DC=adds,DC=google,DC=com'),
        '4': ('Germania', 'OU=rDE,OU=Client,DC=adds,DC=google,DC=com'),
        '5': ('Tutte', 'DC=adds,DC=google,DC=com')
    }
    
    sede_name, search_base = sede_filters.get(sede_choice, sede_filters['5'])
    print(f"\n→ Sede selezionata: {sede_name}")
    print(f"→ Ricerca in: {search_base}")
    
    USERNAME = input("\nUsername AD: ").strip() or "dom.Bob"
    
    # Password in modo sicuro
    import getpass
    try:
        PASSWORD = getpass.getpass("Password: ")
    except:
        PASSWORD = input("Password: ").strip()
    
    print("\n" + "=" * 60)
    
    # Inizializza manager
    manager = ADSignatureManager(AD_SERVER, DOMAIN, search_base, USERNAME, PASSWORD)
    
    # Opzione: personalizza info azienda
    manager.company_info.update({
        'nome': 'Google',
        'indirizzo': 'Address',  # PERSONALIZZA
        'website': 'www.google.com',
        'logo_url': 'https://www.google.com/logo.png',  # PERSONALIZZA
        'colore_primario': '#0066cc',
        'colore_secondario': '#666666',
        'disclaimer': 'Questo messaggio è confidenziale. Se non siete il destinatario, vi preghiamo di eliminarlo.'
    })
    
    # Connetti ad AD
    if not manager.connect_to_ad():
        print("Impossibile connettersi ad Active Directory. Verifica le credenziali.")
        input("\nPremi Invio per chiudere...")
        return
    
    # Cerca utenti
    print("\nCerca utenti in Active Directory...")
    search_query = input("Inserisci filtro ricerca (lascia vuoto per tutti della sede): ").strip()
    
    if search_query:
        ldap_filter = f"(&(objectClass=user)(|(sAMAccountName=*{search_query}*)(displayName=*{search_query}*)(mail=*{search_query}*)))"
    else:
        ldap_filter = "(&(objectClass=user)(mail=*))"
    
    print(f"\nAvvio ricerca con filtro LDAP...")
    
    try:
        users = manager.search_users(ldap_filter)
    except Exception as e:
        print(f"\n✗ ERRORE durante la ricerca: {e}")
        import traceback
        traceback.print_exc()
        input("\nPremi Invio per chiudere...")
        return
    
    if not users:
        print("\n⚠ Nessun utente trovato con i criteri di ricerca.")
        print("\nPossibili soluzioni:")
        print("1. Verifica che ci siano utenti con email nel dominio")
        print("2. Controlla i permessi dell'account AD")
        print("3. Prova con un filtro di ricerca più specifico")
        
        retry = input("\nVuoi riprovare con un BASE DN diverso? (s/n): ").strip().lower()
        if retry == 's':
            new_base_dn = input("Inserisci nuovo BASE DN (es: OU=Users,DC=cartongrp,DC=com): ").strip()
            if new_base_dn:
                manager.base_dn = new_base_dn
                try:
                    users = manager.search_users(ldap_filter)
                except Exception as e:
                    print(f"✗ Errore: {e}")
        
        if not users:
            input("\nPremi Invio per chiudere...")
            return
    
    # Mostra utenti
    display_users(users)
    
    # Menu selezione
    while True:
        print("\nOpzioni:")
        print("1. Aggiorna firma per un singolo utente")
        print("2. Aggiorna firma per più utenti (selezione multipla)")
        print("3. Salva firme in cartella locale (distribuzione manuale)")
        print("4. Esci")
        
        choice = input("\nScegli opzione (1-4): ").strip()
        
        if choice == "1":
            # Singolo utente
            try:
                user_num = int(input(f"Inserisci numero utente (1-{len(users)}): "))
                if 1 <= user_num <= len(users):
                    selected_user = users[user_num - 1]
                    print(f"\n→ Aggiornamento firma per: {selected_user['display_name']}")
                    
                    target_username = input(f"Username Windows (default: {selected_user['username']}): ").strip()
                    if not target_username:
                        target_username = selected_user['username']
                    
                    if manager.deploy_signature_to_user(selected_user, target_username):
                        print(f"✓ Firma aggiornata con successo per {selected_user['display_name']}")
                    else:
                        print(f"✗ Errore nell'aggiornamento della firma")
                else:
                    print("Numero non valido")
            except ValueError:
                print("Input non valido")
        
        elif choice == "2":
            # Multipli utenti
            user_numbers = input(f"Inserisci numeri utenti separati da virgola (es: 1,3,5): ").strip()
            try:
                selected_nums = [int(n.strip()) for n in user_numbers.split(',')]
                valid_nums = [n for n in selected_nums if 1 <= n <= len(users)]
                
                if valid_nums:
                    print(f"\n→ Aggiornamento firme per {len(valid_nums)} utenti...")
                    
                    for num in valid_nums:
                        selected_user = users[num - 1]
                        print(f"\n  [{num}/{len(valid_nums)}] {selected_user['display_name']}")
                        
                        if manager.deploy_signature_to_user(selected_user):
                            print(f"  ✓ Completato")
                        else:
                            print(f"  ✗ Errore")
                    
                    print(f"\n✓ Processo completato per {len(valid_nums)} utenti")
                else:
                    print("Nessun numero valido inserito")
            except ValueError:
                print("Input non valido")
        
        elif choice == "3":
            # Salva in cartella locale
            output_folder = input("Cartella output (default: ./firme): ").strip()
            if not output_folder:
                output_folder = "./firme"
            
            user_numbers = input(f"Utenti da esportare (es: 1,2,3 o 'tutti'): ").strip().lower()
            
            if user_numbers == "tutti":
                selected_users = users
            else:
                try:
                    selected_nums = [int(n.strip()) for n in user_numbers.split(',')]
                    selected_users = [users[n-1] for n in selected_nums if 1 <= n <= len(users)]
                except:
                    print("Input non valido")
                    continue
            
            print(f"\n→ Salvataggio {len(selected_users)} firme in {output_folder}...")
            for user in selected_users:
                manager.save_signature_to_file(user, output_folder)
            
            print(f"\n✓ {len(selected_users)} firme salvate con successo")
        
        elif choice == "4":
            print("\nArrivederci!")
            break
        
        else:
            print("Opzione non valida")


if __name__ == "__main__":
    try:
        main()
    except KeyboardInterrupt:
        print("\n\nInterrotto dall'utente.")
    except Exception as e:
        print(f"\n✗ Errore: {e}")
        import traceback
        traceback.print_exc()