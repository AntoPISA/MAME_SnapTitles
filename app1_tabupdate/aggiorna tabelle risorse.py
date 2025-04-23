import os
import pyodbc
import configparser
import logging
from datetime import datetime

# Funzione per configurare il logger
def setup_logger(log_folder):
    # Verifica che il percorso sia valido
    if not log_folder or not isinstance(log_folder, str) or not os.path.isabs(log_folder):
        logging.error(f"Percorso della cartella dei log non valido: {log_folder}")
        return
    
    # Crea la cartella dei log se non esiste
    try:
        if not os.path.exists(log_folder):
            os.makedirs(log_folder)
            logging.info(f"Cartella dei log creata: {log_folder}")
    except Exception as e:
        logging.error(f"Errore durante la creazione della cartella dei log: {e}")
        return
    
    # Genera un nome file di log con data e ora
    log_filename = f"log_{datetime.now().strftime('%Y%m%d_%H%M%S')}.log"
    log_file = os.path.join(log_folder, log_filename)
    
    # Configura il logger
    logging.basicConfig(
        filename=log_file,
        level=logging.INFO,
        format='%(asctime)s - %(levelname)s - %(message)s',
        datefmt='%Y-%m-%d %H:%M:%S'
    )
    logging.info("File di log creato.")
    logging.info("Applicazione: Aggiorna progetto-SNAPs v1.0")

# Funzione per connettersi al database Access
def connect_to_access(db_path):
    try:
        # Stringa di connessione per Access
        conn_str = r'DRIVER={Microsoft Access Driver (*.mdb, *.accdb)};DBQ=' + db_path
        conn = pyodbc.connect(conn_str)
        return conn
    except Exception as e:
        logging.error(f"Errore durante la connessione al database: {e}")
        return None

# Funzione per verificare e creare la tabella
def ensure_table_exists(conn, table_name):
    cursor = conn.cursor()
    try:
        # Verifica se la tabella esiste
        cursor.execute(f"""
            SELECT COUNT(*) 
            FROM MSysObjects 
            WHERE Name='{table_name}' AND Type=1
        """)
        table_exists = cursor.fetchone()[0]
        
        if not table_exists:
            # Crea la tabella con un campo corrispondente al nome della cartella
            cursor.execute(f"""
                CREATE TABLE [{table_name}] (
                    ID AUTOINCREMENT PRIMARY KEY,
                    [{table_name}] TEXT
                )
            """)
            conn.commit()
            logging.info(f"Tabella '{table_name}' creata nel database.")
    except Exception as e:
        logging.error(f"Errore durante la verifica o la creazione della tabella '{table_name}': {e}")
    finally:
        cursor.close()

# Funzione per svuotare una tabella
def clear_table(conn, table_name):
    cursor = conn.cursor()
    try:
        # Svuota la tabella
        cursor.execute(f"DELETE FROM [{table_name}]")
        conn.commit()
        logging.info(f"Tabella '{table_name}' svuotata.")
    except Exception as e:
        logging.error(f"Errore durante lo svuotamento della tabella '{table_name}': {e}")
        conn.rollback()
    finally:
        cursor.close()

# Funzione per inserire i nomi dei file nella tabella
def insert_file_names(conn, folder_path, table_name):
    cursor = conn.cursor()
    try:
        # Leggi tutti i file nella cartella e filtra solo quelli con estensione .png
        files = [f for f in os.listdir(folder_path) if f.lower().endswith('.png')]
        
        # Rimuovi l'estensione .png dai nomi dei file
        files_no_extension = [os.path.splitext(f)[0] for f in files]
        
        total_files = len(files_no_extension)
        logging.info(f"Elaborazione della cartella '{table_name}' - File PNG trovati: {total_files}")
        
        # Inserisci ogni nome di file (senza estensione) nella tabella
        for file_name in files_no_extension:
            cursor.execute(f"INSERT INTO [{table_name}] ([{table_name}]) VALUES (?)", file_name)
        
        # Conferma le modifiche
        conn.commit()
        logging.info(f"Inseriti {total_files} file PNG (senza estensione) nella tabella '{table_name}'.")
    except Exception as e:
        logging.error(f"Errore durante l'inserimento dei dati: {e}")
        conn.rollback()
    finally:
        cursor.close()

# Funzione per leggere il file .ini
def read_config(config_path):
    config = configparser.ConfigParser()
    try:
        config.read(config_path)
        folders = dict(config['Folders'])  # Percorsi delle cartelle
        db_path = config['Database']['Path']  # Percorso del database
        log_folder = config['Log']['LogFolder']  # Percorso della cartella dei log
        return folders, db_path, log_folder
    except Exception as e:
        logging.error(f"Errore durante la lettura del file di configurazione: {e}")
        return None, None, None

# Funzione principale
def main():
    # Percorso del file di configurazione
    config_path = r"W:\AGGIORNA_DB_RISORSE\config.ini"  # Modifica con il tuo percorso
    
    # Leggi il file di configurazione
    folders, db_path, log_folder = read_config(config_path)
    if not folders or not db_path or not log_folder:
        logging.error("Impossibile proseguire a causa di errori nel file di configurazione.")
        return
    
    # Configura il logger
    setup_logger(log_folder)
    logging.info("Avvio dell'applicazione: Aggiorna progetto-SNAPs v1.0")
    
    # Connessione al database
    conn = connect_to_access(db_path)
    if conn:
        try:
            # Elabora ogni cartella
            for table_name, folder_path in folders.items():
                logging.info(f"Inizio elaborazione della cartella '{table_name}'...")
                
                # Verifica e crea la tabella se necessario
                ensure_table_exists(conn, table_name)
                
                # Svuota la tabella prima di inserire nuovi dati
                clear_table(conn, table_name)
                
                # Inserisci i nomi dei file nella tabella
                insert_file_names(conn, folder_path, table_name)
        finally:
            # Chiudi la connessione
            conn.close()
            logging.info("Connessione al database chiusa.")
    
    # Registra la fine dell'applicazione
    logging.info("Applicazione terminata.")

if __name__ == "__main__":
    main()