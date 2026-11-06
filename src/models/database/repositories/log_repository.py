import mysql.connector
from mysql.connector import Error
from datetime import datetime

class LogRepository:
    def __init__(self, cfg):
        self.cfg = cfg

    def inserir_log(self, rpa: str, processo: str, status: str, descricao: str):
        """
        Insere log com carimbo de data explícito.
        - processo: 'Iniciado' | 'Finalizado'
        - status:   'Sucesso'  | 'Erro'
        - descricao: até 1000 chars
        """
        data_log = datetime.now().strftime("%Y-%m-%d %H:%M:%S")

        conn = None
        cursor = None
        try:
            # Tenta conectar
            conn = mysql.connector.connect(
                **self.cfg,
                use_pure=True,
                connection_timeout=10
                )
            cursor = conn.cursor()
            sql = """
                INSERT INTO tbl_rpa_log_roaming_gestao_tarifas
                (rpa, processo, status, descricao, data_log)
                VALUES (%s, %s, %s, %s, %s)
            """
            cursor.execute(sql, (rpa, processo, status, descricao, data_log))
            conn.commit()

        except Error as e:
            raise RuntimeError(f"Falha ao conectar/gravar na Base de Dados de logs: {e}") from e

        except Exception as e:
            # Qualquer outra exceção inesperada
            raise RuntimeError(f"Erro inesperado ao inserir log: {e}") from e

        finally:
            # Cleanup seguro
            try:
                if cursor is not None:
                    cursor.close()
            finally:
                if conn is not None and conn.is_connected():
                    conn.close()