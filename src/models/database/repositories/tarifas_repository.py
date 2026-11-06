import mysql.connector
from mysql.connector import Error
from typing import List, Tuple, Any, Dict

from src.services.log_service import registrar_log

class TarifasRepository:
    """Repository responsável por operações de INSERT na Base de Dados (histórico TAP IN/OUT)."""

    def __init__(self, mysql_cfg: Dict[str, Any]):
        self.mysql_cfg = mysql_cfg
        self.conn = None
        self.cursor = None

    def connect(self):
        """Abre conexão e cursor."""
        try:
            registrar_log("Iniciado", "Sucesso", f"Tentativa de conexão com Base de Dados.")
            self.conn = mysql.connector.connect(
                **self.mysql_cfg,
                use_pure=True,
                connection_timeout=10
            )
            self.cursor = self.conn.cursor()
        except Error as e:
            registrar_log("Iniciado", "Erro", f"Erro ao conectar na Base de Dados: {e}")
            raise 

    def close(self):
        """Fecha cursor e conexão."""
        if self.cursor:
            self.cursor.close()
        if self.conn:
            self.conn.close()

    def insert_many(self, sql: str, chunk: List[Tuple]) -> int:
        """
        Executa executemany() com commit e retorna número de linhas afetadas.
        Preserva o comportamento original (rowcount pode ser -1).
        """
        self.cursor.executemany(sql, chunk)
        self.conn.commit()
        afetadas = (self.cursor.rowcount if self.cursor.rowcount not in (None, -1) else len(chunk))
        return afetadas