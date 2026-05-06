# modelo.py

from sqlmodel import Field, SQLModel, create_engine
from enum import Enum
# modelo.py

from sqlmodel import Field, SQLModel, create_engine, Relationship
from enum import Enum
from datetime import date

class Bancos(Enum):
    NUBANK = 'Nubank'
    SANTANDER = 'Santander'
    INTER = 'Inter'

class Status(Enum):
    ATIVO = 'Ativo'
    INATIVO = 'Inativo'

class Tipos(Enum):
    ENTRADA = 'Entrada'
    SAIDA = 'Saida'

class Conta(SQLModel, table=True):
    id: int = Field(primary_key=True)
    banco: str = Field(default=Bancos.NUBANK.value)
    status: str = Field(default=Status.ATIVO.value)
    valor: float

class Historico(SQLModel, table=True):
    id: int = Field(primary_key=True)
    conta_id: int = Field(foreign_key="conta.id")
    tipo: str = Field(default=Tipos.ENTRADA.value)
    valor: float
    data: date

sqlite_file_name = "database.db"
sqlite_url = f"sqlite:///{sqlite_file_name}"
engine = create_engine(sqlite_url, echo=False)

def create_db_and_tables():
    SQLModel.metadata.create_all(engine)

create_db_and_tables()