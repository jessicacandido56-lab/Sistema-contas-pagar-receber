from dataclasses import dataclass
from datetime import date


@dataclass
class ContaPagar:
    data: date
    fornecedor: str
    documento: str
    valor: float


@dataclass
class ContaReceber:
    vencimento: date
    ordem_servico: str
    cliente: str
    numero_documento: str
    valor: float