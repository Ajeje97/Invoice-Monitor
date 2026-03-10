"""Monitor simples de Notas Fiscais no Outlook para iniciantes.

Objetivo:
- Procurar e-mails com sinais de Nota Fiscal (assunto, corpo ou anexos)
- Criar uma tarefa no Outlook com o alerta "Dar entrada no fiscal"
- Registrar em CSV os e-mails detectados

Requisitos:
- Windows
- Outlook instalado e configurado
- pywin32: pip install pywin32
"""

from __future__ import annotations

import argparse
import csv
import re
import sys
from dataclasses import dataclass, field
from datetime import datetime, timedelta
from pathlib import Path
from typing import Iterable

# ---------------------------------------------------------------------------
# Palavras-chave e padrões de detecção
# ---------------------------------------------------------------------------

PALAVRAS_CHAVE_NF = [
    "nota fiscal",
    "nfe",
    "nf-e",
    "danfe",
    "xml de nota",
]

# Padrão "nota" removido: causa muitos falsos positivos.
# Use padrões mais específicos para nomes de anexo.
PADROES_ANEXO_NF = [
    re.compile(r"nfe", re.IGNORECASE),
    re.compile(r"nf[-_ ]?e", re.IGNORECASE),
    re.compile(r"danfe", re.IGNORECASE),
    re.compile(r"nota[-_ ]?fiscal", re.IGNORECASE),
]

# Tamanho máximo do corpo a inspecionar (evita varrer e-mails gigantes)
_MAX_CORPO_CHARS = 5_000

# Colunas do CSV de registro
_CABECALHO_CSV = [
    "data_registro",
    "remetente",
    "assunto",
    "recebido_em",
    "motivo_deteccao",
    "tarefa_criada",
    "erro",
]


# ---------------------------------------------------------------------------
# Detecção
# ---------------------------------------------------------------------------

@dataclass
class ResultadoEmail:
    """Resultado da análise de um único e-mail."""

    remetente: str
    assunto: str
    recebido_em: str
    motivo_deteccao: str  # "assunto" | "corpo" | "anexo"
    tarefa_criada: bool = False
    erro: str = ""


class EmailNFDetector:
    def __init__(self, dias_para_busca: int = 7, apenas_nao_lidos: bool = True) -> None:
        self.dias_para_busca = dias_para_busca
        self.apenas_nao_lidos = apenas_nao_lidos

    # ------------------------------------------------------------------
    # Conexão
    # ------------------------------------------------------------------

    def conectar_outlook(self):
        try:
            import win32com.client  # type: ignore
        except ImportError as exc:
            raise RuntimeError(
                "pywin32 não encontrado. Instale com: pip install pywin32"
            ) from exc

        outlook = win32com.client.Dispatch("Outlook.Application")
        namespace = outlook.GetNamespace("MAPI")
        return outlook, namespace

    # ------------------------------------------------------------------
    # Busca com filtro MAPI (muito mais rápido que iterar tudo)
    # ------------------------------------------------------------------

    def obter_emails(self, caixa_entrada):
        """Retorna coleção já filtrada por data (e leitura, se configurado)."""
        emails = caixa_entrada.Items
        emails.Sort("[ReceivedTime]", True)

        limite = (datetime.now() - timedelta(days=self.dias_para_busca)).strftime(
            "%m/%d/%Y %H:%M %p"
        )
        filtro = f"[ReceivedTime] >= '{limite}'"
        if self.apenas_nao_lidos:
            filtro += " AND [UnRead] = True"

        return emails.Restrict(filtro)

    # ------------------------------------------------------------------
    # Detecção de palavras-chave
    # ------------------------------------------------------------------

    def texto_indica_nf(self, texto: str) -> bool:
        texto_normalizado = texto.lower()
        return any(chave in texto_normalizado for chave in PALAVRAS_CHAVE_NF)

    def anexos_indicam_nf(self, anexos) -> bool:
        if anexos is None:
            return False
        quantidade = getattr(anexos, "Count", 0)
        if quantidade == 0:
            return False
        for i in range(1, quantidade + 1):
            anexo = anexos.Item(i)
            nome = str(getattr(anexo, "FileName", ""))
            if any(p.search(nome) for p in PADROES_ANEXO_NF):
                return True
        return False

    def motivo_deteccao(self, email) -> str | None:
        """Retorna o primeiro motivo encontrado ou None se não for NF."""
        assunto = str(getattr(email, "Subject", ""))
        if self.texto_indica_nf(assunto):
            return "assunto"

        corpo = str(getattr(email, "Body", ""))[:_MAX_CORPO_CHARS]
        if self.texto_indica_nf(corpo):
            return "corpo"

        if self.anexos_indicam_nf(getattr(email, "Attachments", None)):
            return "anexo"

        return None

    # ------------------------------------------------------------------
    # Filtragem
    # ------------------------------------------------------------------

    def _to_datetime(self, valor) -> datetime | None:
        """Converte pywintypes.datetime (ou datetime) para datetime padrão."""
        if valor is None:
            return None
        if isinstance(valor, datetime):
            return valor.replace(tzinfo=None)
        # pywintypes.datetime tem .timestamp()
        try:
            return datetime.fromtimestamp(valor.timestamp())
        except Exception:
            return None

    def filtrar_emails(self, emails) -> list[ResultadoEmail]:
        """Filtra e classifica e-mails como NF, retornando ResultadoEmail."""
        limite = datetime.now() - timedelta(days=self.dias_para_busca)
        encontrados: list[ResultadoEmail] = []

        for email in emails:
            try:
                recebido_em = self._to_datetime(getattr(email, "ReceivedTime", None))
                if recebido_em is None or recebido_em < limite:
                    continue

                # Quando obter_emails usa Restrict, este check é redundante
                # mas mantemos para o autoteste (que passa lista pura).
                if self.apenas_nao_lidos and not bool(getattr(email, "UnRead", False)):
                    continue

                motivo = self.motivo_deteccao(email)
                if motivo is None:
                    continue

                encontrados.append(
                    ResultadoEmail(
                        remetente=str(getattr(email, "SenderName", "Desconhecido")),
                        assunto=str(getattr(email, "Subject", "(sem assunto)")),
                        recebido_em=recebido_em.strftime("%Y-%m-%d %H:%M"),
                        motivo_deteccao=motivo,
                    )
                )
            except Exception as exc:  # noqa: BLE001
                # E-mail corrompido ou com propriedade inacessível — registra e segue
                print(f"  [aviso] E-mail ignorado por erro: {exc}")

        return encontrados


# ---------------------------------------------------------------------------
# Criação de tarefa
# ---------------------------------------------------------------------------

def criar_tarefa_fiscal(outlook, resultado: ResultadoEmail, nome_responsavel: str) -> None:
    ol_task_item = 3
    tarefa = outlook.CreateItem(ol_task_item)

    tarefa.Subject = f"Dar entrada no fiscal | {resultado.assunto}"
    tarefa.Body = (
        "Alerta automático:\n"
        "Foi detectada possível Nota Fiscal em e-mail.\n\n"
        f"Responsável: {nome_responsavel}\n"
        f"Remetente: {resultado.remetente}\n"
        f"Assunto: {resultado.assunto}\n"
        f"Recebido em: {resultado.recebido_em}\n"
        f"Motivo da detecção: {resultado.motivo_deteccao}\n"
    )
    tarefa.DueDate = datetime.now() + timedelta(days=1)
    tarefa.ReminderSet = True
    tarefa.ReminderTime = datetime.now() + timedelta(minutes=5)
    tarefa.Save()


# ---------------------------------------------------------------------------
# Registro CSV
# ---------------------------------------------------------------------------

def registrar_csv(caminho: Path, resultados: list[ResultadoEmail]) -> None:
    """Acrescenta (append) os resultados ao arquivo CSV."""
    novo_arquivo = not caminho.exists()
    with caminho.open("a", newline="", encoding="utf-8-sig") as f:
        writer = csv.DictWriter(f, fieldnames=_CABECALHO_CSV)
        if novo_arquivo:
            writer.writeheader()
        agora = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        for r in resultados:
            writer.writerow(
                {
                    "data_registro": agora,
                    "remetente": r.remetente,
                    "assunto": r.assunto,
                    "recebido_em": r.recebido_em,
                    "motivo_deteccao": r.motivo_deteccao,
                    "tarefa_criada": "sim" if r.tarefa_criada else "nao",
                    "erro": r.erro,
                }
            )


# ---------------------------------------------------------------------------
# Orquestração principal
# ---------------------------------------------------------------------------

def processar_alertas(
    dias: int,
    apenas_nao_lidos: bool,
    responsavel: str,
    limite_alertas: int,
    caminho_csv: Path,
) -> list[ResultadoEmail]:
    detector = EmailNFDetector(dias_para_busca=dias, apenas_nao_lidos=apenas_nao_lidos)
    outlook, namespace = detector.conectar_outlook()
    caixa_entrada = namespace.GetDefaultFolder(6)  # 6 = Inbox
    emails = detector.obter_emails(caixa_entrada)
    encontrados = detector.filtrar_emails(emails)

    for resultado in encontrados[:limite_alertas]:
        try:
            criar_tarefa_fiscal(outlook, resultado, responsavel)
            resultado.tarefa_criada = True
        except Exception as exc:  # noqa: BLE001
            resultado.erro = str(exc)
            print(f"  [erro] Falha ao criar tarefa para '{resultado.assunto}': {exc}")

    registrar_csv(caminho_csv, encontrados[:limite_alertas])
    return encontrados


# ---------------------------------------------------------------------------
# Autoteste (sem Outlook)
# ---------------------------------------------------------------------------

@dataclass
class FakeAttachment:
    FileName: str


class FakeAttachments:
    def __init__(self, names: list[str]) -> None:
        self._items = [FakeAttachment(name) for name in names]

    @property
    def Count(self) -> int:
        return len(self._items)

    def Item(self, index: int) -> FakeAttachment:
        return self._items[index - 1]


@dataclass
class FakeEmail:
    Subject: str
    Body: str
    ReceivedTime: datetime
    UnRead: bool
    SenderName: str
    Attachments: FakeAttachments


def executar_autoteste() -> None:
    detector = EmailNFDetector(dias_para_busca=7, apenas_nao_lidos=True)

    emails: list[FakeEmail] = [
        FakeEmail(
            Subject="Envio de NF-e referente pedido 123",
            Body="Segue XML da nota.",
            ReceivedTime=datetime.now() - timedelta(hours=2),
            UnRead=True,
            SenderName="Fornecedor A",
            Attachments=FakeAttachments(["NFe_123.xml"]),
        ),
        FakeEmail(
            Subject="Reunião de planejamento",
            Body="Pauta da semana",
            ReceivedTime=datetime.now() - timedelta(hours=1),
            UnRead=True,
            SenderName="Time Interno",
            Attachments=FakeAttachments(["agenda.pdf"]),
        ),
        FakeEmail(
            # Não lido=False → deve ser ignorado com apenas_nao_lidos=True
            Subject="DANFE em anexo",
            Body="",
            ReceivedTime=datetime.now() - timedelta(days=2),
            UnRead=False,
            SenderName="Fornecedor B",
            Attachments=FakeAttachments(["danfe_2024.pdf"]),
        ),
        FakeEmail(
            # Palavra "nota" no nome do anexo NÃO deve mais detectar — padrão removido
            Subject="Contrato aprovado",
            Body="",
            ReceivedTime=datetime.now() - timedelta(hours=3),
            UnRead=True,
            SenderName="Jurídico",
            Attachments=FakeAttachments(["nota_reuniao.pdf"]),
        ),
    ]

    encontrados = detector.filtrar_emails(emails)

    assert len(encontrados) == 1, (
        f"Autoteste falhou: esperado 1 e-mail, encontrado {len(encontrados)}."
    )
    assert encontrados[0].motivo_deteccao == "assunto", (
        f"Autoteste falhou: motivo esperado 'assunto', obtido '{encontrados[0].motivo_deteccao}'."
    )

    # Testa registro CSV em arquivo temporário
    import tempfile

    with tempfile.NamedTemporaryFile(suffix=".csv", delete=False) as tmp:
        caminho_tmp = Path(tmp.name)

    registrar_csv(caminho_tmp, encontrados)
    linhas = caminho_tmp.read_text(encoding="utf-8-sig").splitlines()
    assert len(linhas) == 2, f"Autoteste CSV falhou: esperado 2 linhas, obtido {len(linhas)}."
    caminho_tmp.unlink()

    print("Autoteste OK: detecção e registro CSV funcionando sem Outlook.")


# ---------------------------------------------------------------------------
# CLI
# ---------------------------------------------------------------------------

def parse_args(argv: Iterable[str] | None = None) -> argparse.Namespace:
    parser = argparse.ArgumentParser(
        description=(
            "Detecta possíveis Notas Fiscais em e-mails do Outlook, "
            "cria alerta no fiscal e registra em CSV."
        )
    )
    parser.add_argument(
        "--dias",
        type=int,
        default=7,
        help="Quantos dias atrás verificar (padrão: 7)",
    )
    parser.add_argument(
        "--inclui-lidos",
        action="store_true",
        help="Inclui e-mails já lidos na varredura",
    )
    parser.add_argument(
        "--responsavel",
        type=str,
        default="Equipe Fiscal",
        help="Nome do responsável no texto do alerta",
    )
    parser.add_argument(
        "--limite-alertas",
        type=int,
        default=20,
        help="Máximo de tarefas criadas por execução (padrão: 20)",
    )
    parser.add_argument(
        "--csv",
        type=Path,
        default=Path("alertas_nf.csv"),
        help="Caminho do arquivo CSV de registro (padrão: alertas_nf.csv)",
    )
    parser.add_argument(
        "--autoteste",
        action="store_true",
        help="Executa autoteste local sem Outlook para validar regras de detecção",
    )
    return parser.parse_args(argv)


def main() -> None:
    args = parse_args()

    if args.autoteste:
        executar_autoteste()
        sys.exit(0)

    try:
        resultados = processar_alertas(
            dias=args.dias,
            apenas_nao_lidos=not args.inclui_lidos,
            responsavel=args.responsavel,
            limite_alertas=args.limite_alertas,
            caminho_csv=args.csv,
        )
    except RuntimeError as exc:
        print(f"Erro: {exc}")
        sys.exit(1)

    total = len(resultados)
    criadas = sum(1 for r in resultados if r.tarefa_criada)

    if total == 0:
        print("Nenhum e-mail com sinal de Nota Fiscal foi encontrado.")
    else:
        print(f"{total} e-mail(s) detectado(s). {criadas} tarefa(s) criada(s).")
        print(f"Registro salvo em: {args.csv.resolve()}")


if __name__ == "__main__":
    main()