"""
Tests for invoice_monitor.py — detection logic only (no Outlook required).
Run with: pytest tests/ -v
"""

from datetime import datetime, timedelta
from invoice_monitor import EmailNFDetector, FakeEmail, FakeAttachments


# ---------------------------------------------------------------------------
# Fixtures
# ---------------------------------------------------------------------------

def make_detector() -> EmailNFDetector:
    return EmailNFDetector(dias_para_busca=7, apenas_nao_lidos=True)


def recent_unread(subject="", body="", attachments=None) -> FakeEmail:
    """Helper: creates a recent unread email with given properties."""
    return FakeEmail(
        Subject=subject,
        Body=body,
        ReceivedTime=datetime.now() - timedelta(hours=2),
        UnRead=True,
        SenderName="Test Sender",
        Attachments=FakeAttachments(attachments or []),
    )


# ---------------------------------------------------------------------------
# Subject detection
# ---------------------------------------------------------------------------

def test_detects_nfe_in_subject():
    detector = make_detector()
    email = recent_unread(subject="Envio de NF-e referente pedido 123")
    results = detector.filtrar_emails([email])
    assert len(results) == 1
    assert results[0].motivo_deteccao == "assunto"


def test_detects_nota_fiscal_in_subject():
    detector = make_detector()
    email = recent_unread(subject="Nota Fiscal referente ao pedido 456")
    results = detector.filtrar_emails([email])
    assert len(results) == 1
    assert results[0].motivo_deteccao == "assunto"


def test_detects_danfe_in_subject():
    detector = make_detector()
    email = recent_unread(subject="DANFE em anexo")
    results = detector.filtrar_emails([email])
    assert len(results) == 1


# ---------------------------------------------------------------------------
# Body detection
# ---------------------------------------------------------------------------

def test_detects_nf_keyword_in_body():
    detector = make_detector()
    email = recent_unread(subject="Documentos fiscais", body="Segue nota fiscal para sua análise.")
    results = detector.filtrar_emails([email])
    assert len(results) == 1
    assert results[0].motivo_deteccao == "corpo"


def test_body_scan_limited_to_5000_chars():
    """Keyword buried beyond 5000 chars should NOT be detected."""
    detector = make_detector()
    padding = "x" * 5001
    email = recent_unread(subject="Assunto normal", body=padding + "nota fiscal")
    results = detector.filtrar_emails([email])
    assert len(results) == 0


# ---------------------------------------------------------------------------
# Attachment detection
# ---------------------------------------------------------------------------

def test_detects_nfe_xml_attachment():
    detector = make_detector()
    email = recent_unread(subject="Documentos", attachments=["NFe_123.xml"])
    results = detector.filtrar_emails([email])
    assert len(results) == 1
    assert results[0].motivo_deteccao == "anexo"


def test_detects_danfe_pdf_attachment():
    detector = make_detector()
    email = recent_unread(subject="Envio", attachments=["danfe_2024.pdf"])
    results = detector.filtrar_emails([email])
    assert len(results) == 1


# ---------------------------------------------------------------------------
# False positive prevention
# ---------------------------------------------------------------------------

def test_ignores_generic_nota_in_attachment():
    """'nota_reuniao.pdf' should NOT be detected — generic 'nota' pattern was removed."""
    detector = make_detector()
    email = recent_unread(subject="Ata de reunião", attachments=["nota_reuniao.pdf"])
    results = detector.filtrar_emails([email])
    assert len(results) == 0


def test_ignores_unrelated_email():
    detector = make_detector()
    email = recent_unread(subject="Reunião de planejamento", body="Pauta da semana.")
    results = detector.filtrar_emails([email])
    assert len(results) == 0


# ---------------------------------------------------------------------------
# Read/unread filter
# ---------------------------------------------------------------------------

def test_ignores_read_email_when_unread_only():
    detector = make_detector()  # apenas_nao_lidos=True
    email = FakeEmail(
        Subject="NF-e em anexo",
        Body="",
        ReceivedTime=datetime.now() - timedelta(hours=1),
        UnRead=False,  # already read
        SenderName="Fornecedor",
        Attachments=FakeAttachments([]),
    )
    results = detector.filtrar_emails([email])
    assert len(results) == 0


def test_detects_read_email_when_all_emails_mode():
    detector = EmailNFDetector(dias_para_busca=7, apenas_nao_lidos=False)
    email = FakeEmail(
        Subject="NF-e em anexo",
        Body="",
        ReceivedTime=datetime.now() - timedelta(hours=1),
        UnRead=False,
        SenderName="Fornecedor",
        Attachments=FakeAttachments([]),
    )
    results = detector.filtrar_emails([email])
    assert len(results) == 1


# ---------------------------------------------------------------------------
# Date filter
# ---------------------------------------------------------------------------

def test_ignores_email_outside_date_window():
    detector = make_detector()  # dias_para_busca=7
    email = FakeEmail(
        Subject="NF-e antiga",
        Body="",
        ReceivedTime=datetime.now() - timedelta(days=10),  # too old
        UnRead=True,
        SenderName="Fornecedor",
        Attachments=FakeAttachments([]),
    )
    results = detector.filtrar_emails([email])
    assert len(results) == 0


# ---------------------------------------------------------------------------
# Result metadata
# ---------------------------------------------------------------------------

def test_result_contains_correct_sender():
    detector = make_detector()
    email = recent_unread(subject="Envio de NF-e")
    email.SenderName = "Fornecedor ABC Ltda"
    results = detector.filtrar_emails([email])
    assert results[0].remetente == "Fornecedor ABC Ltda"


def test_result_contains_correct_subject():
    detector = make_detector()
    email = recent_unread(subject="NF-e pedido 789")
    results = detector.filtrar_emails([email])
    assert results[0].assunto == "NF-e pedido 789"
