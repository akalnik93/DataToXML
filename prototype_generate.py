"""Create a proof-of-concept MP XML from an application, contract and consent DOCX."""

from __future__ import annotations

import datetime as dt
import re
import uuid
import zipfile
from pathlib import Path
from xml.etree import ElementTree as ET


def docx_text(path: Path) -> str:
    with zipfile.ZipFile(path) as archive:
        root = ET.fromstring(archive.read("word/document.xml"))
    ns = {"w": "http://schemas.openxmlformats.org/wordprocessingml/2006/main"}
    return " ".join((node.text or "") for node in root.findall(".//w:t", ns))


MONTHS = {
    "января": 1, "февраля": 2, "марта": 3, "апреля": 4,
    "мая": 5, "июня": 6, "июля": 7, "августа": 8,
    "сентября": 9, "октября": 10, "ноября": 11, "декабря": 12,
}


def identify_documents(paths: list[Path]) -> tuple[str, str | None, str | None]:
    texts = [(path, docx_text(path)) for path in paths if path.suffix.lower() == ".docx"]
    application = next((text for _, text in texts if re.search(r"ЗАЯВЛЕНИЕ\s*№", text, re.I)), None)
    contract = next((text for _, text in texts if re.search(r"ДОГОВОР\s+ПОДРЯДА", text, re.I)), None)
    consent = next((text for _, text in texts if re.search(r"СОГЛАСИЕ\s+на\s+обработку\s+персональных\s+данных", text, re.I)), None)
    if application is None:
        raise ValueError("Среди выбранных DOCX не найдено заявление (ожидается текст «ЗАЯВЛЕНИЕ №»)")
    return application, contract, consent


def extract_fields(application: str, contract: str | None, consent: str | None) -> tuple[str, str, str, str, str, str | None]:
    client = re.search(r"от\s+([А-ЯЁ][а-яё-]+)\s+([А-ЯЁ][а-яё-]+)\s+([А-ЯЁ][а-яё-]+)", application)
    agreement = re.search(r"ЗАЯВЛЕНИЕ\s*№\s*([\wА-ЯЁа-яё/-]+)", application, re.IGNORECASE)
    # The application also contains passport and law dates; use the date labeled "Дата".
    date = re.search(r"\b(\d{2}\.\d{2}\.\d{4})\s+Дата\b", application, re.IGNORECASE)
    if not (client and agreement and date):
        raise ValueError("Не удалось уверенно извлечь ФИО, номер заявления или дату из DOCX")
    number, iso_date = agreement.group(1).upper(), f"{date.group(1)[6:10]}-{date.group(1)[3:5]}-{date.group(1)[:2]}"
    if contract:
        contract_number = re.search(r"ДОГОВОР\s+ПОДРЯДА[^\n]{0,160}?№\s*([\wА-ЯЁа-яё/-]+)", contract, re.I)
        contract_date = re.search(r"[«\"]\s*(\d{1,2})\s*[»\"]\s+([А-ЯЁа-яё]+)\s+(\d{4})\s*г", contract, re.I)
        if contract_number:
            number = contract_number.group(1).upper()
        if contract_date and contract_date.group(2).lower() in MONTHS:
            day, month, year = int(contract_date.group(1)), MONTHS[contract_date.group(2).lower()], int(contract_date.group(3))
            iso_date = dt.date(year, month, day).isoformat()
    snils = None
    if consent:
        labeled = re.search(r"(?:СНИЛС|страховой\s+номер[^:]{0,60}:?)\s*([\d\s-]{11,24})", consent, re.I)
        if labeled:
            digits = re.sub(r"\D", "", labeled.group(1))
            if len(digits) == 11:
                snils = digits
    return (*client.groups(), number, iso_date, snils)


def extract_additional_fields(application: str, contract: str | None, consent: str | None) -> list[str]:
    """Extract useful values for review when this MP template has no safe target node."""
    found: list[str] = []
    if consent:
        normalized = re.sub(r"П\s+а(спорт)", r"Паспорт", consent, flags=re.I)
        normalized = re.sub(r"В\s+ыдан", "Выдан", normalized, flags=re.I)
        passport = re.search(
            r"паспорт\s+(\d{4})\s*№?\s*(\d{6})\s+выдан\s+(.+?),\s*(\d{2}\.\d{2}\.\d{4})",
            normalized, re.I,
        )
        if passport:
            found.append(f"Паспорт: {passport.group(1)} № {passport.group(2)}, выдан {passport.group(4)} ({passport.group(3)})")
        address = re.search(r"зарегистрированный\s+по\s+адресу:\s*(.+?)\s*,\s*даю", consent, re.I)
        if address:
            found.append(f"Адрес регистрации: {address.group(1).strip()}")
    phone = re.search(r"([+]?[78][\d\s()\-]{9,18})\s+контактный\s+телефон", application, re.I)
    if phone:
        phone_value = "".join(phone.group(1).split())
        found.append(f"Телефон: {phone_value}")
    parcel_address = None
    if contract:
        parcel_address = re.search(r"расположен\s+по\s+адресу:\s*(.+?)\s*\.\s*1\.3", contract, re.I)
    if parcel_address:
        found.append(f"Адрес участка: {parcel_address.group(1).strip()}")
        cost = re.search(r"общая\s+стоимость\s+работ\s+составляет\s+([\d\s]+(?:[.,]\d{2})?)", contract, re.I)
        if cost:
            cost_value = "".join(cost.group(1).split())
            found.append(f"Стоимость работ: {cost_value} руб.")
        advance = re.search(r"2\.3\..{0,260}?составляет\s+([\d\s]+[.,]\d{2})\s*\(", contract, re.I | re.S)
        if advance:
            advance_value = "".join(advance.group(1).split())
            found.append(f"Аванс: {advance_value} руб.")
        duration = re.search(r"выполнить\s+работы\s+в\s+течение\s+(\d+)\s*\(([^)]*)\)\s*дней", contract, re.I)
        if duration:
            found.append(f"Срок работ: {duration.group(1)} дней")
    return found


def _load_fields(documents: list[Path]) -> tuple[tuple[str, str, str, str, str, str | None], list[str]]:
    application, contract, consent = identify_documents(documents)
    return extract_fields(application, contract, consent), extract_additional_fields(application, contract, consent)


def _summary(family: str, first: str, patronymic: str, number: str, date: str, snils: str | None, additional: list[str]) -> str:
    text = f"{family} {first} {patronymic}; договор № {number} от {date}; СНИЛС {'извлечён' if snils else 'не найден в ОПД'}"
    if additional:
        text += "\n\nИзвлечено для проверки, но не записано в XML (в шаблоне нет подтверждённых для этого полей):\n• " + "\n• ".join(additional)
    return text


def _set_client(root: ET.Element, family: str, first: str, patronymic: str, snils: str | None) -> None:
    client = root.find(".//Clients/Client/Person")
    if client is None:
        raise ValueError("В шаблоне не найден узел клиента Person")
    for tag, value in (("FamilyName", family), ("FirstName", first), ("Patronymic", patronymic)):
        node = client.find(tag)
        if node is None:
            raise ValueError(f"В шаблоне не найдено поле {tag}")
        node.text = value
    snils_node = client.find("SNILS")
    if snils_node is not None:
        # Never leave the previous client's SNILS in a generated draft.
        snils_node.text = snils


def _set_agreement(root: ET.Element, number: str, date: str) -> None:
    agreements = root.findall(".//AgreementCadWork")
    if not agreements:
        raise ValueError("В шаблоне не найден узел AgreementCadWork")
    for agreement in agreements:
        for tag, value in (("NumberAgreement", number), ("DateAgreement", date)):
            node = agreement.find(tag)
            if node is not None:
                node.text = value


def build_application_only_xml(documents: list[Path], template_xml: Path) -> tuple[bytes, str]:
    """Clear sample-specific data and file references, retaining XML structure."""
    (family, first, patronymic, number, date, snils), additional = _load_fields(documents)
    tree = ET.parse(template_xml)
    root = tree.getroot()

    # Remove all references to supporting files so ARGO does not look for absent PDFs.
    for parent in root.iter():
        for child in list(parent):
            if child.tag in {"AppliedFile", "AppliedFiles"}:
                parent.remove(child)
    # These two attachments are represented as top-level elements with Name
    # attributes rather than AppliedFile child nodes.
    for tag in ("SchemeGeodesicPlotting", "SchemeDisposition"):
        node = root.find(tag)
        if node is not None:
            root.remove(node)

    parcel = root.find(".//ExistParcel")
    parcel_node_ids = {id(node) for node in parcel.iter()} if parcel is not None else set()

    # Keep the source template's structure and attributes. ARGO parses Date as
    # DateTime, geodetic-base OrdX/OrdY as numbers, and numeric parcel values as
    # integers/doubles, so retain those source values.
    for node in root.iter():
        if node.tag in {"Date", "OrdX", "OrdY"}:
            continue
        value = (node.text or "").strip()
        numeric_parcel_value = (
            id(node) in parcel_node_ids
            and re.fullmatch(r"[+-]?\d+(?:[.,]\d+)?", value) is not None
        )
        if not numeric_parcel_value:
            node.text = None
    root.attrib["GUID"] = str(uuid.uuid4())

    _set_client(root, family, first, patronymic, snils)
    _set_agreement(root, number, date)

    data = ET.tostring(root, encoding="utf-8", xml_declaration=True)
    summary = _summary(family, first, patronymic, number, date, snils, additional)
    return data, summary
