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
        passport = extract_passport(consent)
        if passport:
            found.append(
                f"Паспорт: {passport['series']} № {passport['number']}, "
                f"выдан {passport['date_display']} ({passport['issuer']})"
            )
        address = re.search(r"зарегистрированный\s+по\s+адресу:\s*(.+?)\s*,\s*даю", consent, re.I)
        if address:
            found.append(f"Адрес регистрации: {address.group(1).strip()}")
    phone = re.search(r"([+]?[78][\d\s()\-]{9,18})\s+контактный\s+телефон", application, re.I)
    if phone:
        phone_value = "".join(phone.group(1).split())
        found.append(f"Телефон: {phone_value}")
    parcel_address = extract_parcel_address(contract)
    if parcel_address:
        found.append(f"Адрес участка: {parcel_address}")
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


def extract_passport(consent: str | None) -> dict[str, str] | None:
    """Extract a complete Russian passport record from the consent document."""
    if not consent:
        return None
    normalized = re.sub(r"П\s+а(спорт)", r"Паспорт", consent, flags=re.I)
    normalized = re.sub(r"В\s+ыдан", "Выдан", normalized, flags=re.I)
    match = re.search(
        r"паспорт\s+(\d{4})\s*№?\s*(\d{6})\s+выдан\s+(.+?),\s*(\d{2}\.\d{2}\.\d{4})",
        normalized, re.I,
    )
    if not match:
        return None
    series, number, issuer, date_display = (part.strip() for part in match.groups())
    date_iso = f"{date_display[6:10]}-{date_display[3:5]}-{date_display[:2]}"
    return {
        "series": series,
        "number": number,
        "issuer": issuer,
        "date_display": date_display,
        "date_iso": date_iso,
    }


def extract_parcel_address(contract: str | None) -> str | None:
    if not contract:
        return None
    match = re.search(r"расположен\s+по\s+адресу:\s*(.+?)\s*\.\s*1\.3", contract, re.I)
    if not match:
        return None
    return match.group(1).strip().rstrip(" .")


def _load_fields(documents: list[Path]) -> tuple[tuple[str, str, str, str, str, str | None], list[str]]:
    application, contract, consent = identify_documents(documents)
    return extract_fields(application, contract, consent), extract_additional_fields(application, contract, consent)


def _summary(family: str, first: str, patronymic: str, number: str, date: str, snils: str | None, passport: dict[str, str] | None, parcel_address: str | None, additional: list[str]) -> str:
    identifier = "СНИЛС извлечён и записан" if snils else (
        "СНИЛС не найден; паспорт извлечён и записан" if passport else "СНИЛС и паспорт не найдены"
    )
    text = f"{family} {first} {patronymic}; договор № {number} от {date}; {identifier}"
    if parcel_address:
        text += "; адрес участка из договора добавлен в заключение"
    additional = [
        item for item in additional
        if not item.startswith(("Паспорт:", "Адрес участка:"))
    ]
    if additional:
        text += "\n\nИзвлечено для проверки, но не записано в XML (в шаблоне нет подтверждённых для этого полей):\n• " + "\n• ".join(additional)
    return text


def _set_client(
    root: ET.Element,
    family: str,
    first: str,
    patronymic: str,
    snils: str | None,
    passport: dict[str, str] | None,
) -> None:
    client = root.find(".//Clients/Client/Person")
    if client is None:
        raise ValueError("В шаблоне не найден узел клиента Person")
    for tag, value in (("FamilyName", family), ("FirstName", first), ("Patronymic", patronymic)):
        node = client.find(tag)
        if node is None:
            raise ValueError(f"В шаблоне не найдено поле {tag}")
        node.text = value
    snils_node = client.find("SNILS")
    # Keep exactly one identification route: prefer SNILS whenever it was found.
    if snils:
        if snils_node is None:
            snils_node = ET.SubElement(client, "SNILS")
        snils_node.text = snils
        for document in client.findall("Document"):
            client.remove(document)
    else:
        if snils_node is not None:
            client.remove(snils_node)
        for document in client.findall("Document"):
            client.remove(document)
        if passport:
            document = ET.SubElement(client, "Document")
            for tag, value in (
                ("CodeDocument", "008001001000"),
                ("Name", "Паспорт гражданина Российской Федерации"),
                ("Series", passport["series"]),
                ("Number", passport["number"]),
                ("Date", passport["date_iso"]),
                ("IssueOrgan", passport["issuer"]),
            ):
                ET.SubElement(document, tag).text = value


def _set_agreement(root: ET.Element, number: str, date: str) -> None:
    agreements = root.findall(".//AgreementCadWork")
    if not agreements:
        raise ValueError("В шаблоне не найден узел AgreementCadWork")
    for agreement in agreements:
        for tag, value in (("NumberAgreement", number), ("DateAgreement", date)):
            node = agreement.find(tag)
            if node is not None:
                node.text = value


def _set_reason(root: ET.Element) -> None:
    node = root.find("./GeneralCadastralWorks/Reason")
    if node is None:
        raise ValueError("В шаблоне не найдено поле Reason")
    node.text = (
        "уточнением описания местоположения границ и (или) площади "
        "земельного участка с кадастровым номером"
    )


def build_application_only_xml(documents: list[Path], template_xml: Path) -> tuple[bytes, str]:
    """Fill DOCX-derived fields into the shared, case-neutral XML scaffold."""
    (family, first, patronymic, number, date, snils), additional = _load_fields(documents)
    _application_text, contract_text, consent_text = identify_documents(documents)
    tree = ET.parse(template_xml)
    root = tree.getroot()

    root.attrib["GUID"] = str(uuid.uuid4())

    passport = extract_passport(consent_text)
    _set_client(root, family, first, patronymic, snils, passport)
    _set_agreement(root, number, date)
    _set_reason(root)

    conclusion_node = root.find("Conclusion")
    if conclusion_node is None:
        raise ValueError("В шаблоне не найден узел Conclusion")
    conclusion_template_path = template_xml.with_name("conclusion_template.txt")
    if not conclusion_template_path.is_file():
        raise ValueError("Рядом с программой не найден файл conclusion_template.txt")
    conclusion_template = conclusion_template_path.read_text(encoding="utf-8")
    parcel_address = extract_parcel_address(contract_text)
    if parcel_address:
        address_line = parcel_address
    else:
        address_line = "_"
    conclusion_node.text = conclusion_template.replace("{{PARCEL_ADDRESS}}", address_line).strip()

    data = ET.tostring(root, encoding="utf-8", xml_declaration=True)
    summary = _summary(family, first, patronymic, number, date, snils, passport, parcel_address, additional)
    return data, summary
