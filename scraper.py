"""
KursKompass API – QIS/LSF Scraper für Goethe-Universität Frankfurt
Backend-API für das statische Frontend.
"""

import requests
from bs4 import BeautifulSoup
import time
import re
import json
import logging
from dataclasses import dataclass, field, asdict
from typing import Optional
from urllib.parse import urljoin, unquote

logging.basicConfig(level=logging.INFO, format='%(asctime)s - %(levelname)s - %(message)s')
logger = logging.getLogger(__name__)

BASE_URL = "https://qis.server.uni-frankfurt.de"
START_ROOT = "118146%7C118447"
REQUEST_DELAY = 1.2


@dataclass
class Veranstaltung:
    pfad: str = ""
    kennung: str = ""
    titel: str = ""
    dozent: str = ""
    veranstaltungsart: str = ""
    semester: str = ""
    sws: str = ""
    gruppe: str = ""
    tag: str = ""
    zeit: str = ""
    rhythmus: str = ""
    raum: str = ""
    max_teilnehmer: str = ""
    belegung: str = ""
    belegungsfristen: str = ""
    credits: str = ""
    sprache: str = ""
    kuerzel: str = ""
    studiengaenge: str = ""
    kommentar: str = ""
    voraussetzungen: str = ""
    detail_url: str = ""
    weitere_termine: list = field(default_factory=list)


@dataclass
class BaumKnoten:
    name: str
    root_path: str
    url: str
    children: list = field(default_factory=list)
    has_veranstaltungen: bool = False


class QISScraper:

    def __init__(self):
        self.session = requests.Session()
        self.session.headers.update({
            "User-Agent": "Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_7) AppleWebKit/537.36",
            "Accept": "text/html,application/xhtml+xml",
            "Accept-Language": "de-DE,de;q=0.9",
        })
        self.veranstaltungen = []
        self.progress = {
            "phase": "idle",
            "status": "Bereit",
            "current": 0,
            "total": 0,
            "details": []
        }

    def _get_page(self, url):
        try:
            time.sleep(REQUEST_DELAY)
            response = self.session.get(url, timeout=30)
            response.raise_for_status()
            response.encoding = "utf-8"
            return BeautifulSoup(response.text, "lxml")
        except Exception as e:
            logger.error(f"Fehler beim Abrufen von {url}: {e}")
            return None

    def _build_tree_url(self, root_path):
        return (
            f"{BASE_URL}/qisserver/rds"
            f"?state=wtree&search=1&trex=step"
            f"&root120261={root_path}&P.vx=kurz"
        )

    # === TOP-LEVEL: Quick scan for Lehramtstypen ===

    def scan_top_level(self):
        """Scannt nur die erste Ebene (L1, L2, L3, L5 etc.)."""
        url = self._build_tree_url(START_ROOT)
        soup = self._get_page(url)
        if not soup:
            return []
        nodes = self._find_tree_children(soup, START_ROOT)
        return [{"name": n.name, "root_path": n.root_path} for n in nodes]

    # === PHASE 1: Tree Scan ===

    def scan_tree(self, start_root=None):
        self.progress = {"phase": "scan", "status": "Scanne Baumstruktur...", "current": 0, "total": 0, "details": []}

        root = start_root or START_ROOT
        url = self._build_tree_url(root)
        soup = self._get_page(url)
        if not soup:
            self.progress["status"] = "Fehler beim Laden der Startseite"
            self.progress["phase"] = "error"
            return []

        top_nodes = self._find_tree_children(soup, root)

        # If no children found, this node itself might have courses
        if not top_nodes and start_root:
            # Create a single node for this root
            node = BaumKnoten(name="Ausgewählter Bereich", root_path=root, url=url)
            ver_table = soup.find("table", summary="Übersicht über alle Veranstaltungen")
            if ver_table:
                node.has_veranstaltungen = True
            top_nodes = [node]

        for node in top_nodes:
            logger.info(f"Scanne: {node.name}")
            self.progress["status"] = f"Scanne {node.name}..."
            self._scan_node_recursive(node, 0, 6)

        self.progress["phase"] = "scan_done"
        self.progress["status"] = f"Struktur geladen: {self._count_nodes(top_nodes)} Bereiche"
        return top_nodes

    def _scan_node_recursive(self, node, depth, max_depth):
        if depth >= max_depth:
            return

        soup = self._get_page(node.url)
        if not soup:
            return

        self.progress["current"] += 1
        self.progress["details"].append(node.name)
        if len(self.progress["details"]) > 5:
            self.progress["details"] = self.progress["details"][-5:]

        ver_table = soup.find("table", summary="Übersicht über alle Veranstaltungen")
        if ver_table:
            node.has_veranstaltungen = True

        children = self._find_tree_children(soup, node.root_path)
        node.children = children

        for child in children:
            self._scan_node_recursive(child, depth + 1, max_depth)

    def _find_tree_children(self, soup, parent_root):
        children = []
        parent_decoded = unquote(parent_root)
        parent_segments = parent_decoded.split("|")
        parent_depth = len(parent_segments)

        for a_tag in soup.find_all("a", class_="ueb"):
            href = a_tag.get("href", "")
            if "state=wtree" not in href or "root120261=" not in href:
                continue

            match = re.search(r'root120261=([^&]+)', href)
            if not match:
                continue

            root_path = match.group(1)
            root_decoded = unquote(root_path)
            segments = root_decoded.split("|")

            if len(segments) == parent_depth + 1 and root_decoded.startswith(parent_decoded):
                name = a_tag.get_text(strip=True)
                if name and name not in ["kurz", "mittel", "lang"]:
                    children.append(BaumKnoten(
                        name=name,
                        root_path=root_path,
                        url=self._build_tree_url(root_path)
                    ))

        return children

    def _count_nodes(self, nodes):
        count = len(nodes)
        for node in nodes:
            count += self._count_nodes(node.children)
        return count

    # === PHASE 2: Scrape Selected ===

    def scrape_selected(self, tree, selected_paths):
        self.veranstaltungen = []
        self.progress = {"phase": "scrape", "status": "Starte...", "current": 0, "total": len(selected_paths), "details": []}

        for node in tree:
            self._scrape_node_recursive(node, selected_paths, [])

        self.progress["phase"] = "done"
        self.progress["status"] = f"Fertig! {len(self.veranstaltungen)} Veranstaltungen gefunden."
        return self.veranstaltungen

    def _scrape_node_recursive(self, node, selected, path):
        current_path = path + [node.name]

        if node.root_path in selected:
            logger.info(f"Scrape: {' > '.join(current_path)}")
            self.progress["status"] = f"Scrape: {node.name}..."
            self._scrape_page_veranstaltungen(node.url, current_path)
            self.progress["current"] += 1

        for child in node.children:
            self._scrape_node_recursive(child, selected, current_path)

    def _scrape_page_veranstaltungen(self, url, path):
        soup = self._get_page(url)
        if not soup:
            return

        table = soup.find("table", summary="Übersicht über alle Veranstaltungen")
        if not table:
            return

        rows = table.find_all("tr")
        for row in rows:
            cells = row.find_all("td")
            if len(cells) < 2:
                continue

            first_cell = cells[0]
            link = first_cell.find("a", class_="regular")
            if not link or "state=verpublish" not in link.get("href", ""):
                continue

            link_text = link.get_text(strip=True)
            detail_url = urljoin(BASE_URL, link["href"])

            dozent_links = first_cell.find_all("a", class_="klein")
            dozent_parts = []
            for dl in dozent_links:
                dozent_parts.append(dl.get_text(" ", strip=True))
            dozent = ", ".join(dozent_parts)

            vst_art = cells[1].get_text(strip=True) if len(cells) > 1 else ""

            kennung = ""
            titel = link_text
            if ":" in link_text:
                parts = link_text.split(":", 1)
                kennung = parts[0].strip()
                titel = parts[1].strip()

            self.progress["details"].append(f"{titel[:50]}...")
            if len(self.progress["details"]) > 5:
                self.progress["details"] = self.progress["details"][-5:]

            detail = self._scrape_detail(detail_url)

            # Modul-Fix: Wenn kennung leer, aus Pfad den letzten Bereich nehmen
            if not kennung and path:
                kennung = path[-1] if len(path) > 0 else ""

            base_dozent = dozent or (detail.get("dozent", "") if detail else "")
            gruppen = detail.get("gruppen", []) if detail else []

            # Wenn Gruppen vorhanden: Pro Gruppe einen Eintrag
            if gruppen and len(gruppen) > 1:
                # Prüfe ob es echte Gruppen sind (verschiedene Gruppennamen)
                gruppe_namen = set(g.get("gruppe", "") for g in gruppen)
                has_real_groups = any(g for g in gruppe_namen if g)

                if has_real_groups:
                    for g in gruppen:
                        v = Veranstaltung(
                            pfad=" > ".join(path),
                            kennung=kennung,
                            titel=titel,
                            dozent=g.get("dozent", "") or base_dozent,
                            veranstaltungsart=vst_art,
                            semester=detail.get("semester", "") if detail else "",
                            sws=detail.get("sws", "") if detail else "",
                            gruppe=g.get("gruppe", ""),
                            tag=g.get("tag", ""),
                            zeit=g.get("zeit", ""),
                            rhythmus=g.get("rhythmus", ""),
                            raum=g.get("raum", ""),
                            max_teilnehmer=detail.get("max_teilnehmer", "") if detail else "",
                            belegung=detail.get("belegung", "") if detail else "",
                            belegungsfristen=detail.get("belegungsfristen", "") if detail else "",
                            credits=detail.get("credits", "") if detail else "",
                            sprache=detail.get("sprache", "") if detail else "",
                            kuerzel=detail.get("kuerzel", "") if detail else "",
                            studiengaenge=detail.get("studiengaenge", "") if detail else "",
                            kommentar=detail.get("kommentar", "") if detail else "",
                            voraussetzungen=detail.get("voraussetzungen", "") if detail else "",
                            detail_url=detail_url,
                        )
                        self.veranstaltungen.append(v)
                    continue

            # Normaler Eintrag (keine Gruppen oder nur eine Gruppe)
            v = Veranstaltung(
                pfad=" > ".join(path),
                kennung=kennung,
                titel=titel,
                dozent=base_dozent,
                veranstaltungsart=vst_art,
                semester=detail.get("semester", "") if detail else "",
                sws=detail.get("sws", "") if detail else "",
                gruppe=gruppen[0].get("gruppe", "") if gruppen else "",
                tag=detail.get("tag", "") if detail else "",
                zeit=detail.get("zeit", "") if detail else "",
                rhythmus=detail.get("rhythmus", "") if detail else "",
                raum=detail.get("raum", "") if detail else "",
                max_teilnehmer=detail.get("max_teilnehmer", "") if detail else "",
                belegung=detail.get("belegung", "") if detail else "",
                belegungsfristen=detail.get("belegungsfristen", "") if detail else "",
                credits=detail.get("credits", "") if detail else "",
                sprache=detail.get("sprache", "") if detail else "",
                kuerzel=detail.get("kuerzel", "") if detail else "",
                studiengaenge=detail.get("studiengaenge", "") if detail else "",
                kommentar=detail.get("kommentar", "") if detail else "",
                voraussetzungen=detail.get("voraussetzungen", "") if detail else "",
                detail_url=detail_url,
                weitere_termine=gruppen[1:] if len(gruppen) > 1 else [],
            )
            self.veranstaltungen.append(v)

    def _scrape_detail(self, url):
        soup = self._get_page(url)
        if not soup:
            return None

        result = {}

        # GRUNDDATEN
        grunddaten = soup.find("table", summary="Grunddaten zur Veranstaltung")
        if grunddaten:
            for th in grunddaten.find_all("th", class_="mod"):
                label = th.get_text(strip=True)
                td = th.find_next_sibling("td")
                if not td:
                    continue
                value = td.get_text(strip=True)

                if "Veranstaltungsart" in label:
                    result["veranstaltungsart"] = value
                elif "Kürzel" in label:
                    result["kuerzel"] = value
                elif "Semester" in label:
                    result["semester"] = value
                elif label == "SWS":
                    result["sws"] = value
                elif "Max. Teilnehmer" in label:
                    result["max_teilnehmer"] = value
                elif "Sprache" in label:
                    result["sprache"] = value
                elif "Credits" in label:
                    result["credits"] = value
                elif "Belegung" == label:
                    result["belegung"] = value

            fristen = []
            for td in grunddaten.find_all("td", headers="basic_14"):
                frist_text = td.get_text(strip=True)
                if frist_text:
                    fristen.append(frist_text)
            result["belegungsfristen"] = " | ".join(fristen)

        # TERMINE – mit Gruppen-Erkennung
        gruppen = []

        # Suche nach Gruppen-Überschriften ("Termine Gruppe: Gruppe 1")
        gruppe_headers = soup.find_all(string=re.compile(r"Termine\s+Gruppe.*Gruppe\s*\d+"))
        if not gruppe_headers:
            # Auch nach Überschriften-Tags suchen
            for tag in soup.find_all(["h2", "h3", "caption", "b", "strong"]):
                text = tag.get_text(strip=True)
                if re.search(r"Gruppe\s*\d+", text) and "Termin" in text:
                    gruppe_headers.append(tag)

        if gruppe_headers:
            # Gruppen-Modus: Jede Gruppe hat eigene Termine
            for gh in gruppe_headers:
                gruppe_name_match = re.search(r"Gruppe\s*(\d+)", str(gh) if isinstance(gh, str) else gh.get_text())
                gruppe_name = f"Gruppe {gruppe_name_match.group(1)}" if gruppe_name_match else ""

                # Finde die nächste Tabelle nach diesem Header
                parent = gh.parent if not isinstance(gh, str) else gh.parent
                next_table = None
                for sibling in parent.find_all_next():
                    if sibling.name == "table" and sibling.find("th"):
                        # Prüfe ob es eine Termine-Tabelle ist (hat Tag/Zeit Spalten)
                        headers_text = " ".join(th.get_text(strip=True) for th in sibling.find_all("th"))
                        if "Tag" in headers_text and "Zeit" in headers_text:
                            next_table = sibling
                            break
                    # Stop wenn nächste Gruppe kommt
                    if sibling.name in ["h2", "h3"] and "Gruppe" in sibling.get_text():
                        break

                if next_table:
                    termine = self._parse_termine_table(next_table)
                    if termine:
                        # Dozent aus der Gruppen-Tabelle extrahieren
                        gruppe_dozent = ""
                        for row in next_table.find_all("tr")[1:]:
                            cells = row.find_all("td")
                            for cell in cells:
                                # Lehrperson-Spalte finden
                                links = cell.find_all("a")
                                for link in links:
                                    href = link.get("href", "")
                                    if "personal" in href:
                                        gruppe_dozent = link.get_text(strip=True)
                                        break

                        for t in termine:
                            t["gruppe"] = gruppe_name
                            if gruppe_dozent and not t.get("dozent"):
                                t["dozent"] = gruppe_dozent
                        gruppen.extend(termine)
        else:
            # Kein Gruppen-Modus: Normale Termine-Tabelle
            termine_table = soup.find("table", summary="Übersicht über alle Veranstaltungstermine")
            if termine_table:
                termine = self._parse_termine_table(termine_table)
                gruppen.extend(termine)

        result["gruppen"] = gruppen

        # Rückwärtskompatibilität: Erste Gruppe als Haupttermin
        if gruppen:
            result["tag"] = gruppen[0].get("tag", "")
            result["zeit"] = gruppen[0].get("zeit", "")
            result["rhythmus"] = gruppen[0].get("rhythmus", "")
            result["raum"] = gruppen[0].get("raum", "")

        # DOZENTEN
        dozenten_table = soup.find("table", summary="Verantwortliche Dozenten")
        if not dozenten_table:
            dozenten_table = soup.find("table", summary="Zugeordnete Personen")
        if dozenten_table:
            dozenten = []
            for row in dozenten_table.find_all("tr")[1:]:
                td = row.find("td")
                if td:
                    dozent_text = td.get_text(strip=True)
                    if dozent_text:
                        dozenten.append(dozent_text)
            if dozenten:
                result["dozent"] = "; ".join(dozenten)

        # STUDIENGÄNGE
        stg_table = soup.find("table", summary="Übersicht über die zugehörigen Studiengänge")
        if stg_table:
            stg_list = []
            for row in stg_table.find_all("tr")[1:]:
                cells = row.find_all("td")
                if len(cells) >= 2:
                    abschluss = cells[0].get_text(strip=True)
                    stg = cells[1].get_text(strip=True)
                    stg_list.append(f"{abschluss}: {stg}")
            result["studiengaenge"] = "; ".join(stg_list)

        # INHALT
        inhalt_table = soup.find("table", summary="Weitere Angaben zur Veranstaltung")
        if inhalt_table:
            for th in inhalt_table.find_all("th"):
                label = th.get_text(strip=True)
                td = th.find_next_sibling("td")
                if not td:
                    continue
                text = td.get_text(strip=True)[:500]
                if "Kommentar" in label:
                    result["kommentar"] = text
                elif "Voraussetzungen" in label:
                    result["voraussetzungen"] = text

        return result

    def _parse_termine_table(self, table):
        """Parst eine Termine-Tabelle und gibt Liste von Terminen zurück."""
        termine = []
        rows = table.find_all("tr")
        for row in rows[1:]:
            cells = row.find_all("td")
            if len(cells) < 3:
                continue

            termin = {}
            cell_texts = [c.get_text(strip=True) for c in cells]

            for i, text in enumerate(cell_texts):
                if re.match(r'^(Mo|Di|Mi|Do|Fr|Sa|So)\.?$', text):
                    termin["tag"] = text
                    if i + 1 < len(cell_texts):
                        termin["zeit"] = cell_texts[i + 1].replace('\xa0', ' ')
                    if i + 2 < len(cell_texts):
                        termin["rhythmus"] = cell_texts[i + 2]
                    break

            for cell in cells:
                raum_link = cell.find("a", title=re.compile(r"Details ansehen zu Raum"))
                if not raum_link:
                    raum_link = cell.find("a", href=re.compile(r"raum"))
                if raum_link:
                    termin["raum"] = raum_link.get_text(strip=True)
                    break

            if termin.get("tag"):
                termine.append(termin)

        return termine


def tree_to_dict(nodes):
    return [{
        "name": n.name,
        "root_path": n.root_path,
        "has_veranstaltungen": n.has_veranstaltungen,
        "children": tree_to_dict(n.children)
    } for n in nodes]


def dict_to_tree(data):
    result = []
    for d in data:
        node = BaumKnoten(
            name=d["name"],
            root_path=d["root_path"],
            url=f"{BASE_URL}/qisserver/rds?state=wtree&search=1&trex=step&root120261={d['root_path']}&P.vx=kurz",
            has_veranstaltungen=d.get("has_veranstaltungen", False),
            children=dict_to_tree(d.get("children", []))
        )
        result.append(node)
    return result
