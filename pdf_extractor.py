"""
PDF Data Extractor - Offline version for Pyodide.
Extracts structured data from Appraisal and ISBS PDFs.
Uses pdfplumber table extraction for clean bank/FI name separation.
"""
import pdfplumber
import re
from typing import Dict, List, Optional


def format_number(value: str) -> str:
    if value is None:
        return ""
    try:
        clean = str(value).replace(",", "").replace(" ", "").strip()
        if not clean:
            return ""
        num = float(clean)
        if num == int(num):
            return f"{int(num):,}"
        return f"{num:,.2f}"
    except (ValueError, TypeError):
        return str(value)


def extract_text_from_pdf(pdf_path: str) -> List[str]:
    pages = []
    with pdfplumber.open(pdf_path) as pdf:
        for page in pdf.pages:
            text = page.extract_text() or ""
            pages.append(text)
    return pages


class AppraisalExtractor:
    def __init__(self, pdf_path: str):
        self.pdf_path = pdf_path
        self.pages_text = extract_text_from_pdf(pdf_path)
        self.full_text = "\n".join(self.pages_text)
        self.confidence = {}
        self.missing_fields = []

    def _record(self, field, value, conf):
        self.confidence[field] = conf
        if not value:
            self.missing_fields.append(field)
        return value

    def extract_business_name(self):
        p1 = self.pages_text[0] if self.pages_text else ""
        m = re.search(r"Business\s*Name:\s*(.+?)(?:\s+Business\s*Since|\n)", p1)
        if m:
            return self._record("business_name", m.group(1).strip(), 0.95)
        return self._record("business_name", None, 0.0)

    def extract_address(self):
        p1 = self.pages_text[0] if self.pages_text else ""
        m = re.search(r"Business\s*Address:\s*(.+?)(?:\n)", p1)
        if m:
            return self._record("address", m.group(1).strip(), 0.95)
        return self._record("address", None, 0.0)

    def extract_key_person(self):
        p1 = self.pages_text[0] if self.pages_text else ""
        m = re.search(r"Proprietor.*?/\s*(.+?)\s+Education:", p1)
        if m:
            name = m.group(1).strip()
            name = re.sub(r"Director(?:'s)?\s*Name:\s*", "", name).strip()
            return self._record("key_person", name, 0.95)
        m = re.search(r"Director(?:'s)?\s*Name:\s*\n?\s*(.+?)(?:\n|Education)", p1)
        if m:
            return self._record("key_person", m.group(1).strip(), 0.85)
        return self._record("key_person", None, 0.0)

    def extract_contact_number(self):
        p1 = self.pages_text[0] if self.pages_text else ""
        m = re.search(r"Phone:\s*(01\d{9,10})", p1)
        if m:
            return self._record("contact_number", m.group(1).strip(), 0.95)
        m = re.search(r"Business\s+(01\d{9,10})", p1)
        if m:
            return self._record("contact_number", m.group(1), 0.8)
        m = re.search(r"Business.*?(01\d{9,10})", p1, re.DOTALL)
        if m:
            return self._record("contact_number", m.group(1), 0.6)
        return self._record("contact_number", None, 0.0)

    def extract_ownership_type(self):
        p1 = self.pages_text[0] if self.pages_text else ""
        m = re.search(r"Legal\s*Status/?\s*Type\s+(.+?)\s+Organization", p1)
        if m:
            val = m.group(1).strip()
            val = re.sub(r"of\s*ownership:\s*", "", val).strip()
            return self._record("ownership_type", val, 0.95)
        m = re.search(r"Legal\s*Status.*?ownership:\s*(.+?)(?:\n|Organization|Wholesaler)", p1, re.DOTALL)
        if m:
            return self._record("ownership_type", m.group(1).strip(), 0.8)
        return self._record("ownership_type", None, 0.0)

    def extract_rm_name(self):
        for page in self.pages_text[:3]:
            m = re.search(r"Initiator.*?Code:\s*(.+?)(?:\s+Visit|\n)", page)
            if m:
                return self._record("rm_name", m.group(1).strip(), 0.95)
        return self._record("rm_name", None, 0.0)

    def extract_date(self):
        p1 = self.pages_text[0] if self.pages_text else ""
        # Look for Date: near CIF (General Information), not Report Generation Date
        m = re.search(r"CIF:.*?Date:\s*(\d{1,2}[-/]\w+[-/]\d{2,4})", p1, re.DOTALL)
        if m:
            return self._record("date", m.group(1).strip(), 0.95)
        # Fallback: find any date after avoiding Report Generation
        text_after_report = p1
        rg_idx = text_after_report.find("Report Generation Date")
        if rg_idx >= 0:
            text_after_report = text_after_report[rg_idx + 100:]
        m = re.search(r"\bDate:\s*(\d{1,2}[-/]\w+[-/]\d{2,4})", text_after_report)
        if m:
            return self._record("date", m.group(1).strip(), 0.9)
        return self._record("date", None, 0.0)

    def extract_background(self):
        bg_page = ""
        for p in self.pages_text:
            if "Background of the client" in p or "Dynamics of the business" in p:
                bg_page = p
                break
        if not bg_page:
            return self._record("background", None, 0.0)

        m = re.search(
            r"Background\s+of\s+the\s+client/?\s*\n?(.+?)(?=\n\s*\d+\s*\n?Factory|\n\s*Factory\s*\n?Premise|\nPremise Ownership|\nBusiness\s+Dynamic|\nBusiness\s+Details|\nRisk\s+Factors|\nShowrooms:|\nPrepared\s+BY)",
            bg_page, re.DOTALL | re.IGNORECASE
        )
        if m:
            text = m.group(1).strip()
            # Remove "Dynamics of the business:" label
            text = re.sub(r"^Dynamics\s+of\s+the\s+business:\s*", "", text, flags=re.IGNORECASE).strip()
            # Strip out "Person(s) Met:" line if present
            text = re.sub(r"Person\(s\)\s*Met:.*?\n", "", text).strip()
            # Remove trailing markers
            text = re.sub(r"\n\s*Showrooms:.*$", "", text, flags=re.DOTALL).strip()
            text = re.sub(r"\n\s*Premise Ownership.*$", "", text, flags=re.DOTALL).strip()
            text = re.sub(r"\n\s*\d+\s*$", "", text).strip()
            return self._record("background", text, 0.85)
        return self._record("background", None, 0.0)

    def extract_suppliers(self, background):
        if not background:
            return self._record("major_suppliers", None, 0.0)
        sentences = re.split(r'(?<=[.!?])\s+', background)
        sents = []
        for s in sentences:
            s = s.strip()
            if re.search(r'(?:import|procure|purchase|source|buy|bought)', s, re.IGNORECASE):
                sents.append(s)
        if sents:
            return self._record("major_suppliers", " ".join(sents), 0.8)
        return self._record("major_suppliers", None, 0.0)

    def extract_clients(self, background):
        if not background:
            return self._record("major_clients", None, 0.0)
        sentences = re.split(r'(?<=[.!?])\s+', background)
        client_sents = []
        for s in sentences:
            s = s.strip()
            has_sell_to = re.search(r'(?:sells?|supplies?|caters?|distribut|exports?)\s+(?:to|for|his\s+product|its\s+product|mainly)', s, re.IGNORECASE)
            has_buyer_mention = re.search(r'(?:wholesale\s+and\s+retail|wholesale.*?retail|retail.*?wholesale|customer|dealer|buyer)', s, re.IGNORECASE)
            is_generic = re.search(r'(?:opened\s+his\s+own|started|experience|expanded|sister\s+concern|imports?\s+through|conveniently\s+located)', s, re.IGNORECASE)
            if (has_sell_to or has_buyer_mention) and not is_generic:
                client_sents.append(s)

        if client_sents:
            return self._record("major_clients", " ".join(client_sents), 0.85)

        for s in sentences:
            s = s.strip()
            if re.search(r'sells\s+all\s+types|sells\s+.*?on\s+wholesale|exports?\s+its\s+products?|main\s+customer', s, re.IGNORECASE):
                return self._record("major_clients", s, 0.7)

        return self._record("major_clients", None, 0.0)

    def _extract_liability_tables(self):
        """Extract LT and ST liability data using pdfplumber TABLE extraction."""
        lt_results = []
        st_results = []

        with pdfplumber.open(self.pdf_path) as pdf:
            for page in pdf.pages:
                tables = page.extract_tables()
                for table in tables:
                    if not table or len(table) < 2:
                        continue
                    header = [str(c or "").replace("\n", " ").strip().lower() for c in table[0]]

                    has_cash_security = any("cash" in h and "security" in h for h in header)
                    has_cr_summation = any("summation" in h for h in header)
                    has_fi_col = any("financial" in h or "institution" in h for h in header)
                    has_facility_col = any("facility" in h for h in header)
                    has_installment = any("installment" in h for h in header)

                    is_st = has_fi_col and has_facility_col and has_cr_summation
                    is_lt = has_fi_col and has_facility_col and has_installment and not has_cr_summation

                    if not is_lt and not is_st:
                        continue

                    def find_col(keywords):
                        for ki, h in enumerate(header):
                            if any(kw in h for kw in keywords):
                                return ki
                        return -1

                    col_fi = find_col(["financial", "institution"])
                    col_facility = find_col(["facility"])
                    col_limit = find_col(["sanction", "limit"])
                    col_outstanding = find_col(["current", "outstandin"])
                    col_term = find_col(["term"])
                    col_emi = find_col(["installment\n(bdt)", "installment (bdt)", "(bdt)"])
                    col_recycle = find_col(["recycle"])

                    for row in table[1:]:
                        if not row or len(row) < 4:
                            continue
                        cells = [str(c or "").replace("\n", " ").strip() for c in row]

                        if cells[0].lower().startswith("security"):
                            continue

                        fi_name = cells[col_fi] if 0 <= col_fi < len(cells) else ""
                        facility = cells[col_facility] if 0 <= col_facility < len(cells) else ""
                        limit_val = cells[col_limit] if 0 <= col_limit < len(cells) else ""

                        if not limit_val or not re.search(r"[\d,]+", limit_val):
                            continue

                        if is_lt and not is_st:
                            outstanding = cells[col_outstanding] if 0 <= col_outstanding < len(cells) else ""
                            tenure = cells[col_term] if 0 <= col_term < len(cells) else ""
                            emi = cells[col_emi] if 0 <= col_emi < len(cells) else ""
                            lt_results.append({
                                "bank_name": fi_name,
                                "facility_type": facility.replace("\n", " "),
                                "limit": format_number(limit_val),
                                "outstanding": format_number(outstanding),
                                "emi": format_number(emi),
                                "tenure": tenure,
                                "repayment_status": "Regular"
                            })
                        elif is_st:
                            outstanding = cells[col_outstanding] if 0 <= col_outstanding < len(cells) else ""
                            recycle = cells[col_recycle] if 0 <= col_recycle < len(cells) else ""
                            st_results.append({
                                "bank_name": fi_name,
                                "facility_type": facility.replace("\n", " "),
                                "limit": format_number(limit_val),
                                "outstanding": format_number(outstanding),
                                "recycle_times": recycle,
                                "expiry_date": "",
                                "repayment_status": "Regular"
                            })

        return lt_results, st_results

    def extract_long_term_liabilities(self):
        if not hasattr(self, '_lt_st_cache'):
            self._lt_st_cache = self._extract_liability_tables()
        results = self._lt_st_cache[0]
        self.confidence["long_term_liabilities"] = 0.85 if results else 0.0
        if not results:
            self.missing_fields.append("long_term_liabilities")
        return results

    def extract_short_term_liabilities(self):
        if not hasattr(self, '_lt_st_cache'):
            self._lt_st_cache = self._extract_liability_tables()
        results = list(self._lt_st_cache[1])

        # Text fallback for ST entries that spill across pages (not in tables)
        existing_keys = {(e["limit"].replace(",",""), e["outstanding"].replace(",","")) for e in results}
        facility_types = ["CC-Hypo", "Overdraft", "LTR", "SOD", "Letter of Credit",
                          "Mudaraba", "Bi-Muraba", "Other direct", "Demand Loan"]
        for page in self.pages_text:
            for ftype in facility_types:
                pattern = re.compile(
                    re.escape(ftype).replace(r"\ ", r"\s*") + r"[a-z\s]*\s+([\d,]+)\s+([\d,]+)",
                    re.IGNORECASE
                )
                for m in pattern.finditer(page):
                    limit_clean = m.group(1).replace(",", "")
                    out_clean = m.group(2).replace(",", "")
                    if (limit_clean, out_clean) in existing_keys:
                        continue
                    # Find bank name in preceding text
                    before = page[:m.start()]
                    bank = "Unknown"
                    bank_matches = re.findall(
                        r"([\w\s]+?(?:Bank|Finance|Islami Bank)[\w\s]*?(?:Limited)?)",
                        before[-200:], re.IGNORECASE
                    )
                    if bank_matches:
                        bank = re.sub(r"\s+", " ", bank_matches[-1]).strip()
                        parts = bank.split()
                        for bi in range(len(parts)):
                            if parts[bi] in ("Bank", "Finance", "Islami", "Limited"):
                                start = max(0, bi - 3)
                                bank = " ".join(parts[start:])
                                break
                    ftype_display = re.sub(r"\s+", " ", page[m.start():m.start()+40].split(str(m.group(1)))[0]).strip()
                    results.append({
                        "bank_name": bank,
                        "facility_type": ftype_display,
                        "limit": format_number(m.group(1)),
                        "outstanding": format_number(m.group(2)),
                        "recycle_times": "",
                        "expiry_date": "",
                        "repayment_status": "Regular"
                    })
                    existing_keys.add((limit_clean, out_clean))

        self.confidence["short_term_liabilities"] = 0.85 if results else 0.0
        if not results:
            self.missing_fields.append("short_term_liabilities")
        return results

    def extract_import_export(self):
        for p in self.pages_text:
            if re.search(r"Import\s*/?\s*Export\s*Performance", p, re.IGNORECASE):
                m = re.search(r"Import\s*/?\s*Export\s*Performance.*?Remarks:\s*(.*?)(?=\n\s*\d+\.\s)", p, re.DOTALL | re.IGNORECASE)
                if m and m.group(1).strip():
                    return self._record("import_export", m.group(1).strip(), 0.8)
                return self._record("import_export", "N/A", 0.9)
        return self._record("import_export", "N/A", 0.5)

    def extract_credit_summation(self):
        for p in self.pages_text:
            m = re.search(r"Total\s+Credit\s*\n?\s*([\d,]+)\s*\n?\s*Summation", p)
            if m:
                return self._record("credit_summation", format_number(m.group(1)), 0.95)
            m = re.search(r"Total\s+Credit\s+Summation:\s*([\d,]+)", p)
            if m:
                return self._record("credit_summation", format_number(m.group(1)), 0.95)
        return self._record("credit_summation", None, 0.0)

    def extract_all(self):
        business_name = self.extract_business_name()
        address = self.extract_address()
        key_person = self.extract_key_person()
        contact = self.extract_contact_number()
        ownership = self.extract_ownership_type()
        rm_name = self.extract_rm_name()
        date_val = self.extract_date()
        background = self.extract_background()
        suppliers = self.extract_suppliers(background)
        clients = self.extract_clients(background)
        lt = self.extract_long_term_liabilities()
        st = self.extract_short_term_liabilities()
        imp_exp = self.extract_import_export()
        credit_sum = self.extract_credit_summation()
        return {
            "business_name": business_name, "address": address, "key_person": key_person,
            "contact_number": contact, "ownership_type": ownership, "rm_name": rm_name,
            "rm_contact": "", "background": background, "major_suppliers": suppliers,
            "major_clients": clients, "date": date_val, "long_term_liabilities": lt,
            "short_term_liabilities": st, "import_export": imp_exp,
            "credit_summation": credit_sum, "confidence": self.confidence,
            "missing_fields": self.missing_fields,
        }


class ISBSExtractor:
    def __init__(self, pdf_path: str):
        self.pdf_path = pdf_path
        self.pages_text = extract_text_from_pdf(pdf_path)
        self.full_text = "\n".join(self.pages_text)
        self.confidence = {}
        self.missing_fields = []

    def _record(self, field, value, conf):
        self.confidence[field] = conf
        if not value:
            self.missing_fields.append(field)
        return value

    def extract_all(self):
        fp_text = ""
        for page in self.pages_text:
            if "Financial Performance" in page:
                fp_text = page
                break
        if not fp_text:
            fp_text = self.full_text

        result = {}
        for field, pattern in [
            ("sales_revenue", r"Sales/Revenue\s+([\d,.]+)"),
            ("inventory", r"\bInventory\s+([\d,.]+)"),
            ("ar", r"A/C\s*Receivable\s+([\d,.]+)"),
            ("ap", r"A/C\s*Payable\s+([\d,.]+)"),
        ]:
            m = re.search(pattern, fp_text)
            if m:
                result[field] = self._record(field, format_number(m.group(1)), 0.95)

        m = re.search(r"Gross\s+Margin\s+\(%\)\s+([\d.]+)", fp_text)
        if m:
            result["gp_pct"] = self._record("gp_pct", f"{m.group(1)}%", 0.95)

        m = re.search(r"Net\s+Profit\s+Margin\s+\(After\s+Tax\)\s+([\d.]+%?)", fp_text)
        if m:
            val = m.group(1).replace("%", "")
            result["np_pct"] = self._record("np_pct", f"{val}%", 0.95)

        return {**result, "confidence": self.confidence, "missing_fields": self.missing_fields}


def extract_all_data(appraisal_path: str, isbs_path: str) -> Dict:
    a = AppraisalExtractor(appraisal_path).extract_all()
    i = ISBSExtractor(isbs_path).extract_all()
    all_missing = a.get("missing_fields", []) + i.get("missing_fields", [])
    return {
        "business_name": a.get("business_name", ""), "address": a.get("address", ""),
        "key_person": a.get("key_person", ""), "contact_number": a.get("contact_number", ""),
        "ownership_type": a.get("ownership_type", ""), "rm_name": a.get("rm_name", ""),
        "rm_contact": "", "background": a.get("background", ""),
        "major_suppliers": a.get("major_suppliers", ""), "major_clients": a.get("major_clients", ""),
        "date": a.get("date", ""), "sales_revenue": i.get("sales_revenue", ""),
        "inventory": i.get("inventory", ""), "ar": i.get("ar", ""),
        "ap": i.get("ap", ""), "gp_pct": i.get("gp_pct", ""),
        "np_pct": i.get("np_pct", ""),
        "long_term_liabilities": a.get("long_term_liabilities", []),
        "short_term_liabilities": a.get("short_term_liabilities", []),
        "import_export": a.get("import_export", "N/A"),
        "credit_summation": a.get("credit_summation", ""),
        "confidence": {**a.get("confidence", {}), **i.get("confidence", {})},
        "missing_fields": all_missing,
    }
