import re
import pandas as pd

DRUG_SOURCE_MAP = {
    'W': 'Single Source Brand (W)',
    'X': 'Multi Source Brand (X)',
    'Y': 'Generic (Y)',
}

# Categories where Qty Updated tracking applies
QTY_APPLICABLE = {'Claims Substitution', 'NDC Substituted (Drug Alt Service)'}
# Obsolete Drug / Substituted also applies — handled inline


def classify_row(log_text, status_id):
    log_text = str(log_text) if log_text is not None else ''
    status_id = str(status_id).strip() if status_id is not None else ''
    qty_flag = bool(re.search(r'Qty has been changed from:', log_text))

    # Priority 1 — Exclusion List
    if re.search(r'GCN in Exclusion List (25200|94200)', log_text):
        return 'Exclusion List', None, False

    # Priority 2 — Obsolete Drug  (A and (B or C))
    if re.search(r'Requested NDC: \S+ is Obsolete', log_text):
        if re.search(r'Substituted with Alternate NDC', log_text):      # B
            return 'Obsolete Drug', 'Substituted', qty_flag
        if re.search(r'No drug Alternatives were found', log_text) or status_id == '1302':  # C
            return 'Obsolete Drug', 'Not Substituted', False
        return 'Obsolete Drug', None, False

    # Priority 3 — Claims Substitution
    if re.search(r'has been switched to CLAIM NDC:', log_text):
        return 'Claims Substitution', None, qty_flag

    # Priority 4 — NDC Substituted from Drug Alternate Service
    m = re.search(
        r'Substituted with Alternate NDC: \S+ For DAW Code: (\d) Drug Source: (\w)',
        log_text,
    )
    if m:
        daw = m.group(1)
        src = m.group(2).upper()
        src_label = DRUG_SOURCE_MAP.get(src, src)
        sub = f'DAW {daw} - {src_label}'
        return 'NDC Substituted (Drug Alt Service)', sub, qty_flag

    return 'Unclassified', None, False


def classify_dataframe(df: pd.DataFrame) -> pd.DataFrame:
    df = df.copy()
    results = df.apply(
        lambda r: classify_row(r.get('LogTXT', ''), r.get('StatusID', '')),
        axis=1,
        result_type='expand',
    )
    results.columns = ['Category', 'Sub Category', 'Qty Updated']
    return pd.concat([df, results], axis=1)
