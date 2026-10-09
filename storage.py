import json, os
from datetime import datetime, timezone

DEFAULT_STATE = {
    'visits': {},
    'daily': {},
    'weekly': {},
    'settings': {'hourly_rate': 21.0, 'tax_reserve_pct': 30.0},
}

class Storage:
    def __init__(self):
        self.mode = 'local'
        self.client = None
        self.sheet = None
        self.state_ws = None
        self.snap_ws = None

        # 1) Google Sheets (preferred when configured)
        try:
            import streamlit as st
            sheet_id = st.secrets.get('GOOGLE_SHEET_ID', '')
            svc = st.secrets.get('gcp_service_account', None)
            if sheet_id and svc:
                import gspread
                from google.oauth2.service_account import Credentials
                scopes = [
                    'https://www.googleapis.com/auth/spreadsheets',
                    'https://www.googleapis.com/auth/drive',
                ]
                creds = Credentials.from_service_account_info(dict(svc), scopes=scopes)
                gc = gspread.authorize(creds)
                self.sheet = gc.open_by_key(sheet_id)
                self.state_ws = self._ensure_ws('_app_state', rows=20, cols=5)
                self.snap_ws = self._ensure_ws('_pdf_snapshots', rows=1000, cols=5)
                if not self.state_ws.acell('A1').value:
                    self.state_ws.update('A1:B1', [['state_json','updated_at']])
                if not self.snap_ws.acell('A1').value:
                    self.snap_ws.update('A1:D1', [['filename','sha256','parsed_json','created_at']])
                self.mode = 'google_sheets'
        except Exception:
            pass

        # 2) Supabase fallback
        if self.mode == 'local':
            try:
                import streamlit as st
                url = st.secrets.get('SUPABASE_URL', '')
                key = st.secrets.get('SUPABASE_SERVICE_ROLE_KEY', '')
                if url and key:
                    from supabase import create_client
                    self.client = create_client(url, key)
                    self.mode = 'supabase'
            except Exception:
                pass

        self.local_path = os.path.join(os.path.dirname(__file__), '.mca_state.json')

    def _ensure_ws(self, title, rows=100, cols=10):
        try:
            return self.sheet.worksheet(title)
        except Exception:
            return self.sheet.add_worksheet(title=title, rows=rows, cols=cols)

    def load_state(self):
        if self.mode == 'google_sheets':
            try:
                raw = self.state_ws.acell('A2').value
                if raw:
                    return json.loads(raw)
            except Exception:
                pass

        if self.mode == 'supabase':
            try:
                r = self.client.table('mca_state').select('data').eq('id','singleton').execute()
                if r.data:
                    return r.data[0]['data']
            except Exception:
                pass

        if os.path.exists(self.local_path):
            try:
                with open(self.local_path,'r',encoding='utf-8') as f:
                    return json.load(f)
            except Exception:
                pass
        return json.loads(json.dumps(DEFAULT_STATE))

    def save_state(self, state):
        if self.mode == 'google_sheets':
            payload = json.dumps(state, ensure_ascii=False, separators=(',',':'))
            self.state_ws.update('A2:B2', [[payload, datetime.now(timezone.utc).isoformat()]])
            return

        if self.mode == 'supabase':
            self.client.table('mca_state').upsert({
                'id':'singleton','data':state,'updated_at':datetime.now(timezone.utc).isoformat()
            }).execute()
            return

        with open(self.local_path,'w',encoding='utf-8') as f:
            json.dump(state,f,ensure_ascii=False,indent=2)

    def latest_snapshot(self):
        if self.mode == 'google_sheets':
            try:
                rows = self.snap_ws.get_all_values()
                if len(rows) <= 1:
                    return None
                filename, sha, parsed_json, created_at = rows[-1][:4]
                return {
                    'filename': filename,
                    'sha256': sha,
                    'parsed': json.loads(parsed_json) if parsed_json else [],
                    'created_at': created_at,
                }
            except Exception:
                return None

        if self.mode == 'supabase':
            try:
                r = self.client.table('mca_pdf_snapshots').select('*').order('created_at', desc=True).limit(1).execute()
                return r.data[0] if r.data else None
            except Exception:
                return None

        path = os.path.join(os.path.dirname(__file__), '.mca_snapshot.json')
        if os.path.exists(path):
            try:
                return json.load(open(path,encoding='utf-8'))
            except Exception:
                return None
        return None

    def save_snapshot(self, filename, sha256, parsed):
        created = datetime.now(timezone.utc).isoformat()
        obj = {'filename':filename,'sha256':sha256,'parsed':parsed,'created_at':created}

        if self.mode == 'google_sheets':
            self.snap_ws.append_row([
                filename,
                sha256,
                json.dumps(parsed, ensure_ascii=False, separators=(',',':')),
                created
            ], value_input_option='RAW')
            return

        if self.mode == 'supabase':
            self.client.table('mca_pdf_snapshots').insert(obj).execute()
            return

        path = os.path.join(os.path.dirname(__file__), '.mca_snapshot.json')
        with open(path,'w',encoding='utf-8') as f:
            json.dump(obj,f,ensure_ascii=False,indent=2)

    def history(self, limit=10):
        if self.mode == 'google_sheets':
            try:
                rows = self.snap_ws.get_all_values()[1:]
                out = []
                for row in reversed(rows[-limit:]):
                    while len(row) < 4:
                        row.append('')
                    out.append({
                        'filename': row[0],
                        'sha256': row[1],
                        'parsed': json.loads(row[2]) if row[2] else [],
                        'created_at': row[3],
                    })
                return out
            except Exception:
                return []

        if self.mode == 'supabase':
            try:
                r = self.client.table('mca_pdf_snapshots').select('filename,sha256,created_at,parsed').order('created_at', desc=True).limit(limit).execute()
                return r.data or []
            except Exception:
                return []

        snap = self.latest_snapshot()
        return [snap] if snap else []
