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

    def load_state(self):
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
        if self.mode == 'supabase':
            self.client.table('mca_state').upsert({
                'id':'singleton','data':state,'updated_at':datetime.now(timezone.utc).isoformat()
            }).execute()
            return
        with open(self.local_path,'w',encoding='utf-8') as f:
            json.dump(state,f,ensure_ascii=False,indent=2)

    def latest_snapshot(self):
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
        obj = {'filename':filename,'sha256':sha256,'parsed':parsed,'created_at':datetime.now(timezone.utc).isoformat()}
        if self.mode == 'supabase':
            self.client.table('mca_pdf_snapshots').insert(obj).execute()
            return
        path = os.path.join(os.path.dirname(__file__), '.mca_snapshot.json')
        with open(path,'w',encoding='utf-8') as f:
            json.dump(obj,f,ensure_ascii=False,indent=2)

    def history(self, limit=10):
        if self.mode == 'supabase':
            try:
                r = self.client.table('mca_pdf_snapshots').select('filename,sha256,created_at,parsed').order('created_at', desc=True).limit(limit).execute()
                return r.data or []
            except Exception:
                return []
        snap = self.latest_snapshot()
        return [snap] if snap else []
