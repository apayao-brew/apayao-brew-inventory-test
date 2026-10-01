"""Apayao Brew main warehouse inventory — TEST MODULE, Supabase-backed.
Run warehouse_inventory_schema.sql in a TEST Supabase project first.
No local SQLite database is read or written by this module.
"""
import io
import json
import urllib.parse
import urllib.request
import urllib.error
from datetime import datetime
from zoneinfo import ZoneInfo

import pandas as pd
import streamlit as st

CATEGORIES = ('Cups & Lids', 'Powders', 'Torani')
TZ = ZoneInfo('Asia/Manila')

def _request(path, method='GET', params=None, data=None):
    url = str(st.secrets.get('SUPABASE_URL', '')).rstrip('/')
    key = str(st.secrets.get('SUPABASE_SECRET_KEY', ''))
    if not url or not key:
        raise RuntimeError('Set SUPABASE_URL and SUPABASE_SECRET_KEY in the TEST app secrets.')
    endpoint = url + '/rest/v1/' + path
    if params:
        endpoint += '?' + urllib.parse.urlencode(params)
    headers = {'apikey': key, 'Authorization': 'Bearer ' + key,
               'Content-Type': 'application/json', 'Prefer': 'return=representation'}
    body = None if data is None else json.dumps(data, allow_nan=False).encode('utf-8')
    try:
        with urllib.request.urlopen(urllib.request.Request(endpoint, data=body, headers=headers, method=method), timeout=30) as response:
            content = response.read()
            return json.loads(content) if content else []
    except urllib.error.HTTPError as exc:
        raise RuntimeError('Inventory database error: ' + exc.read().decode('utf-8', 'replace')[:400]) from exc

def _rpc(name, payload):
    return _request('rpc/' + name, 'POST', data=payload)

def _items(category):
    return _request('warehouse_items', params={'select':'id,category,name,unit,active', 'category':'eq.'+category, 'active':'eq.true', 'order':'name.asc'})

def _balance(category):
    return _request('warehouse_balances', params={'select':'item_id,category,name,unit,balance', 'category':'eq.'+category, 'order':'name.asc'})

def _template():
    frame = pd.DataFrame({'Category':['Cups & Lids','Powders','Torani'], 'Item':['22oz Cup','Chocolate Powder','Caramel Syrup'], 'Quantity':[0,0,0], 'Reference':['','','']})
    buf=io.BytesIO()
    with pd.ExcelWriter(buf, engine='openpyxl') as writer:
        frame.to_excel(writer,index=False,sheet_name='Receiving')
    return buf.getvalue()

def render_warehouse_inventory(username, role):
    if role not in ('Admin','Super Admin'):
        st.error('Warehouse Inventory is available to Admin and Super Admin only.')
        return
    st.title('Main Warehouse Inventory')
    st.caption('TEST VERSION • Cups & Lids, Powders and Torani • Weekly count every Monday • Stored in Supabase')
    category = st.selectbox('Category', CATEGORIES, key='wh_category')
    try:
        items = _items(category)
        balances = _balance(category)
    except Exception as exc:
        st.error(str(exc)); return
    tabs = st.tabs(['Stock Balance','Receive Stock','Monday Stock Count','Item Masterlist','Dispatch History'])
    with tabs[0]:
        if balances:
            st.dataframe(pd.DataFrame(balances).rename(columns={'name':'Item','unit':'Unit','balance':'Available Stock'})[['Item','Unit','Available Stock']],hide_index=True,use_container_width=True)
        else:
            st.info('No items yet. Add them under Item Masterlist.')
    with tabs[1]:
        st.subheader('Receive Stock')
        mode=st.radio('Receiving method',['Manual Entry','Excel Upload'],horizontal=True)
        if mode=='Manual Entry':
            if not items: st.info('Add an item to the masterlist first.')
            else:
                item_map={f"{r['name']} ({r['unit']})":r['id'] for r in items}
                selected=st.selectbox('Item',list(item_map),key='wh_receive_item')
                qty=st.number_input('Received quantity',min_value=0.0,step=1.0,key='wh_receive_qty')
                ref=st.text_input('Supplier / receiving reference',key='wh_receive_ref')
                if st.button('SAVE RECEIVED STOCK',type='primary',disabled=qty<=0):
                    try:
                        _rpc('warehouse_receive',{'p_item_id':item_map[selected],'p_qty':qty,'p_reference':ref.strip(),'p_actor':username,'p_batch_key':None})
                        st.success('Stock received and recorded.'); st.rerun()
                    except Exception as exc: st.error(str(exc))
        else:
            st.download_button('Download Excel Template',_template(),'Warehouse_Receiving_Template.xlsx')
            upload=st.file_uploader('Upload completed receiving sheet',type=['xlsx'],key='wh_excel')
            if upload:
                try:
                    frame=pd.read_excel(upload)
                    required={'Category','Item','Quantity','Reference'}
                    if not required.issubset(frame.columns): raise ValueError('Required columns: Category, Item, Quantity, Reference.')
                    frame=frame.dropna(how='all').copy()
                    frame['Quantity']=pd.to_numeric(frame['Quantity'],errors='raise')
                    if frame.empty or frame['Quantity'].isna().any() or (frame['Quantity']<=0).any(): raise ValueError('Every row must have a positive quantity.')
                    all_items=_request('warehouse_items',params={'select':'id,category,name,active','active':'eq.true'})
                    lookup={(r['category'].strip().casefold(),r['name'].strip().casefold()):r['id'] for r in all_items}
                    payload=[]
                    for _,row in frame.iterrows():
                        key=(str(row['Category']).strip().casefold(),str(row['Item']).strip().casefold())
                        if key not in lookup: raise ValueError('Unknown item/category: '+str(row['Category'])+' / '+str(row['Item']))
                        payload.append({'item_id':lookup[key],'qty':float(row['Quantity']),'reference':'' if pd.isna(row['Reference']) else str(row['Reference'])})
                    st.dataframe(frame,hide_index=True,use_container_width=True)
                    st.caption('Preview only until you press Save. The whole upload is saved in one database transaction.')
                    if st.button('SAVE UPLOADED RECEIVING',type='primary'):
                        import hashlib
                        digest=hashlib.sha256(upload.getvalue()).hexdigest()
                        _rpc('warehouse_receive_batch',{'p_rows':payload,'p_actor':username,'p_batch_key':digest})
                        st.success('Receiving sheet saved. Uploading the same file again will not duplicate stock.');st.rerun()
                except Exception as exc: st.error(str(exc))
    with tabs[2]:
        today=datetime.now(TZ).date()
        st.caption('Scheduled every Monday. Counts on other days require Super Admin approval in the future.')
        if today.weekday()!=0: st.warning('Today is not Monday. You can view balances; submission opens on Monday.')
        if balances:
            frame=pd.DataFrame([{'Item ID':r['item_id'],'Item':r['name'],'Expected':float(r['balance']),'Actual':None} for r in balances])
            edited=st.data_editor(frame,hide_index=True,use_container_width=True,disabled=['Item ID','Item','Expected'],column_config={'Actual':st.column_config.NumberColumn('Actual',min_value=0.0,required=True)},key='wh_count_'+category)
            edited['Variance']=pd.to_numeric(edited['Expected'])-pd.to_numeric(edited['Actual'],errors='coerce')
            st.dataframe(edited[['Item','Expected','Actual','Variance']],hide_index=True,use_container_width=True)
            if st.button('SUBMIT MONDAY STOCK COUNT',type='primary',disabled=today.weekday()!=0):
                if edited['Actual'].isna().any(): st.error('Count every item; enter 0 for no stock.')
                else:
                    try:
                        rows=[{'item_id':r['Item ID'],'expected':float(r['Expected']),'actual':float(r['Actual'])} for _,r in edited.iterrows()]
                        _rpc('warehouse_submit_count',{'p_category':category,'p_count_date':today.isoformat(),'p_rows':rows,'p_actor':username})
                        st.success('Monday count submitted; differences are recorded without silently changing the stock ledger.');st.rerun()
                    except Exception as exc: st.error(str(exc))
    with tabs[3]:
        st.subheader('Add warehouse item')
        name=st.text_input('Item name',key='wh_new_item')
        unit=st.selectbox('Unit',['pcs','packs','boxes','bottles','kg','g','L','mL'],key='wh_new_unit')
        if st.button('ADD ITEM',disabled=not name.strip()):
            try:
                _request('warehouse_items','POST',data={'category':category,'name':name.strip(),'unit':unit,'active':True})
                st.success('Item added.');st.rerun()
            except Exception as exc: st.error(str(exc))
        if items: st.dataframe(pd.DataFrame(items)[['name','unit']],hide_index=True,use_container_width=True)
    with tabs[4]:
        st.caption('Dispatch is recorded by a unique delivery ID. Repeated confirmation cannot deduct the same delivery twice.')
        try:
            rows=_request('warehouse_movements',params={'select':'created_at,item_id,kind,quantity,reference,actor','kind':'eq.dispatch','order':'created_at.desc','limit':'200'})
            st.dataframe(pd.DataFrame(rows),hide_index=True,use_container_width=True) if rows else st.info('No warehouse dispatches yet.')
        except Exception as exc: st.error(str(exc))

# When the entire order/delivery workflow has migrated to Supabase, call this
# from the dispatch transaction, NOT from an independent SQLite commit.
def dispatch_once(delivery_id, lines, actor):
    """Atomic, idempotent Supabase deduction; lines: [{item_id, qty}].
    Raises on insufficient stock, unknown items, or a changed retry payload.
    This is deliberately not connected to the legacy SQLite dispatch button.
    """
    return _rpc('warehouse_dispatch',{'p_delivery_id':str(delivery_id),'p_rows':lines,'p_actor':actor})
