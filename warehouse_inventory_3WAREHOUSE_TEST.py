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

WAREHOUSE_CATEGORIES = {
    'Buntun': ('Torani', 'Powders', 'Pastries', 'Store Items'),
    'Echague': ('Cups & Lids', 'Powders'),
    'Santa Maria': ('Cups & Lids', 'Powders'),
}
TZ = ZoneInfo('Asia/Manila')

def _request(path, method='GET', params=None, data=None):
    url = str(st.secrets.get('SUPABASE_URL', '')).strip().rstrip('/')
    # Accept either the base Supabase Project URL or a Data API URL copied from the dashboard.
    # REST paths below are appended by this module, so normalize any copied /rest/v1 suffix away.
    if url.endswith('/rest/v1'):
        url = url[:-8].rstrip('/')
    key = str(st.secrets.get('SUPABASE_SECRET_KEY', '')).strip()
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

def _warehouses():
    rows = _request('warehouses', params={'select':'id,name,active','active':'eq.true','order':'name.asc'})
    return {r['name']: r['id'] for r in rows if r.get('name') in WAREHOUSE_CATEGORIES}

def _items(warehouse_id, category):
    return _request('warehouse_items', params={
        'select':'id,warehouse_id,category,name,unit,active',
        'warehouse_id':'eq.'+warehouse_id,
        'category':'eq.'+category,
        'active':'eq.true',
        'order':'name.asc'
    })

def _balance(warehouse_id, category):
    return _request('warehouse_balances', params={
        'select':'item_id,warehouse_id,category,name,unit,balance',
        'warehouse_id':'eq.'+warehouse_id,
        'category':'eq.'+category,
        'order':'name.asc'
    })

def _template(warehouse_name):
    cats = WAREHOUSE_CATEGORIES[warehouse_name]
    frame = pd.DataFrame({
        'Category': list(cats),
        'Item': [''] * len(cats),
        'Quantity': [0] * len(cats),
        'Reference': [''] * len(cats)
    })
    buf=io.BytesIO()
    with pd.ExcelWriter(buf, engine='openpyxl') as writer:
        frame.to_excel(writer,index=False,sheet_name='Receiving')
    return buf.getvalue()

def render_warehouse_inventory(username, role):
    if role not in ('Admin','Super Admin'):
        st.error('Warehouse Inventory is available to Admin and Super Admin only.')
        return

    st.title('Warehouse Inventory')
    st.caption('TEST VERSION • Separate inventory for Buntun, Echague and Santa Maria • Stored in Supabase')

    try:
        warehouse_map = _warehouses()
    except Exception as exc:
        st.error(str(exc)); return

    if not warehouse_map:
        st.error('No TEST warehouses found. Run the warehouse migration first.')
        return

    preferred = [w for w in ('Buntun','Echague','Santa Maria') if w in warehouse_map]
    warehouse_name = st.selectbox('Warehouse', preferred, key='wh_warehouse')
    warehouse_id = warehouse_map[warehouse_name]
    category = st.selectbox('Category', WAREHOUSE_CATEGORIES[warehouse_name], key='wh_category_v2')

    try:
        items = _items(warehouse_id, category)
        balances = _balance(warehouse_id, category)
    except Exception as exc:
        st.error(str(exc)); return

    tabs = st.tabs(['Stock Balance','Receive Stock','Monday Stock Count','Item Masterlist','Dispatch History'])

    with tabs[0]:
        st.subheader(f'{warehouse_name} — {category}')
        if balances:
            view = pd.DataFrame(balances).rename(columns={'name':'Item','unit':'Unit','balance':'Available Stock'})
            st.dataframe(view[['Item','Unit','Available Stock']],hide_index=True,use_container_width=True)
        else:
            st.info('No items yet for this warehouse/category. Add them under Item Masterlist.')

    with tabs[1]:
        st.subheader(f'Receive Stock — {warehouse_name}')
        if category in ('Pastries','Store Items'):
            st.info('For Buntun Pastries and Store Items, the dedicated Receive → Dispatch → Reconcile workflow will be added in the next TEST stage. For now, do not use normal receiving for these categories.')
        else:
            mode=st.radio('Receiving method',['Manual Entry','Excel Upload'],horizontal=True,key='wh_receive_mode_v2')
            if mode=='Manual Entry':
                if not items:
                    st.info('Add an item to the masterlist first.')
                else:
                    item_map={f"{r['name']} ({r['unit']})":r['id'] for r in items}
                    selected=st.selectbox('Item',list(item_map),key='wh_receive_item_v2')
                    qty=st.number_input('Received quantity',min_value=0.0,step=1.0,key='wh_receive_qty_v2')
                    ref=st.text_input('Supplier / receiving reference',key='wh_receive_ref_v2')
                    if st.button('SAVE RECEIVED STOCK',type='primary',disabled=qty<=0,key='wh_receive_save_v2'):
                        try:
                            _rpc('warehouse_receive',{'p_item_id':item_map[selected],'p_qty':qty,'p_reference':ref.strip(),'p_actor':username,'p_batch_key':None})
                            st.success(f'Stock received into {warehouse_name}.'); st.rerun()
                        except Exception as exc: st.error(str(exc))
            else:
                st.download_button('Download Excel Template',_template(warehouse_name),f'{warehouse_name}_Receiving_Template.xlsx',key='wh_template_v2')
                upload=st.file_uploader('Upload completed receiving sheet',type=['xlsx'],key='wh_excel_v2')
                if upload:
                    try:
                        frame=pd.read_excel(upload)
                        required={'Category','Item','Quantity','Reference'}
                        if not required.issubset(frame.columns):
                            raise ValueError('Required columns: Category, Item, Quantity, Reference.')
                        frame=frame.dropna(how='all').copy()
                        frame['Quantity']=pd.to_numeric(frame['Quantity'],errors='raise')
                        if frame.empty or frame['Quantity'].isna().any() or (frame['Quantity']<=0).any():
                            raise ValueError('Every row must have a positive quantity.')
                        allowed=set(WAREHOUSE_CATEGORIES[warehouse_name])
                        if any(str(x).strip() not in allowed for x in frame['Category']):
                            raise ValueError(f'One or more categories are not assigned to {warehouse_name}.')
                        all_items=_request('warehouse_items',params={
                            'select':'id,warehouse_id,category,name,active',
                            'warehouse_id':'eq.'+warehouse_id,
                            'active':'eq.true'
                        })
                        lookup={(r['category'].strip().casefold(),r['name'].strip().casefold()):r['id'] for r in all_items}
                        payload=[]
                        for _,row in frame.iterrows():
                            key=(str(row['Category']).strip().casefold(),str(row['Item']).strip().casefold())
                            if key not in lookup:
                                raise ValueError('Unknown item/category for '+warehouse_name+': '+str(row['Category'])+' / '+str(row['Item']))
                            payload.append({'item_id':lookup[key],'qty':float(row['Quantity']),'reference':'' if pd.isna(row['Reference']) else str(row['Reference'])})
                        st.dataframe(frame,hide_index=True,use_container_width=True)
                        st.caption('Preview only until Save. This upload applies only to the selected warehouse.')
                        if st.button('SAVE UPLOADED RECEIVING',type='primary',key='wh_upload_save_v2'):
                            import hashlib
                            digest=hashlib.sha256((warehouse_id+':').encode()+upload.getvalue()).hexdigest()
                            _rpc('warehouse_receive_batch',{'p_rows':payload,'p_actor':username,'p_batch_key':digest})
                            st.success(f'Receiving sheet saved to {warehouse_name}.'); st.rerun()
                    except Exception as exc: st.error(str(exc))

    with tabs[2]:
        today=datetime.now(TZ).date()
        st.subheader(f'Monday Stock Count — {warehouse_name}')
        if category in ('Pastries','Store Items'):
            st.info('Pastries and Store Items use Receive → Dispatch → Reconcile rather than the regular Monday warehouse count.')
        else:
            st.caption('A variance will be shown for review. Automatic stock adjustment is not enabled in this TEST stage.')
            if today.weekday()!=0:
                st.warning('Today is not Monday. You can view balances; submission opens on Monday.')
            if balances:
                frame=pd.DataFrame([{'Item ID':r['item_id'],'Item':r['name'],'Expected':float(r['balance']),'Actual':None} for r in balances])
                edited=st.data_editor(
                    frame,hide_index=True,use_container_width=True,
                    disabled=['Item ID','Item','Expected'],
                    column_config={'Actual':st.column_config.NumberColumn('Actual',min_value=0.0,required=True)},
                    key='wh_count_'+warehouse_name+'_'+category
                )
                edited['Variance']=pd.to_numeric(edited['Expected'])-pd.to_numeric(edited['Actual'],errors='coerce')
                st.dataframe(edited[['Item','Expected','Actual','Variance']],hide_index=True,use_container_width=True)
                has_variance = edited['Actual'].notna().all() and (edited['Variance'].fillna(0).abs() > 0.000001).any()
                if has_variance:
                    st.warning('Variance detected. This count should be reviewed before any stock adjustment.')
                if st.button('SUBMIT MONDAY STOCK COUNT',type='primary',disabled=today.weekday()!=0,key='wh_count_submit_v2'):
                    if edited['Actual'].isna().any():
                        st.error('Count every item; enter 0 for no stock.')
                    elif has_variance:
                        st.error('Variance detected. TEST rule: the count will not proceed until the variance review/proof workflow is added.')
                    else:
                        try:
                            rows=[{'item_id':r['Item ID'],'expected':float(r['Expected']),'actual':float(r['Actual'])} for _,r in edited.iterrows()]
                            _rpc('warehouse_submit_count',{
                                'p_warehouse_id':warehouse_id,
                                'p_category':category,
                                'p_count_date':today.isoformat(),
                                'p_rows':rows,
                                'p_actor':username
                            })
                            st.success('Monday count submitted with no variance.'); st.rerun()
                        except Exception as exc: st.error(str(exc))

    with tabs[3]:
        st.subheader(f'Item Masterlist — {warehouse_name} / {category}')
        name=st.text_input('Item name',key='wh_new_item_v2')
        unit=st.selectbox('Unit',['pcs','packs','boxes','bottles','kg','g','L','mL'],key='wh_new_unit_v2')
        if st.button('ADD ITEM',disabled=not name.strip(),key='wh_add_item_v2'):
            try:
                _request('warehouse_items','POST',data={
                    'warehouse_id':warehouse_id,
                    'category':category,
                    'name':name.strip(),
                    'unit':unit,
                    'active':True
                })
                st.success(f'Item added to {warehouse_name}.'); st.rerun()
            except Exception as exc: st.error(str(exc))
        if items:
            st.dataframe(pd.DataFrame(items)[['name','unit']],hide_index=True,use_container_width=True)

    with tabs[4]:
        st.subheader(f'Dispatch History — {warehouse_name}')
        st.caption('Automatic deduction after branch receiving is not connected yet. This tab will become the detailed permanent dispatch audit trail.')
        try:
            item_rows=_request('warehouse_items',params={'select':'id,name,category','warehouse_id':'eq.'+warehouse_id})
            item_lookup={r['id']:r for r in item_rows}
            if item_lookup:
                rows=_request('warehouse_movements',params={
                    'select':'created_at,item_id,kind,quantity,reference,actor',
                    'kind':'eq.dispatch',
                    'order':'created_at.desc',
                    'limit':'500'
                })
                rows=[r for r in rows if r.get('item_id') in item_lookup]
                for r in rows:
                    meta=item_lookup[r['item_id']]
                    r['category']=meta['category']
                    r['item']=meta['name']
                if rows:
                    view=pd.DataFrame(rows).rename(columns={
                        'created_at':'Date/Time','category':'Category','item':'Item',
                        'quantity':'Quantity','reference':'Reference','actor':'Processed By'
                    })
                    st.dataframe(view[['Date/Time','Category','Item','Quantity','Reference','Processed By']],hide_index=True,use_container_width=True)
                else:
                    st.info('No warehouse dispatches yet.')
            else:
                st.info('No items yet for this warehouse.')
        except Exception as exc:
            st.error(str(exc))

# When the entire order/delivery workflow has migrated to Supabase, call this
# from the dispatch transaction, NOT from an independent SQLite commit.
def dispatch_once(delivery_id, lines, actor):
    """Atomic, idempotent Supabase deduction; lines: [{item_id, qty}].
    Raises on insufficient stock, unknown items, or a changed retry payload.
    This is deliberately not connected to the legacy SQLite dispatch button.
    """
    return _rpc('warehouse_dispatch',{'p_delivery_id':str(delivery_id),'p_rows':lines,'p_actor':actor})
