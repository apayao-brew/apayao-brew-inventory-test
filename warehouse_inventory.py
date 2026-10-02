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


def _weekly_inventory(warehouse_id, category):
    return _request('warehouse_weekly_inventory', params={
        'select':'item_id,warehouse_id,category,item_name,unit,beginning,deliveries,total,pull_out,stock,previous_count_date',
        'warehouse_id':'eq.'+warehouse_id,
        'category':'eq.'+category,
        'order':'item_name.asc'
    })

def _inventory_history(warehouse_id, category):
    return _request('warehouse_inventory_history', params={
        'select':'count_id,warehouse_id,warehouse,category,count_date,status,submitted_by,submitted_at,variance_explanation,proof_reference,reviewed_by,reviewed_at,item_id,item_name,unit,beginning,deliveries,total,pull_out,stock,actual,variance',
        'warehouse_id':'eq.'+warehouse_id,
        'category':'eq.'+category,
        'order':'count_date.desc,item_name.asc'
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

def render_warehouse_inventory(username, role, assigned_warehouse=""):
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
    if role == 'Admin':
        # Admin is the warehouse account. It is locked to one assigned warehouse.
        normalized = str(assigned_warehouse or '').strip()
        if normalized.endswith(' Warehouse'):
            normalized = normalized[:-10].strip()
        if normalized not in warehouse_map:
            st.error('This Admin account has no valid Assigned Warehouse. Ask Super Admin to assign Buntun, Echague, or Santa Maria in User Management.')
            return
        warehouse_name = normalized
        st.text_input('Assigned Warehouse', value=warehouse_name, disabled=True, key='wh_assigned_warehouse')
    else:
        warehouse_name = st.selectbox('Warehouse', preferred, key='wh_warehouse')
    warehouse_id = warehouse_map[warehouse_name]
    category = st.selectbox('Category', WAREHOUSE_CATEGORIES[warehouse_name], key='wh_category_v2')

    try:
        items = _items(warehouse_id, category)
        balances = _balance(warehouse_id, category)
    except Exception as exc:
        st.error(str(exc)); return

    tabs = st.tabs(['Stock Balance','Monday Stock Count','Item Masterlist','Direct Dispatch','Dispatch History'])

    with tabs[0]:
        st.subheader(f'{warehouse_name} — {category}')
        if balances:
            view = pd.DataFrame(balances).rename(columns={'name':'Item','unit':'Unit','balance':'Available Stock'})
            st.dataframe(view[['Item','Unit','Available Stock']],hide_index=True,use_container_width=True)
        else:
            st.info('No items yet for this warehouse/category. Add them under Item Masterlist.')
        st.caption('Supplier receiving is handled only in Warehouse Deliveries to prevent duplicate stock posting.')

    with tabs[1]:
        today=datetime.now(TZ).date()
        st.subheader(f'Monday Stock Count — {warehouse_name}')
        if category in ('Pastries','Store Items'):
            st.info('Pastries and Store Items use their separate Receive → Dispatch → Reconcile workflow.')
        else:
            try:
                weekly=_weekly_inventory(warehouse_id,category)
            except Exception as exc:
                st.error(str(exc)); weekly=[]

            st.caption('Beginning = previous approved Monday Actual • Deliveries = confirmed warehouse deliveries • Pull Out = completed dispatches • Actual is the only editable inventory quantity.')
            if today.weekday()!=0:
                st.warning('Today is not Monday. Current balances and past inventory remain available; submission opens on Monday.')

            if weekly:
                frame=pd.DataFrame([{
                    'Item ID':r['item_id'],
                    'Item':r.get('item_name',''),
                    'Unit':r.get('unit',''),
                    'Beginning':float(r.get('beginning') or 0),
                    'Deliveries':float(r.get('deliveries') or 0),
                    'Total':float(r.get('total') or 0),
                    'Pull Out':float(r.get('pull_out') or 0),
                    'Stock':float(r.get('stock') or 0),
                    'Actual':None
                } for r in weekly])
                edited=st.data_editor(
                    frame,hide_index=True,use_container_width=True,
                    disabled=['Item ID','Item','Unit','Beginning','Deliveries','Total','Pull Out','Stock'],
                    column_config={'Actual':st.column_config.NumberColumn('Actual',min_value=0.0,required=True)},
                    key='wh_count_final_'+warehouse_name+'_'+category
                )
                edited['Variance']=pd.to_numeric(edited['Stock'])-pd.to_numeric(edited['Actual'],errors='coerce')
                st.dataframe(edited[['Item','Unit','Beginning','Deliveries','Total','Pull Out','Stock','Actual','Variance']],hide_index=True,use_container_width=True)
                complete=edited['Actual'].notna().all()
                has_variance=complete and (edited['Variance'].fillna(0).abs()>0.000001).any()
                explanation=''; proof_reference=''
                if has_variance:
                    st.warning('Variance detected. Explanation and proof/reference are required and the count will be sent FOR REVIEW.')
                    explanation=st.text_area('Variance Explanation',key='wh_variance_explanation')
                    proof_reference=st.text_input('Variance Proof / Reference',help='Enter the proof filename, reference number, or other traceable proof reference.',key='wh_variance_proof_ref')

                submit_disabled=today.weekday()!=0 or not complete
                if st.button('SUBMIT MONDAY STOCK COUNT',type='primary',disabled=submit_disabled,key='wh_count_submit_final'):
                    if has_variance and (not explanation.strip() or not proof_reference.strip()):
                        st.error('Variance Explanation and Variance Proof / Reference are required.')
                    else:
                        try:
                            rows=[]
                            for _,r in edited.iterrows():
                                rows.append({
                                    'item_id':r['Item ID'],
                                    'beginning':float(r['Beginning']),
                                    'deliveries':float(r['Deliveries']),
                                    'total':float(r['Total']),
                                    'pull_out':float(r['Pull Out']),
                                    'stock':float(r['Stock']),
                                    'actual':float(r['Actual'])
                                })
                            _rpc('warehouse_submit_count',{
                                'p_warehouse_id':warehouse_id,
                                'p_category':category,
                                'p_count_date':today.isoformat(),
                                'p_rows':rows,
                                'p_actor':username,
                                'p_explanation':explanation.strip() or None,
                                'p_proof_reference':proof_reference.strip() or None
                            })
                            if has_variance:
                                st.success('Monday count submitted FOR REVIEW with the full inventory snapshot saved.')
                            else:
                                st.success('Monday count submitted and APPROVED. The full inventory snapshot was saved.')
                            st.rerun()
                        except Exception as exc: st.error(str(exc))
            else:
                st.info('No regular inventory items found for this warehouse/category.')

            st.divider()
            st.subheader('Past Inventory & Balances')
            try:
                history=_inventory_history(warehouse_id,category)
                if history:
                    hdf=pd.DataFrame(history)
                    dates=sorted([str(x) for x in hdf['count_date'].dropna().unique()],reverse=True)
                    selected_date=st.selectbox('Inventory Date',dates,key='wh_history_date_'+warehouse_name+'_'+category)
                    selected=hdf[hdf['count_date'].astype(str)==selected_date].copy()
                    if not selected.empty:
                        first=selected.iloc[0]
                        st.caption(
                            f"Status: {first.get('status','')} • Submitted by: {first.get('submitted_by','')}"
                            + (f" • Reviewed by: {first.get('reviewed_by','')}" if pd.notna(first.get('reviewed_by')) and first.get('reviewed_by') else '')
                        )
                        hist_view=selected.rename(columns={
                            'item_name':'Item','unit':'Unit','beginning':'Beginning','deliveries':'Deliveries',
                            'total':'Total','pull_out':'Pull Out','stock':'Stock','actual':'Actual','variance':'Variance'
                        })
                        cols=['Item','Unit','Beginning','Deliveries','Total','Pull Out','Stock','Actual','Variance']
                        st.dataframe(hist_view[cols],hide_index=True,use_container_width=True)
                        if first.get('variance_explanation'):
                            st.write('Variance Explanation:',first.get('variance_explanation'))
                        if first.get('proof_reference'):
                            st.write('Proof / Reference:',first.get('proof_reference'))
                        export=io.BytesIO()
                        with pd.ExcelWriter(export,engine='openpyxl') as writer:
                            hist_view[cols].to_excel(writer,index=False,sheet_name='Monday Inventory')
                            pd.DataFrame([{
                                'Warehouse':warehouse_name,'Category':category,'Inventory Date':selected_date,
                                'Status':first.get('status',''),'Submitted By':first.get('submitted_by',''),
                                'Reviewed By':first.get('reviewed_by',''),'Variance Explanation':first.get('variance_explanation',''),
                                'Proof / Reference':first.get('proof_reference','')
                            }]).to_excel(writer,index=False,sheet_name='Record Details')
                        st.download_button('Download Historical Inventory Excel',export.getvalue(),f'{warehouse_name}_{category}_{selected_date}_Inventory.xlsx',key='wh_history_download_'+warehouse_name+'_'+category)
                else:
                    st.info('No past Monday inventory records yet for this warehouse/category.')
            except Exception as exc:
                st.error(str(exc))

    with tabs[2]:
        st.subheader(f'Item Masterlist — {warehouse_name} / {category}')
        if role == 'Super Admin':
            master_mode = st.radio('Masterlist method', ['Manual Add','Excel Upload'], horizontal=True, key='wh_master_mode_v3')
            allowed_units=['pcs','packs','boxes','bottles','kg','g','L','mL']
            if master_mode == 'Manual Add':
                name=st.text_input('Item name',key='wh_new_item_v3')
                unit=st.selectbox('Unit',allowed_units,key='wh_new_unit_v3')
                if st.button('ADD ITEM',disabled=not name.strip(),key='wh_add_item_v3'):
                    try:
                        existing={str(r['name']).strip().casefold() for r in items}
                        if name.strip().casefold() in existing:
                            raise ValueError('This item already exists in the selected warehouse/category.')
                        _request('warehouse_items','POST',data={'warehouse_id':warehouse_id,'category':category,'name':name.strip(),'unit':unit,'active':True})
                        st.success(f'Item added to {warehouse_name} / {category}.'); st.rerun()
                    except Exception as exc: st.error(str(exc))
            else:
                template=pd.DataFrame({'Item Name':['Example Item'],'Unit':['pcs']})
                buf=io.BytesIO()
                with pd.ExcelWriter(buf,engine='openpyxl') as writer: template.to_excel(writer,index=False,sheet_name='Item Masterlist')
                st.download_button('Download Excel Template',buf.getvalue(),f'{warehouse_name}_{category}_Item_Masterlist.xlsx',key='wh_master_template_v3')
                upload=st.file_uploader('Upload Item Masterlist Excel',type=['xlsx'],key='wh_master_upload_v3')
                if upload:
                    try:
                        frame=pd.read_excel(upload).dropna(how='all').copy()
                        required={'Item Name','Unit'}
                        if not required.issubset(frame.columns): raise ValueError('Required columns: Item Name, Unit.')
                        frame=frame[['Item Name','Unit']]
                        frame['Item Name']=frame['Item Name'].fillna('').astype(str).str.strip()
                        frame['Unit']=frame['Unit'].fillna('').astype(str).str.strip()
                        if frame.empty or (frame['Item Name']=='').any(): raise ValueError('Blank Item Name is not allowed.')
                        invalid=frame[~frame['Unit'].isin(allowed_units)]
                        if not invalid.empty: raise ValueError('Invalid Unit found. Allowed: '+', '.join(allowed_units))
                        dup=frame[frame['Item Name'].str.casefold().duplicated(keep=False)]
                        if not dup.empty: raise ValueError('Duplicate Item Name found in the uploaded Excel.')
                        existing={str(r['name']).strip().casefold() for r in items}
                        frame['Status']=frame['Item Name'].apply(lambda x:'Already Exists' if x.casefold() in existing else 'Ready')
                        st.dataframe(frame,hide_index=True,use_container_width=True)
                        ready=frame[frame['Status']=='Ready']
                        if st.button('SAVE ITEM MASTERLIST',type='primary',disabled=ready.empty,key='wh_master_save_v3'):
                            payload=[{'warehouse_id':warehouse_id,'category':category,'name':r['Item Name'],'unit':r['Unit'],'active':True} for _,r in ready.iterrows()]
                            _request('warehouse_items','POST',data=payload)
                            st.success(f'{len(payload)} item(s) added to {warehouse_name} / {category}.'); st.rerun()
                    except Exception as exc: st.error(str(exc))
        else:
            st.info('View only. Only Super Admin can add or upload Warehouse Item Masterlist items.')
        if items:
            st.dataframe(pd.DataFrame(items)[['name','unit']].rename(columns={'name':'Item Name','unit':'Unit'}),hide_index=True,use_container_width=True)
        else:
            st.info('No items yet for this warehouse/category.')


    with tabs[3]:
        st.subheader(f'Direct Dispatch — {warehouse_name}')
        st.caption('No branch order is required. Stock is reserved while IN TRANSIT and is deducted only after confirmed branch receiving.')

        if category in ('Pastries','Store Items'):
            st.info('Pastries and Store Items use the separate Buntun preparation/distribution workflow and are not posted through this regular Direct Dispatch screen.')
        elif not items:
            st.info('No items yet for this warehouse/category.')
        else:
            try:
                reserved_rows = _request('warehouse_reserved_stock', params={
                    'select':'item_id,reserved_quantity',
                    'warehouse_id':'eq.'+warehouse_id
                })
                reserved_lookup = {str(r['item_id']): float(r.get('reserved_quantity') or 0) for r in reserved_rows}
                balance_lookup = {str(r['item_id']): float(r.get('balance') or 0) for r in balances}

                item_by_name = {i['name']: i for i in items}
                mode_manual, mode_excel = st.tabs(['Manual Selection', 'Upload Allocation Excel'])

                with mode_manual:
                    selected_names = st.multiselect(
                        'Select Item(s) to Dispatch',
                        options=list(item_by_name.keys()),
                        placeholder='Choose only the items you want to dispatch',
                        key='wh_direct_selected_items'
                    )
                    branch = st.text_input('Branch', key='wh_direct_branch')
                    reference = st.text_input('Reference / DR No. (optional)', key='wh_direct_reference')
                    remarks = st.text_area('Remarks (optional)', key='wh_direct_remarks')
                    selected = pd.DataFrame()
                    invalid_qty = False

                    if selected_names:
                        dispatch_rows = []
                        for item_name in selected_names:
                            item = item_by_name[item_name]
                            item_id = str(item['id'])
                            available = balance_lookup.get(item_id, 0.0) - reserved_lookup.get(item_id, 0.0)
                            dispatch_rows.append({
                                'Item ID': item_id, 'Item': item['name'], 'Unit': item['unit'],
                                'Available Stock': available, 'Qty to Dispatch': 0.0
                            })
                        st.markdown('#### Selected Items')
                        edited_dispatch = st.data_editor(
                            pd.DataFrame(dispatch_rows), hide_index=True, use_container_width=True,
                            disabled=['Item ID','Item','Unit','Available Stock'],
                            column_config={'Qty to Dispatch': st.column_config.NumberColumn(
                                'Qty to Dispatch', min_value=0.0, step=1.0, format='%.3f')},
                            key='wh_direct_editor'
                        )
                        selected = edited_dispatch[pd.to_numeric(
                            edited_dispatch['Qty to Dispatch'], errors='coerce').fillna(0) > 0].copy()
                        invalid_qty = any(
                            float(r['Qty to Dispatch']) > float(r['Available Stock'])
                            for _, r in selected.iterrows()
                        ) if not selected.empty else False
                        if invalid_qty:
                            st.error('A dispatch quantity is greater than the available stock.')
                    else:
                        st.info('Select only the items you want to include. Unselected items are not part of the dispatch.')

                    if st.button('CREATE DIRECT DISPATCH', type='primary',
                                 disabled=(not str(branch).strip() or not selected_names or selected.empty or invalid_qty),
                                 key='wh_direct_create'):
                        dispatch_no = 'DD-' + datetime.now(TZ).strftime('%Y%m%d-%H%M%S-%f')
                        _rpc('warehouse_create_direct_dispatch', {
                            'p_dispatch_no': dispatch_no, 'p_warehouse_id': warehouse_id,
                            'p_branch': str(branch).strip(), 'p_category': category,
                            'p_reference': str(reference).strip(), 'p_remarks': str(remarks).strip(),
                            'p_rows': [{'item_id': str(r['Item ID']), 'qty': float(r['Qty to Dispatch'])}
                                       for _, r in selected.iterrows()],
                            'p_actor': username
                        })
                        st.success(f'Direct Dispatch {dispatch_no} created FOR PREPARATION.')
                        st.rerun()

                with mode_excel:
                    st.caption('Upload one allocation file for multiple branches. Uploading does not deduct or reserve stock until you confirm the allocation.')
                    template = pd.DataFrame([
                        {'Branch':'TEST BRANCH','Item Name':items[0]['name'],'Qty to Dispatch':10}
                    ])
                    template_buf = io.BytesIO()
                    with pd.ExcelWriter(template_buf, engine='openpyxl') as writer:
                        template.to_excel(writer, index=False, sheet_name='Direct Dispatch Allocation')
                    st.download_button(
                        'DOWNLOAD EXCEL TEMPLATE',
                        data=template_buf.getvalue(),
                        file_name=f'{warehouse_name}_{category}_Direct_Dispatch_Template.xlsx',
                        mime='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
                        key='wh_direct_template'
                    )

                    uploaded = st.file_uploader(
                        'Upload Allocation Excel',
                        type=['xlsx'],
                        key='wh_direct_excel'
                    )
                    if uploaded is not None:
                        try:
                            allocation = pd.read_excel(uploaded)
                            required = ['Branch','Item Name','Qty to Dispatch']
                            missing = [c for c in required if c not in allocation.columns]
                            if missing:
                                st.error('Missing column(s): ' + ', '.join(missing))
                            else:
                                allocation = allocation[required].copy()
                                allocation['Branch'] = allocation['Branch'].fillna('').astype(str).str.strip()
                                allocation['Item Name'] = allocation['Item Name'].fillna('').astype(str).str.strip()
                                allocation['Qty to Dispatch'] = pd.to_numeric(allocation['Qty to Dispatch'], errors='coerce')

                                allocation['Status'] = 'OK'
                                allocation.loc[allocation['Branch'].eq(''), 'Status'] = 'Missing Branch'
                                allocation.loc[allocation['Item Name'].eq(''), 'Status'] = 'Missing Item'
                                allocation.loc[allocation['Qty to Dispatch'].isna() | (allocation['Qty to Dispatch'] <= 0), 'Status'] = 'Invalid Qty'
                                allocation.loc[~allocation['Item Name'].isin(item_by_name.keys()), 'Status'] = 'Unknown Item'

                                dup = allocation.duplicated(['Branch','Item Name'], keep=False)
                                allocation.loc[dup & allocation['Status'].eq('OK'), 'Status'] = 'Duplicate Branch + Item'

                                valid_mask = allocation['Status'].eq('OK')
                                requested_by_item = allocation[valid_mask].groupby('Item Name')['Qty to Dispatch'].sum().to_dict()
                                for item_name, requested_qty in requested_by_item.items():
                                    item = item_by_name[item_name]
                                    item_id = str(item['id'])
                                    available = balance_lookup.get(item_id, 0.0) - reserved_lookup.get(item_id, 0.0)
                                    if float(requested_qty) > float(available):
                                        allocation.loc[
                                            allocation['Item Name'].eq(item_name) & allocation['Status'].eq('OK'),
                                            'Status'
                                        ] = f'Insufficient Stock (Available {available:g})'

                                st.markdown('#### Allocation Preview')
                                st.dataframe(allocation, hide_index=True, use_container_width=True)

                                has_errors = not allocation['Status'].eq('OK').all()
                                if has_errors:
                                    st.error('Fix the rows marked above before confirming the allocation.')
                                else:
                                    st.success('Allocation is valid. Nothing has been dispatched yet.')
                                    excel_reference = st.text_input(
                                        'Reference / Batch No. (optional)',
                                        key='wh_direct_excel_reference'
                                    )
                                    if st.button('CONFIRM EXCEL ALLOCATION', type='primary', key='wh_direct_excel_confirm'):
                                        batch_stamp = datetime.now(TZ).strftime('%Y%m%d-%H%M%S-%f')
                                        for branch_name, branch_rows in allocation.groupby('Branch', sort=False):
                                            payload = []
                                            for _, row in branch_rows.iterrows():
                                                item = item_by_name[row['Item Name']]
                                                payload.append({
                                                    'item_id': str(item['id']),
                                                    'qty': float(row['Qty to Dispatch'])
                                                })
                                            dispatch_no = f'DD-{batch_stamp}-{len(payload)}-{abs(hash(branch_name)) % 10000:04d}'
                                            _rpc('warehouse_create_direct_dispatch', {
                                                'p_dispatch_no': dispatch_no,
                                                'p_warehouse_id': warehouse_id,
                                                'p_branch': branch_name,
                                                'p_category': category,
                                                'p_reference': str(excel_reference).strip(),
                                                'p_remarks': 'Created from Direct Dispatch Excel allocation',
                                                'p_rows': payload,
                                                'p_actor': username
                                            })
                                        st.success('Excel allocation created FOR PREPARATION. Branch allocations are now saved separately.')
                                        st.rerun()
                        except Exception as exc:
                            st.error(f'Could not read allocation file: {exc}')

                st.markdown('#### For Preparation / In Transit')
                active_dispatches = _request('warehouse_direct_dispatches', params={
                    'select':'id,dispatch_no,branch,category,source,reference,remarks,status,created_by,created_at,released_by,released_at',
                    'warehouse_id':'eq.'+warehouse_id,
                    'category':'eq.'+category,
                    'status':'in.(FOR_PREPARATION,IN_TRANSIT,WITH_VARIANCE)',
                    'order':'created_at.desc',
                    'limit':'100'
                })

                if not active_dispatches:
                    st.info('No active Direct Dispatch records for this warehouse/category.')
                else:
                    item_names = {str(i['id']): i['name'] for i in items}
                    for d in active_dispatches:
                        with st.expander(f"{d['dispatch_no']} • {d['branch']} • {str(d['status']).replace('_',' ')}"):
                            lines = _request('warehouse_direct_dispatch_lines', params={
                                'select':'id,item_id,quantity_released,quantity_received,variance,receiving_remarks',
                                'dispatch_id':'eq.'+str(d['id']),
                                'order':'id.asc'
                            })
                            if lines:
                                st.dataframe(pd.DataFrame([{
                                    'Item': item_names.get(str(x['item_id']), str(x['item_id'])),
                                    'Released': x.get('quantity_released'),
                                    'Received': x.get('quantity_received'),
                                    'Variance': x.get('variance'),
                                    'Receiving Remarks': x.get('receiving_remarks')
                                } for x in lines]), hide_index=True, use_container_width=True)

                            st.caption(f"Reference: {d.get('reference') or '—'} • Created by: {d.get('created_by') or '—'}")

                            if d.get('status') == 'FOR_PREPARATION':
                                if st.button('MARK DISPATCHED / RELEASE TO BRANCH', key=f"wh_direct_release_{d['id']}"):
                                    _rpc('warehouse_release_direct_dispatch', {
                                        'p_dispatch_id': int(d['id']),
                                        'p_actor': username
                                    })
                                    st.success('Released. Stock is reserved / in transit. It has NOT been deducted yet.')
                                    st.rerun()
                            elif d.get('status') == 'IN_TRANSIT':
                                st.info('Waiting for branch receiving confirmation. Stock is reserved but not yet finally deducted.')
                            elif d.get('status') == 'WITH_VARIANCE':
                                st.warning('Branch receiving has a variance. Final completion/deduction requires variance resolution.')
            except Exception as exc:
                st.error(str(exc))

    with tabs[4]:
        st.subheader(f'Dispatch History — {warehouse_name}')
        st.caption('Final warehouse Pull Out appears only after branch receiving is confirmed. Active Direct Dispatch records remain in the Direct Dispatch tab.')
        try:
            item_rows=_request('warehouse_items',params={'select':'id,name,category','warehouse_id':'eq.'+warehouse_id})
            item_lookup={r['id']:r for r in item_rows}
            if item_lookup:
                rows=_request('warehouse_movements',params={
                    'select':'created_at,item_id,kind,quantity,reference,actor',
                    'kind':'eq.dispatch','order':'created_at.desc','limit':'500'
                })
                rows=[r for r in rows if r.get('item_id') in item_lookup]
                for r in rows:
                    meta=item_lookup[r['item_id']]; r['category']=meta['category']; r['item']=meta['name']
                if rows:
                    view=pd.DataFrame(rows).rename(columns={'created_at':'Date/Time','category':'Category','item':'Item','quantity':'Quantity','reference':'Reference','actor':'Processed By'})
                    st.dataframe(view[['Date/Time','Category','Item','Quantity','Reference','Processed By']],hide_index=True,use_container_width=True)
                else: st.info('No warehouse dispatches yet.')
            else: st.info('No items yet for this warehouse.')
        except Exception as exc: st.error(str(exc))

# When the entire order/delivery workflow has migrated to Supabase, call this
# from the dispatch transaction, NOT from an independent SQLite commit.
def dispatch_once(delivery_id, lines, actor):
    """Atomic, idempotent Supabase deduction; lines: [{item_id, qty}].
    Raises on insufficient stock, unknown items, or a changed retry payload.
    This is deliberately not connected to the legacy SQLite dispatch button.
    """
    return _rpc('warehouse_dispatch',{'p_delivery_id':str(delivery_id),'p_rows':lines,'p_actor':actor})
