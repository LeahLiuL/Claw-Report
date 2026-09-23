#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""遍历 outlook 里所有 store + 所有文件夹，按主题/发件人关键词找 CUL NANSHA 的 NOON 邮件。"""
import sys, re, datetime
import pythoncom
from win32com.client import Dispatch

KEYWORDS = ['NOON RPT', 'NOON REPORT', 'CUL NANSHA', '2635W']

def walk(folder, path='', out=None, cutoff=None):
    if out is None: out = []
    cur = (path + ' / ' + folder.Name) if path else folder.Name
    try:
        items = folder.Items
        try:
            items.Sort('[ReceivedTime]', True)
        except Exception:
            pass
        if cutoff:
            try:
                items = items.Restrict("[ReceivedTime] >= '%s'" % cutoff.strftime('%m/%d/%Y %I:%M %p'))
            except Exception:
                pass
        for it in items:
            try:
                rt = it.ReceivedTime
            except Exception:
                continue
            try:
                subj = str(it.Subject) if it.Subject else ''
            except Exception:
                subj = ''
            if any(k.lower() in subj.lower() for k in KEYWORDS):
                try:
                    snd = it.Sender.Name if it.Sender else ''
                except Exception:
                    snd = ''
                try:
                    addr = it.Sender.Address if it.Sender else ''
                except Exception:
                    addr = ''
                out.append({
                    'folder': cur,
                    'received': rt.strftime('%Y-%m-%d %H:%M'),
                    'sender': snd,
                    'addr': addr,
                    'subject': subj,
                })
    except Exception as e:
        pass  # 联系人/日历等无 ReceivedTime，跳过
    for f in folder.Folders:
        walk(f, cur, out, cutoff)
    return out

def main():
    pythoncom.CoInitialize()
    ns = Dispatch('Outlook.Application').GetNamespace('MAPI')
    print('ExchangeConnectionMode =', ns.ExchangeConnectionMode)
    cutoff = datetime.datetime.now() - datetime.timedelta(days=6)
    all_rows = []
    nstores = ns.Folders.Count
    print('store count =', nstores)
    for i in range(1, nstores + 1):
        store = ns.Folders.Item(i)
        sname = store.Name
        print('  store[%d] = %s' % (i, sname))
        rows = walk(store, sname, None, cutoff)
        all_rows.extend(rows)
    all_rows.sort(key=lambda x: x['received'], reverse=True)
    print('\n近6天含 CUL NANSHA/NOON RPT/2635W 的邮件: %d 封\n' % len(all_rows))
    for x in all_rows[:60]:
        print('%s | %s | %s <%s> | %s' % (x['received'], x['folder'], x['sender'], x['addr'], x['subject']))
    pythoncom.CoUninitialize()

if __name__ == '__main__':
    main()
