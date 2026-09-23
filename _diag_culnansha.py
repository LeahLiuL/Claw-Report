#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""定位 CUL NANSHA 近 5 天所有邮件，按文件夹/主题/ReceivedTime 列明细。"""
import sys, os, re, datetime
import pythoncom
from win32com.client import Dispatch

TARGETS = ['CUL NANSHA', 'CULNansha', 'CUL_NANSHA', '2635W']
DAYS = 5

def norm(s):
    return re.sub(r'[^a-z0-9]', '', str(s).lower())

def walk(folder, path='', out=None, cutoff=None):
    if out is None: out = []
    cur = path + '/' + folder.Name if path else folder.Name
    try:
        items = folder.Items
        items.Sort("[ReceivedTime]", True)
        if cutoff:
            try:
                items = items.Restrict("[ReceivedTime] >= '%s'" % cutoff.strftime('%m/%d/%Y %I:%M %p'))
            except Exception as e:
                print('restrict err in', cur, e)
        for it in items:
            try:
                rt = it.ReceivedTime
            except Exception:
                continue
            if cutoff and datetime.datetime(rt.year, rt.month, rt.day, rt.hour, rt.minute) < cutoff:
                break
            subj = str(it.Subject) if it.Subject else ''
            try:
                snd = it.Sender.Name if it.Sender else ''
            except Exception:
                snd = ''
            try:
                smtp = it.Sender.Address if it.Sender else ''
            except Exception:
                smtp = ''
            flag = any(t.lower() in subj.lower() or t.lower() in snd.lower() for t in TARGETS)
            if flag or 'noon' in subj.lower():
                out.append({
                    'folder': cur,
                    'received': rt.strftime('%Y-%m-%d %H:%M'),
                    'sender_name': snd,
                    'sender_smtp': smtp,
                    'subject': subj,
                    'target_match': flag,
                })
    except Exception as e:
        print('err reading', cur, e)
    for f in folder.Folders:
        walk(f, cur, out, cutoff)
    return out

def main():
    pythoncom.CoInitialize()
    ns = Dispatch('Outlook.Application').GetNamespace('MAPI')
    print('mode=', ns.ExchangeConnectionMode)
    cutoff = datetime.datetime.now() - datetime.timedelta(days=DAYS)
    out = []
    inbox = ns.GetDefaultFolder(6)
    walk(inbox, '', out, cutoff)
    # also root
    root = ns.Folders.Item(1)
    if root.EntryID != inbox.StoreID:
        walk(root, '', out, cutoff)
    out.sort(key=lambda x: x['received'], reverse=True)
    print('\n总计近%d天含 CUL NANSHA/NOON 的邮件: %d 封' % (DAYS, len(out)))
    for x in out[:30]:
        mark = '**' if x['target_match'] else ''
        print('%s %s | %s | %s | %s' % (mark, x['received'], x['folder'], x['sender_name'], x['subject']))
    pythoncom.CoUninitialize()

if __name__ == '__main__':
    main()
