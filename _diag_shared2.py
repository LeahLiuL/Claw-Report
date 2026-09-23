#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""拿 Vessel Report / CUL Fleet 的 SMTP, 并尝试用地址打开共享收件箱, 搜 CUL NANSHA 09-21+ NOON。"""
import datetime
import pythoncom
from win32com.client import Dispatch

def smtp_of(name):
    OL = Dispatch('Outlook.Application')
    NS = OL.GetNamespace('MAPI')
    r = NS.CreateRecipient(name)
    r.Resolve()
    if not r.Resolved:
        return None, None, 'unresolved'
    eu = r.AddressEntry
    smtp = ''
    typ = ''
    try:
        typ = eu.Type
    except Exception:
        pass
    try:
        xu = eu.GetExchangeUser()
        if xu:
            smtp = xu.PrimarySmtpAddress or ''
    except Exception as e:
        smtp = '(GetExchangeUser err %s)' % e
    return smtp, typ, 'ok'

def try_open(smtp):
    OL = Dispatch('Outlook.Application')
    NS = OL.GetNamespace('MAPI')
    try:
        r = NS.CreateRecipient(smtp)
        r.Resolve()
        f = NS.GetSharedDefaultFolder(r, 6)
        print('  共享收件箱打开成功, Count =', f.Items.Count)
        items = f.Items
        items.Sort('[ReceivedTime]', True)
        cutoff = datetime.datetime.now() - datetime.timedelta(days=4)
        rest = items.Restrict("[ReceivedTime] >= '%s'" % cutoff.strftime('%m/%d/%Y %I:%M %p'))
        hit = 0
        for it in rest:
            try:
                subj = str(it.Subject) if it.Subject else ''
            except Exception:
                subj = ''
            if 'CUL NANSHA' in subj.upper() or '2635W' in subj:
                try:
                    rt = it.ReceivedTime.strftime('%Y-%m-%d %H:%M')
                except Exception:
                    rt = '?'
                print('    ', rt, '|', subj[:80])
                hit += 1
        print('  命中 CUL NANSHA/2635W 邮件:', hit)
    except Exception as e:
        print('  打开失败:', e)

def main():
    pythoncom.CoInitialize()
    OL = Dispatch('Outlook.Application')
    NS = OL.GetNamespace('MAPI')
    print('ExchangeConnectionMode =', NS.ExchangeConnectionMode)
    for name in ('Vessel Report', 'CUL Fleet'):
        print('\n=== %s ===' % name)
        smtp, typ, st = smtp_of(name)
        print('  status=%s  type=%s  smtp=%s' % (st, typ, smtp))
        if smtp and 'err' not in smtp and '@' in smtp:
            try_open(smtp)
    pythoncom.CoUninitialize()

if __name__ == '__main__':
    main()
