#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""尝试解析 'Vessel Report' / 'CUL Fleet' 收件人, 看是否为可访问的共享邮箱, 并搜其中
最近 3 天含 CUL NANSHA NOON 的邮件。"""
import datetime
import pythoncom
from win32com.client import Dispatch

def main():
    pythoncom.CoInitialize()
    OL = Dispatch('Outlook.Application')
    NS = OL.GetNamespace('MAPI')
    print('ExchangeConnectionMode =', NS.ExchangeConnectionMode)

    # 当前 profile 里挂的所有 store
    print('\n--- 当前 profile 的所有 store ---')
    for s in NS.Stores:
        print('  store:', repr(s.DisplayName))

    for name in ('Vessel Report', 'CUL Fleet', 'vesselreport', 'culfleet'):
        try:
            r = NS.CreateRecipient(name)
            r.Resolve()
            if r.Resolved:
                eu = r.AddressEntry
                print('\n[解析成功] %s -> %s' % (name, eu.Address))
                try:
                    print('  Type:', eu.Type, 'Name:', eu.Name)
                except Exception as e:
                    print('  type err', e)
                try:
                    folder = NS.GetSharedDefaultFolder(r, 6)  # 6=收件箱
                    print('  共享收件箱 Count =', folder.Items.Count)
                    # 搜 CUL NANSHA NOON
                    items = folder.Items
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
                            print('    ', rt, '|', subj)
                            hit += 1
                    print('  命中 CUL NANSHA/2635W 邮件:', hit)
                except Exception as e:
                    print('  GetSharedDefaultFolder 失败:', e)
            else:
                print('\n[未解析] %s' % name)
        except Exception as e:
            print('\n[异常] %s -> %s' % (name, e))

    pythoncom.CoUninitialize()

if __name__ == '__main__':
    main()
