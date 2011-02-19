# -*- coding: utf-8 -*-

import os
import sys
import sqlite3
import win32con, winioctlcon, winnt, win32file
import win32com.client
import struct
import pywintypes

SOURCES = [r'D:\projects\vlinker\1', r'D:\projects\vlinker\2', r'D:\projects\vlinker\3']
TARGET = r'D:\projects\vlinker\result'

__CSL = None
def symlink(source, link_name):
    global __CSL
    if __CSL is None:
        import ctypes
        csl = ctypes.windll.kernel32.CreateSymbolicLinkW
        csl.argtypes = (ctypes.c_wchar_p, ctypes.c_wchar_p, ctypes.c_uint32)
        csl.restype = ctypes.c_ubyte
        __CSL = csl
    flags = 0
    if source is not None and os.path.isdir(source):
        flags = 1
    if __CSL(link_name, source, flags) == 0:
        raise ctypes.WinError()

def _symlink(source, filename):
    shell = win32com.client.Dispatch("WScript.Shell")
    shortcut = shell.CreateShortCut(filename + '.lnk')
    shortcut.Targetpath = source
    shortcut.save()

def get_reparse_target(fname):
    h = win32file.CreateFile(fname, 0,
        win32con.FILE_SHARE_READ|win32con.FILE_SHARE_WRITE|win32con.FILE_SHARE_DELETE, None,
        win32con.OPEN_EXISTING,
        win32file.FILE_FLAG_OVERLAPPED|win32file.FILE_FLAG_OPEN_REPARSE_POINT|win32file.FILE_FLAG_BACKUP_SEMANTICS, 0)

    output_buf=win32file.AllocateReadBuffer(winnt.MAXIMUM_REPARSE_DATA_BUFFER_SIZE)
    buf=win32file.DeviceIoControl(h, winioctlcon.FSCTL_GET_REPARSE_POINT, None,
            OutBuffer=output_buf, Overlapped=None)
    fixed_fmt='LHHHHHHL'
    fixed_len=struct.calcsize(fixed_fmt)
    tag, datalen, reserved, target_offset, target_len, printname_offset, printname_len, wchar = \
        struct.unpack(fixed_fmt, buf[:fixed_len])

    ## variable size target data follows the fixed portion of the buffer 
    name_buf=buf[fixed_len:]

    target_buf=name_buf[target_offset:target_offset+target_len]
    target=target_buf.decode('utf-16-le')
    return target[4:]

def _get_reparse_target(filename):
    shell = win32com.client.Dispatch("WScript.Shell")
    shortcut = shell.CreateShortCut(filename + '.lnk')
    return shortcut.Targetpath

def create_link(cursor, filename, link):
    try:
        symlink(filename, link)
    except:
        return False
    else:
        cursor.execute('INSERT OR REPLACE INTO files(filename, link) VALUES (?, ?);', (filename, link))
        connection.commit()
    return True


if not os.path.exists(TARGET):
    os.makedirs(TARGET)
connection = sqlite3.connect(os.path.join(TARGET, 'vlinker.sqlite'))
cursor = connection.cursor()
try:
    cursor.execute('CREATE TABLE files(filename TEXT PRIMARY KEY, link TEXT);')
except sqlite3.OperationalError:
    print 'Поиск изменений...'
    cursor.execute('SELECT * FROM files;')
    files = cursor.fetchall()
    links = {}
    for filename, link in files:
        links[filename] = link
    found = 0
    move = []
    for dirpath, dirnames, filenames in os.walk(TARGET):
        for filename in filenames:
            target_link = os.path.join(dirpath, filename).decode('cp1251')
            try:
                source_filename = get_reparse_target(target_link)
            except pywintypes.error:
                pass
            else:
                found += 1
                if source_filename not in links:
                    continue
                source_link = links[source_filename]
                del links[source_filename]
                if source_link == target_link:
                    continue
                path, name = os.path.split(source_filename)
                target_filename = os.path.join(path, os.path.split(target_link)[1])
                move.append((source_filename, source_link, target_filename, target_link))
    print 'Found: %d, changed: %d, removed: %d.' % (found, len(move), len(links))
    print
    if move:
        print 'Files to be moved (%d):' % len(move)
        for source_filename, source_link, target_filename, target_link in move:
            print source_filename, '->', target_filename
        print
        while True:
            confirm = raw_input('Do you realy want to rename this files (yes/no)? ')
            if confirm.lower() in ['y', 'n', 'yes', 'no']:
                break
        if confirm.lower() in ['y', 'yes']:
            skipped = 0
            print
            for source_filename, source_link, target_filename, target_link in move:
                cursor.execute('DELETE FROM files WHERE filename = ?;', (source_filename, ))
                connection.commit()
                print source_filename,
                error = False
                try:
                    os.remove(target_link)
                except:
                    print '- fail to remove link.'
                else:
                    try:
                        os.rename(source_filename, target_filename)
                    except:
                        skipped += 1
                        print '- fail to rename',
                        if create_link(cursor, source_filename, target_link):
                            print '- skipped.'
                        else:
                            print '- fail to create old link.'
                    else:
                        if create_link(cursor, target_filename, target_link):
                            print '- done.'
                        else:
                            print '- fail to create new link.'
            print 'Moved: %d, skipped: %d.' % (len(move) - skipped, skipped)
            print
    if links:
        print 'Files to be removed (%d):' % len(links)
        for filename, link in links.iteritems():
            print filename
        print
        while True:
            confirm = raw_input('Do you realy want to remove this files (yes/no)? ')
            if confirm.lower() in ['y', 'n', 'yes', 'no']:
                break
        if confirm.lower() in ['y', 'yes']:
            skipped = 0
            for filename, link in links.iteritems():
                print filename,
                cursor.execute('DELETE FROM files WHERE filename = ?;', (filename, ))
                connection.commit()
                try:
                    os.remove(filename)
                except:
                    skipped += 1
                    print '- fail to remove.'
                else:
                    print '- ok.'
            print 'Removed: %d, skipped: %d.' % (len(links) - skipped, skipped)
            print
    raw_input('Press enter to continue...')
else:
    print 'Search for files...'
    found = 0
    skipped = 0
    for source in SOURCES:
        for dirpath, dirnames, filenames in os.walk(source):
            for filename in filenames:
                filename = os.path.join(dirpath, filename).decode('cp1251')
                name, ext = os.path.splitext(os.path.split(filename)[1])
                link = os.path.join(TARGET, name + ext)
                if os.path.exists(link):
                    index = 2
                    while True:
                        link = os.path.join(TARGET, '%s-%d%s' % (name, index, ext))
                        if not os.path.exists(link):
                            break
                        index += 1
                if create_link(cursor, filename, link):
                    found += 1
                else:
                    print filename, '- fail to create link.'
                    skipped += 1
    connection.commit()
    print 'Found: %d, skipped: %d' % (found, skipped)
    raw_input('Press enter to continue...')
connection.close()
