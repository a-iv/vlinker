# -*- coding: utf-8 -*-
import os
import sys
import sqlite3
import win32con, winioctlcon, winnt, win32file
import win32com.client
import struct
import pywintypes

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

def create_link(connection, filename, link):
    try:
        symlink(filename, link)
    except:
        return False
    else:
        cursor = connection.cursor()
        cursor.execute('INSERT OR REPLACE INTO files (filename, link) VALUES (?, ?);', (filename, link))
        connection.commit()
    return True

def parse_dir_name(path):
    if not path:
        return path
    path = os.path.abspath(path)
    if path.endswith('/'):
        return path[:-1]
    return path
      
def set_folders(connection):
    cursor = connection.cursor()
    cursor.execute('DROP TABLE IF EXISTS folders;')
    cursor.execute('CREATE TABLE folders (folder TEXT, is_target INTEGER);')
    while True:
        print u'Введите путь до общей папки назначения:'
        target = parse_dir_name(raw_input().decode('cp866'))
        if not os.path.exists(target):
            try:
                os.makedirs(target)
            except:
                continue
            else:
                break
        else:
            break
    cursor.execute('INSERT INTO folders (folder, is_target) VALUES (?, 1);', (target, ))
    sources = []
    while True:
        print u'Введите путь для добавления файлов (пустая строка для окончания ввода):'
        source = parse_dir_name(raw_input().decode('cp866'))
        if source == '':
            if sources:
                break
            print u'Необходимо ввысти хотябы один путь до папки с файлами'
            continue
        if not os.path.exists(source):
            print u'Папка не существует'
        else:
            sources.append(source)
            cursor.execute('INSERT INTO folders (folder, is_target) VALUES (?, 0);', (source, ))
            continue
    connection.commit()

def get_folders(connection):
    cursor = connection.cursor()
    cursor.execute('SELECT * FROM folders;')
    target = None
    sources = []
    for folder, is_target in cursor.fetchall():
        if is_target:
            target = folder
        else:
            sources.append(folder)
    return target, sources

connection = sqlite3.connect(os.path.expanduser('~\\.vlinker.sqlite'))
cursor = connection.cursor()
try:
    cursor.execute('SELECT * FROM folders;')
except sqlite3.OperationalError:
    set_folders(connection)
TARGET, SOURCES = get_folders(connection)
if TARGET is None or not SOURCES:
    set_folders(connection)
    TARGET, SOURCES = get_folders(connection)
   
if not os.path.exists(TARGET):
    print u'Не найдена папка назначения.'
    cursor.execute('DROP TABLE IF EXISTS files;')
    os.makedirs(TARGET)
try:
    cursor.execute('CREATE TABLE files(filename TEXT PRIMARY KEY, link TEXT);')
except sqlite3.OperationalError:
    print u'Поиск изменений...'
    cursor.execute('SELECT * FROM files;')
    files = cursor.fetchall()
    links = {}
    for filename, link in files:
        links[filename] = link
    found = 0
    move = []
    for dirpath, dirnames, filenames in os.walk(TARGET):
        for filename in filenames:
            target_link = os.path.join(dirpath, filename)
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
    print u'Найдено файлов: %d, переименовано: %d, удалено: %d.' % (found, len(move), len(links))
    print
    if move:
        print u'Файлы для переименования (%d):' % len(move)
        for source_filename, source_link, target_filename, target_link in move:
            print source_filename.encode('cp866', 'replace'), '->', target_filename.encode('cp866', 'replace')
        print
        while True:
            print u'Вы действительно хотите переиненовать эти файлы (yes/no)?',
            confirm = raw_input()
            if confirm.lower() in ['y', 'n', 'yes', 'no']:
                break
        if confirm.lower() in ['y', 'yes']:
            skipped = 0
            print
            for source_filename, source_link, target_filename, target_link in move:
                cursor.execute('DELETE FROM files WHERE filename = ?;', (source_filename, ))
                connection.commit()
                print source_filename.encode('cp866', 'replace'),
                error = False
                try:
                    os.remove(target_link)
                except:
                    print u'- ошибка удаления ссылки.'
                else:
                    try:
                        os.rename(source_filename, target_filename)
                    except:
                        skipped += 1
                        print u'- ошибка переименования',
                        if create_link(connection, source_filename, target_link):
                            print u'- пропущен.'
                        else:
                            print u'- ошибка восстановления старой ссылки.'
                    else:
                        if create_link(connection, target_filename, target_link):
                            print u'- готово.'
                        else:
                            print u'- ошибка создания новой ссылки.'
            print u'Перемещено: %d, пропущено: %d.' % (len(move) - skipped, skipped)
            print
    if links:
        print u'Файлы для удаления (%d):' % len(links)
        for filename, link in links.iteritems():
            print filename.encode('cp866', 'replace')
        print
        while True:
            print u'Вы действительно хотите удалить эти файлы (yes/no)?',
            confirm = raw_input()
            if confirm.lower() in ['y', 'n', 'yes', 'no']:
                break
        if confirm.lower() in ['y', 'yes']:
            skipped = 0
            for filename, link in links.iteritems():
                print filename.encode('cp866', 'replace'),
                cursor.execute('DELETE FROM files WHERE filename = ?;', (filename, ))
                connection.commit()
                try:
                    os.remove(filename)
                except:
                    skipped += 1
                    print u'- ошибка удаления.'
                else:
                    print u'- готово.'
            print u'Удалено: %d, пропущено: %d.' % (len(links) - skipped, skipped)
            print
else:
    print u'Поиск новых файлов...'
    found = 0
    skipped = 0
    files = []
    for source in SOURCES:
        for dirpath, dirnames, filenames in os.walk(source):
            for filename in filenames:
                filename = os.path.join(dirpath, filename)
                print filename.encode('cp866', 'replace'),
                if filename in files:
                    print u'- уже добавлен.'
                    continue
                files.append(filename)
                name, ext = os.path.splitext(os.path.split(filename)[1])
                link = os.path.join(TARGET, name + ext)
                if os.path.exists(link):
                    index = 2
                    while True:
                        link = os.path.join(TARGET, '%s-%d%s' % (name, index, ext))
                        if not os.path.exists(link):
                            break
                        index += 1
                if create_link(connection, filename, link):
                    print u'- готово.'
                    found += 1
                else:
                    print u'- ошибка создания ссылки.'
                    skipped += 1
    connection.commit()
    print u'Найдено файлов: %d, пропущено: %d' % (found, skipped)
    print
connection.close()
print u'Нажмите Enter для завершения...',
raw_input()
