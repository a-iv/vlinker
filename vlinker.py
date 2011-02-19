import os
import sys
import sqlite3
import win32con, winioctlcon, winnt, win32file 
import struct 

SOURCES = [r'C:\.test\1', r'C:\.test\2', r'C:\.test\3']
DESTINATION = r'c:\.test\dest'

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

def get_reparse_target(fname):
    h = win32file.CreateFile(fname, 0,
        win32con.FILE_SHARE_READ|win32con.FILE_SHARE_WRITE|win32con.FILE_SHARE_DELETE, None,
        win32con.OPEN_EXISTING,
        win32file.FILE_FLAG_OVERLAPPED|win32file.FILE_FLAG_OPEN_REPARSE_POINT|win32file.FILE_FLAG_BACKUP_SEMANTICS, 0)

    output_buf=win32file.AllocateReadBuffer(winnt.MAXIMUM_REPARSE_DATA_BUFFER_SIZE)
    buf=win32file.DeviceIoControl(h, winioctlcon.FSCTL_GET_REPARSE_POINT, None,
            OutBuffer=output_buf, Overlapped=None)
    fixed_fmt='LHHHHHH'
    fixed_len=struct.calcsize(fixed_fmt)
    tag, datalen, reserved, target_offset, target_len, printname_offset, printname_len = \
        struct.unpack(fixed_fmt, buf[:fixed_len])

    ## variable size target data follows the fixed portion of the buffer 
    name_buf=buf[fixed_len:]

    target_buf=name_buf[target_offset:target_offset+target_len]
    print target_buf
    target=target_buf.decode('utf-16-le')
    return target

if not os.path.exists(DESTINATION):
    os.makedirs(DESTINATION)
connection = sqlite3.connect(os.path.join(DESTINATION, 'vlinker.sqlite'))
cursor = connection.cursor()
try:
    cursor.execute('CREATE TABLE files(source TEXT PRIMARY KEY, target TEXT);')
except sqlite3.OperationalError:
    print 'Search for changes...'
    cursor.execute('SELECT * FROM files;')
    files = cursor.fetchall()
    for source, target in files:
        print source, get_reparse_target(target)
else:
    print 'Search for files...'
    for source in SOURCES:
        for dirpath, dirnames, filenames in os.walk(source):
            for filename in filenames:
                source = os.path.join(dirpath, filename)
                name, ext = os.path.splitext(os.path.split(source)[1])
                target = os.path.join(DESTINATION, name + ext)
                if os.path.exists(target):
                    index = 1
                    while True:
                        target = os.path.join(DESTINATION, '%s-%d%s' % (name, index, ext))
                        if not os.path.exists(target):
                            break
                        index += 1
                symlink(source, target)
                cursor.execute('INSERT OR REPLACE INTO files(source, target) VALUES (?, ?);', (source, target))
    connection.commit()
connection.close()
