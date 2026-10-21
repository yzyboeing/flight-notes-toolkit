#!/usr/bin/env python3
"""drive_verify.py —— 谷歌云盘上传复查（总纲第十节第 11 步；2026-10-14 丢 11 个文件的教训：复制时 md5 对不算数）。

用法：python3 drive_verify.py <文件或文件夹> [...]
  对每个本地文件，查 DriveFS 元数据库里同名、非 local- 开头（已上传）且大小一致的条目；文件夹逐个查其中的文件（不递归）。
  数据库连同 -wal / -shm 复制到临时目录再读（不碰正在用的库）。全部确认退出码 0；有未确认的退出码 1——过几分钟再跑。
  不同步时先 pgrep -fl "Google Drive"；local- 计数长时间不动就退出重开 Google Drive。"""
import glob, os, shutil, sqlite3, sys, tempfile

ROOT = os.path.expanduser('~/Library/Application Support/Google/DriveFS')


def dbs():
    tmp = tempfile.mkdtemp(prefix='drivefs-')
    out = []
    for i, d in enumerate(sorted(glob.glob(os.path.join(ROOT, '*', 'metadata_sqlite_db')))):
        for s in ('', '-wal', '-shm'):
            if os.path.exists(d + s): shutil.copy2(d + s, os.path.join(tmp, 'db%d%s' % (i, s)))
        out.append(sqlite3.connect(os.path.join(tmp, 'db%d' % i)))
    return out


def main():
    files = []
    for a in sys.argv[1:]:
        if os.path.isdir(a): files += [os.path.join(a, f) for f in sorted(os.listdir(a)) if os.path.isfile(os.path.join(a, f)) and not f.startswith('.')]
        elif os.path.isfile(a): files.append(a)
        else: print('-- 本地没有：', a)
    cons = dbs()
    if not cons: sys.exit('找不到 DriveFS 元数据库（%s）' % ROOT)
    pend = sum(c.execute("select count(*) from items where id like 'local-%'").fetchone()[0] for c in cons)
    bad = 0
    for f in files:
        n, sz = os.path.basename(f), os.path.getsize(f)
        rows = [r for c in cons for r in c.execute('select id, file_size from items where local_title=? and is_folder=0', (n,)).fetchall()]
        ok = any(not i.startswith('local-') and s == sz for i, s in rows)
        bad += not ok
        print('OK ' if ok else '-- ', n, sz)
    print('云盘复查：%d 个文件，未确认 %d；待上传（local-）共 %d 项' % (len(files), bad, pend))
    sys.exit(1 if bad else 0)


if __name__ == '__main__':
    main()
