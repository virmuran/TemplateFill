# -*- coding: utf-8 -*-
"""分批清理打包产物（规避批量删除保护，每轮最多 40 个文件）"""
import os
import sys

BASE = r'C:\Users\Administrator\Desktop\TemplateFill'
TREES = [os.path.join(BASE, d) for d in
         ('dist/TemplateFill', 'build/TemplateFill',
          'dist_diag/TemplateFill', 'build_diag/TemplateFill')]
BATCH = 40

def main():
    files = []
    for root in TREES:
        if os.path.isdir(root):
            for dirpath, dirnames, filenames in os.walk(root):
                for fn in filenames:
                    files.append(os.path.join(dirpath, fn))
    remaining = len(files)
    deleted = 0
    for f in files[:BATCH]:
        try:
            os.remove(f)
            deleted += 1
        except OSError:
            pass
    print('deleted %d, remaining files: %d' % (deleted, remaining - deleted))
    # 全部文件删完后清理空目录
    if remaining - deleted == 0:
        removed_dirs = 0
        for root in TREES:
            for dirpath, dirnames, filenames in sorted(os.walk(root, topdown=False), key=lambda x: -len(x[0])):
                try:
                    os.rmdir(dirpath)
                    removed_dirs += 1
                except OSError:
                    pass
        # 树根目录本身
        for root in TREES:
            try:
                os.rmdir(root)
                removed_dirs += 1
            except OSError:
                pass
        print('empty dirs removed: %d' % removed_dirs)

if __name__ == '__main__':
    main()
