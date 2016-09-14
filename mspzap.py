# Based on the script from https://blogs.msdn.microsoft.com/heaths/2007/02/01/how-to-safely-delete-orphaned-patches/
# See also https://www.raymond.cc/blog/safely-delete-unused-msi-and-mst-files-from-windows-installer-folder/

from __future__ import division

import argparse
import os
import shutil
import subprocess
import sys
from contextlib import contextmanager
from functools import partial
from operator import itemgetter
from tempfile import NamedTemporaryFile

INSTALLER_PATH = os.path.join(os.environ['WINDIR'], 'Installer')
SCRIPT = b'''\
Option Explicit
Dim msi : Set msi = CreateObject("WindowsInstaller.Installer")
Dim products : Set products = msi.Products
Dim productCode
For Each productCode in products
    Dim patches : Set patches = msi.Patches(productCode)
    Dim patchCode
    For Each patchCode in patches
        Dim location : location = msi.PatchInfo(patchCode, "LocalPackage")
        WScript.Echo location
    Next
Next
'''


@contextmanager
def deleting(filename):
    try:
        yield
    finally:
        os.unlink(filename)


def get_all_patches():
    tf = NamedTemporaryFile(suffix='.vbs', delete=False)
    with deleting(tf.name):
        tf.write(SCRIPT)
        tf.close()

        output, _ = subprocess.Popen(['cscript', '//Nologo', tf.name], stdout=subprocess.PIPE).communicate()

    patches = output.decode().strip().split()
    patches_result = []
    for p in patches:
        path, filename = os.path.split(p)
        if path.lower() == INSTALLER_PATH.lower():
            patches_result.append(filename.lower())

    return patches_result


def get_msp_files_to_delete():
    patches = get_all_patches()
    msp_files_to_delete = []

    for fn in os.listdir(INSTALLER_PATH):
        fn = fn.lower()
        if fn.endswith('.msp') and fn not in patches:
            size = os.stat(os.path.join(INSTALLER_PATH, fn)).st_size
            msp_files_to_delete.append((fn, size))

    return msp_files_to_delete


def check(list_files=False):
    msp_files_to_delete = get_msp_files_to_delete()
    total_size = sum(size for fn, size in msp_files_to_delete) / (1 << 30)
    num_files = len(msp_files_to_delete)

    if list_files:
        print('The following files are safe to delete:')
        for fn, size in sorted(msp_files_to_delete, key=itemgetter(1)):
            print('{:8.2f} MB: {}'.format(size / (1 << 20), fn))

    print('Safe to delete {} files with total size {:.2f} GB'.format(num_files, total_size))


def move(move_path):
    if not os.path.isdir(move_path):
        sys.stderr.write('The path specified is not a valid directory.')
        return

    total_size = 0
    count = 0
    for fn, size in get_msp_files_to_delete():
        print('Moving file {} ({:.2f} MB)'.format(fn, size / (1 << 20)))
        shutil.move(os.path.join(INSTALLER_PATH, fn), move_path)

        total_size += size
        count += 1

    print('Moved {} files with total size {:.2f} GB'.format(count, total_size / (1 << 30)))


def zap():
    total_size = 0
    count = 0
    for fn, size in get_msp_files_to_delete():
        print('Deleting file {} ({:.2f} MB)'.format(fn, size / (1 << 20)))
        os.unlink(os.path.join(INSTALLER_PATH, fn))

        total_size += size
        count += 1

    print('Deleted {} files with total size {:.2f} GB'.format(count, total_size / (1 << 30)))


def main():
    parser = argparse.ArgumentParser(description='Zap redundant .msp files in the Installer directory.')
    parser.add_argument('--check', action='store_const', dest='action', const=check,
                        help='Count the redundant files and their total size.')
    parser.add_argument('--list', action='store_const', dest='action', const=partial(check, list_files=True),
                        help='List the redundant files and their sizes.')
    parser.add_argument('--zap', action='store_const', dest='action', const=zap,
                        help='Zap the files.')
    parser.add_argument('--move', dest='move_path', metavar='PATH',
                        help='Move the files to the specified directory.')
    args = parser.parse_args()

    if args.action:
        args.action()
    elif args.move_path:
        move(args.move_path)
    else:
        parser.error('You must specify an action (check, zap, move).')


if __name__ == '__main__':
    main()
