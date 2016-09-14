# mspzap

Command line utility that zaps redundant .msp patches in the Windows Installer directory.

Based on [WiMsps.vbs by Microsoft's Heath Stewart](https://blogs.msdn.microsoft.com/heaths/2007/02/01/how-to-safely-delete-orphaned-patches/). See also [this post](https://www.raymond.cc/blog/safely-delete-unused-msi-and-mst-files-from-windows-installer-folder/).

Compatible with Python 2.7 and Python 3.2+.

```
usage: mspzap.py [-h] [--check] [--list] [--zap] [--move PATH]

Zap redundant .msp files in the Installer directory.

optional arguments:
  -h, --help   show this help message and exit
  --check      Count the redundant files and their total size.
  --list       List the redundant files and their sizes.
  --zap        Zap the files (admin required).
  --move PATH  Move the files to the specified directory (admin required).
```

### Disclaimer

This utility is provided without warranty of any kind. Use at your own risk.

If you want to be sure, do what I did and use `--move`, and delete the files later when you ensure that there's no damage.

### License

[WTFPL](http://www.wtfpl.net/).
However, if this is useful to you, I'd love it if you let me know :)
My email address is in my GitHub profile.
