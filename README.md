# PictureMagic

A nice little helper script to keep my iPhone pictures sorted, after syncing them to my PC.

~~~
$ python picture_magic.py --help

usage: picture_magic.py [-h] --mode {0,1,2,3,4,5,6,7} --path PATH [--dry_run] [--verbose]

Organize pictures from iOS

options:
  -h, --help            show this help message and exit
  --mode {0,1,2,3,4,5,6,7}
                        mode of operation.
                          0:show folder statistics (recursive, read-only)
                          1:show capture timestamps grouped by year (recursive, read-only)
                          2:find duplicate files and remove them interactively (recursive)
                          3:move to type subfolders
                          4:move back from type subfolders (not possible if names clash)
                          5:move to monthly subfolders
                          6:safely move back from any direct subfolders (renaming if names clash)
                          7:safely remove renaming-suffix from previous safe moving (recursive)
  --path PATH           path to a folder.
  --dry_run             If true, also the non-READ-ONLY modes will not make any changes and only output the planned operations.
  --verbose             If true, print verbose log information.


