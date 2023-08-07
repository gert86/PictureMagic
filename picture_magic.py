import argparse
from genericpath import isfile
import os
import shutil
import re
import collections

from glob import glob
from pathlib import Path
from typing import Collection

from exif import Image
import exifread
import exiftool
import fnmatch
import platform


import pytz
import datetime
#import win32com


class PictureMagic(object):

    # reserved folder names when sorting by type
    dirname_aae       = "_EditDataAAE"
    dirname_downloads = "_WhatsApp,Downloads,etc"
    dirname_live_imgs = "_LiveImages"
    dirname_originals = "_Originals"
    dirname_remaining = "Pics_and_Movies"
    special_folders = [dirname_aae, dirname_downloads, dirname_live_imgs, dirname_originals, dirname_remaining]

    def main(self):
        mode_map = {0: 'show folder statistics (recursive, read-only)',
                    1: 'show capture timestamps grouped by year (recursive, read-only)',
                    2: 'find duplicate file names (recursive, read-only)',
                    3: 'move to type subfolders',
                    4: 'move back from type subfolders (not possible if names clash)',
                    5: 'move to monthly subfolders',
                    6: 'safely move back from any direct subfolders (renaming if names clash)',
                    7: 'safely remove renaming-suffix from previous safe moving (recursive)',
                    }

        dict_to_str = lambda x: ','.join(" %s:%s" % (str(k), str(v)) for (k, v) in x.items()) if isinstance(x,dict) else x
        parser = argparse.ArgumentParser(description='Organize pictures from iOS')
        parser.add_argument('--mode', dest='mode', required=True, type=int, choices=mode_map.keys(),
                           help=f'mode of operation. {dict_to_str(mode_map)}')
        parser.add_argument('--path', dest='path', required=True, type=str,
                           help='path to a folder.')
        parser.add_argument('--dry_run', dest='dry_run', required=False, action='store_true',
                           help='If true, also the non-READ-ONLY modes will not make any changes and '
                                'only output the planned operations.')
        parser.add_argument('--verbose', dest='verbose', required=False, action='store_true',
                           help='If true, print verbose log information.')

        args = parser.parse_args()
        print("Running program in mode: {}".format(args.mode))
        if not os.path.isdir(args.path):
            print("The given folder {} does not exist. Exit.".format(args.path))
            return

        if args.mode==0:
            self.showStats(args)
        elif args.mode == 1:
            self.showCaptureYears(args)
        elif args.mode == 2:
            self.findDuplicates(args, True)            
        elif args.mode == 3:
            self.moveToSubfolders(args)
        elif args.mode == 4:
            self.moveBackFromSubfolders(args)
        elif args.mode == 5:
            self.moveToMonthlySubfolders(args)
        elif args.mode == 6:
            self.safeMoveFromSubfolders(args)
        elif args.mode == 7:
            self.removeRenamingSuffixes(args)                         

    ###############################################################################################################
    # Mode 0 - show file statistics (READ-ONLY)
    def showStats(self, args):
        folder_path = Path(args.path.rstrip(os.sep))
        files = [path.suffix for path in folder_path.glob("**/*") if path.is_file() and path.suffix]
        data = collections.Counter(files)
        for key, val in data.items():
            print(f'{key}: {val}')
        print(f"-----------------------------")
        print(f"Overall: {len(files)} files")


    ###############################################################################################################
    # Mode 1 - show capture year  (READ-ONLY)
    # Checks all image and video files in a folder for their timestamp and prints the year distribution.
    def showCaptureYears(self, args):               
        folder_path = Path(args.path.rstrip(os.sep))
        collection = collections.defaultdict(list)
        for file_path in Path(folder_path).glob('**/*'):
            dt = None
            lower_name = str(file_path).lower()
            if lower_name.endswith("jpg") or lower_name.endswith("jpeg") or lower_name.endswith("png"):
                dt = self.getImageCaptureTimestamp(file_path)
            elif lower_name.endswith("mov") or lower_name.endswith("mp4"):
                dt = self.getVideoCaptureTimestamp(file_path)
            
            key = str(dt[0:4]) if (dt is not None and len(dt)>7) else 'Unknown'
            collection[key].append(file_path)

        for year,files in collection.items():
            print(f"{year}: {len(files)} files")        


    ###############################################################################################################
    # Mode 2 - find duplicates (READ-ONLY)
    # Also returns duplicates in the form {filename: num_occurrences}
    def findDuplicates(self, args, doPrint): 
        folder_path = Path(args.path.rstrip(os.sep))
        dups = {}
        files = [os.path.basename(path) for path in folder_path.glob("**/*") if path.is_file()]
        data = collections.Counter(files)
        for key, val in data.items():
            if (val > 1):
                dups[key] = val
                if doPrint:
                    print(f'Duplicate file name: {key} -> occurs {val} times')
        if doPrint:
            print(f'Found {len(dups)} duplicates')                       
        return dups


    ###############################################################################################################
    # Mode 3 - move to type subfolders
    def moveToSubfolders(self, args):
        folder_path = Path(args.path.rstrip(os.sep))
        if not self.checkIfDirWithNonReservedName(folder_path):
            return

        # TODO: Since we don't really move in dry_run, the number of matching files is not equal to normal mode
        #       if a file matches multiple criteriums
        # TODO: For performance reasons it would be way better if we only iterate once over all files.
        #       If we classify each file in one run (criterium after criterium) and store what is going to
        #       happen with each file we could avoid this problem AND the one above.

        # Folder: _EditDataAAE
        # All files with type AAE (*.AAE)
        criterion_aae = lambda curr_file, curr_folder: True if curr_file.lower().endswith('aae') else False
        self.createSubfolderAndMove(args.dry_run, args.verbose, folder_path, self.dirname_aae, criterion_aae)

        # Folder: _WhatsApp, Downloads, etc
        # All files that do not start with IMG_*  OR
        # that have a type other than JPG, JPEG or MOV
        # TODO: JPG files named IMG* that have an empty "Date taken" (WhatsApp clears all EXIF data)
        criterion_downloads = lambda curr_file, curr_folder: True if not re.match(r"^img_.*\.(jpg|jpeg|mov)$", curr_file, re.IGNORECASE) else False
        self.createSubfolderAndMove(args.dry_run, args.verbose, folder_path, self.dirname_downloads, criterion_downloads)

        # Folder: _LiveImages
        # All MOV files for which a corresponding JPG file (with the same name) exists.
        # --> IMG_0433.MOV, IMG_E0910.MOV
        def criterion_live_imgs(curr_file, curr_folder):
            suffix = Path(curr_file).suffix
            if suffix.lower() == '.mov':
                for img_suffix in ['.jpg', '.jpeg', '.JPG', '.JPEG']:
                    potential_twin_file = img_suffix.join(curr_file.rsplit(suffix, 1)) # replace only last occurrence of .MOV with .JPG
                    if os.path.isfile(os.path.join(curr_folder, potential_twin_file)):
                        return True
            return False
        self.createSubfolderAndMove(args.dry_run, args.verbose, folder_path, self.dirname_live_imgs, criterion_live_imgs)


        # TODO: This is not so efficient, since we check all IMG_ files for matching IMG_E instead vice versa.
        #       Also no warnings for IMG_E without matching Original is possible.
        # Folder: _Originals
        # Contains all files IMG_<num>.xxx for which an edited version IMG_E<num>.xxx exists.
        def criterion_originals(curr_file, curr_folder):
            if curr_file.startswith('IMG_'):
                potential_twin_file = 'IMG_E'.join(curr_file.split('IMG_', 1))  # replace only first occurrence of IMG_ with IMG_E
                if os.path.isfile(os.path.join(curr_folder, potential_twin_file)):
                    return True
            return False
        self.createSubfolderAndMove(args.dry_run, args.verbose, folder_path, self.dirname_originals, criterion_originals)

        # # Folder: Pics_and_Movies
        # # All remaining files. Should be only JPG, JPEG or MOV, where MOV is NOT a live video.
        criterion_remaining = lambda curr_file, curr_folder: True if re.match(r"^img_.*\.(jpg|jpeg|mov)$", curr_file, re.IGNORECASE) else False
        self.createSubfolderAndMove(args.dry_run, args.verbose, folder_path, self.dirname_remaining, criterion_remaining)


    ###############################################################################################################
    # Mode 4 - move back from type subfolders to revert mode 3 operations
    def moveBackFromSubfolders(self, args):
        folder_path = Path(args.path.rstrip(os.sep))
        if not self.checkIfDirWithNonReservedName(folder_path):
            return
        
        # assure that moving won't cause name clashes
        filename_duplicates = self.findDuplicates(args, False)
        if len(filename_duplicates) > 0:
            print("Cannot move back to parent folder, because subfolders contain files with equal names. Use mode 6 to move manually.")
            for key, val in filename_duplicates.items():
                print(f'Duplicate file name: {key} -> occurs {val} times')   
            return         

        self.moveToParentAndDeleteSubfolder(args.dry_run, args.verbose, folder_path, self.dirname_aae)
        self.moveToParentAndDeleteSubfolder(args.dry_run, args.verbose, folder_path, self.dirname_downloads)
        self.moveToParentAndDeleteSubfolder(args.dry_run, args.verbose, folder_path, self.dirname_live_imgs)
        self.moveToParentAndDeleteSubfolder(args.dry_run, args.verbose, folder_path, self.dirname_originals)
        self.moveToParentAndDeleteSubfolder(args.dry_run, args.verbose, folder_path, self.dirname_remaining)
        print("Done!")

    ###############################################################################################################
    # Mode 5 - move to monthly subfolders
    def moveToMonthlySubfolders(self, args):
        monthly = True  # Set to False for yearly subfolders!!!
        
        folder_path = Path(args.path.rstrip(os.sep))
        collection = collections.defaultdict(list)
        for file_path in Path(folder_path).glob('**/*'):
            lower_name = str(file_path).lower()
            dt = None
            if lower_name.endswith("jpg") or lower_name.endswith("jpeg") or lower_name.endswith("png"):
                dt = self.getImageCaptureTimestamp(file_path)
            elif lower_name.endswith("mov") or lower_name.endswith("mp4"):
                dt = self.getVideoCaptureTimestamp(file_path)
            
            key = 'Unknown'
            if dt is not None and len(dt)>7:
                key = str(dt[0:4]) + "_" + str(dt[5:7]) if monthly else str(dt[0:4])
            collection[key].append(file_path)
            
        for key,files in collection.items():
            subfolder_path = os.path.join(folder_path, key)
            print(f"{subfolder_path}: {len(files)} items")
            if not os.path.exists(subfolder_path):
                print(f"Create new subfolder {subfolder_path}")
                if not args.dry_run: 
                    os.makedirs(subfolder_path)
            for f in files:
                if args.verbose: 
                    print(f"Moving file {f} to subfolder {subfolder_path}")
                if not args.dry_run: 
                    shutil.move(f, subfolder_path)       
    

    ###############################################################################################################
    # Mode 6 - move from direct subfolders (non-recursively) to this folder
    #          name clashes are avoided by renaming a file before moving if necessary
    def safeMoveFromSubfolders(self, args):
        # TODO
        if args.dry_run:
            print("DRY_RUN is not available for this mode. Quit!")
            return        
        
        # non recursive, also ignores files directly in folder_path
        folder_path = Path(args.path.rstrip(os.sep))
        files = [path for path in Path(folder_path).glob("*/*") if path.is_file()]
        for f in files:   
            target_path = os.path.join(folder_path, os.path.basename(f))
            #print(f"{f}   -->  {target_path}")

            if os.path.isfile(target_path):
                print(f"File {target_path} already exists ...")
                num = 0
                while True:                    
                    num = num + 1
                    tp_pre, tp_ext = os.path.splitext(target_path)
                    target_path_renamed = tp_pre + "__" + f"{num:03}" + tp_ext
                    if not os.path.isfile(target_path_renamed):
                        target_path = target_path_renamed
                        print(f"...renaming to {os.path.basename(target_path)} when moving from subfolder {os.path.basename(os.path.dirname(f))}")
                        break
            shutil.move(f, target_path)      


    ###############################################################################################################
    # Mode 7 - move from direct subfolders (non-recursively) to this folder
    #          name clashes are avoided by renaming a file before moving if necessary
    def removeRenamingSuffixes(self, args):
        # TODO: Dry mode can give different results
        #       e.g. if files A/img_001.JPG and B/img_002.jpg exist, we only
        #       for an existing file img.JPG in current folder when in dry_mode -> NO PROBLEM
        #       In reality, the first renaming will create such a file, so the second one won't work! 

        folder_path = Path(args.path.rstrip(os.sep))
        files = [path for path in folder_path.glob("**/*") if path.is_file() and re.match(r".*__[0-9][0-9][0-9]\.[a-z]+$", os.path.basename(path), re.IGNORECASE)]
        renameOk = 0
        for f in files:
            path_pre, path_ext = os.path.splitext(f)
            renamed_f = path_pre[0:path_pre.rfind("__")] + path_ext
            if not os.path.isfile(renamed_f):
                renameOk += 1
                if not args.dry_run:
                    shutil.move(f, renamed_f) 
            else:
                print(f"Cannot rename {f} to {renamed_f} because this filename exists already\n")
        canWill = "can be" if args.dry_run else "were"
        print(f"Detected {len(files)} files with suffix __XXX -> {renameOk} " + canWill + " renamed safely")


    ###############################################################################################################
    # Helper functions start here
    def createSubfolderAndMove(self, is_dry_run, is_verbose, parent_folder_path, subdir_name, match_criterion):
        prefix = "DRY_RUN: " if is_dry_run else ""

        # determine matching files
        matching_files = []
        print("\n{}".format(subdir_name))
        for file in os.listdir(parent_folder_path):
            if os.path.isfile(os.path.join(parent_folder_path, file)):
                if match_criterion(file, parent_folder_path):
                    matching_files.append(os.path.join(parent_folder_path, file))
        if not matching_files:
            print(prefix + "No matching files. Directory {} will not be created.".format(subdir_name))
            return

        # create subfolder if matching files
        subdir_path = os.path.join(parent_folder_path, subdir_name)
        if not os.path.exists(subdir_path):
            print(prefix + "Create new subfolder {}".format(subdir_path))
            if not is_dry_run: os.makedirs(subdir_path)
        else:
            print("WARNING: The subfolder {} already exists!".format(subdir_path))

        # move matching files to subfolder
        for f in matching_files:
            if is_verbose: print("  " + prefix + "Moving file {}".format(f))
            if not is_dry_run: shutil.move(f, subdir_path)
        print(prefix + "Moved {} files into subfolder".format(len(matching_files)))


    def moveToParentAndDeleteSubfolder(self, is_dry_run, is_verbose, parent_dir, subdir_name):
        prefix = "DRY_RUN: " if is_dry_run else ""

        subdir = os.path.join(parent_dir, subdir_name)
        if not os.path.isdir(subdir):
            print("\n" + prefix +  "{} does not exist!".format(subdir_name))
            return

        print("\n{}".format(subdir_name))
        num_moved=0
        for file in os.listdir(subdir):
            subdir_file_path = os.path.join(subdir, file)
            if os.path.isfile(subdir_file_path):
                if is_verbose: print("  " + prefix + "Moving file {} back to parent folder".format(file))
                if not is_dry_run: shutil.move(subdir_file_path, parent_dir)
                num_moved = num_moved + 1
        print(prefix + "Moved {} files back from subfolder".format(num_moved))

        if len(os.listdir(subdir)) == 0:
            print(prefix + "Removing empty subdir {}".format(subdir))
            if not is_dry_run: shutil.rmtree(subdir)
        else:
            if not is_dry_run: print("WARNING: Cannot remove subfolder {} because it is not empty.".format(subdir))

    
    # TODO: Also use exiftool instead exifread for this
    def getImageCaptureTimestamp(self, file_path):
        try:
            f = open(file_path, 'rb')
            tags = exifread.process_file(f)
            if not tags:
                Exception("no exif found!")
            dt = str(tags['EXIF DateTimeOriginal'])
            return dt
        except:
            return None
        finally:
            f.close()


    # capture date of MOV from its stored meta-data (platform specific)
    def getVideoCaptureTimestamp(self, file_path):
        current_platform = platform.system()
        if current_platform == 'Windows':
            return self.getVideoCaptureTimestampWindows(file_path)
        elif current_platform == 'Linux':
            return self.getVideoCaptureTimestampLinux(file_path)
        else:
            print(f"Unknown platform: {current_platform}")

    def getVideoCaptureTimestampWindows(self, file_path):
        try:
            try: 
                import win32com
                print("win32com was already installed")
            except ImportError:
                from pip._internal import main as pip
                pip(['install', '--user', 'win32com'])
                import win32com
                print("Installed and imported win32com")

            properties = win32com.propsys.SHGetPropertyStoreFromParsingName(file_path)
            dt = properties.GetValue(win32com.pscon.PKEY_Media_DateEncoded).GetValue()

            if not isinstance(dt, datetime.datetime):
                # In Python 2, PyWin32 returns a custom time type instead of
                # using a datetime subclass. It has a Format method for strftime
                # style formatting, but let's just convert it to datetime:
                dt = datetime.datetime.fromtimestamp(int(dt))
                dt = dt.replace(tzinfo=pytz.timezone('UTC'))

            # dt_vienna = dt.astimezone(pytz.timezone('Europe/Vienna'))   # if time zone is needed
            return str(dt)
        except:
            return None

    def getVideoCaptureTimestampLinux(self, file_path):
        try:
            with exiftool.ExifToolHelper() as et:
                metadata = et.get_metadata(file_path)
                dt = metadata[0]["QuickTime:CreationDate"]
                return dt
        except:
            return None            

    def checkIfDirWithNonReservedName(self, folder_path):

        # path must be a folder
        if not os.path.isdir(folder_path):
            print("The given path {} is not a folder. Exit.".format(folder_path))
            return False

        # path must not have a reserved name
        folder_name = Path(os.path.basename(folder_path))  # e.g.: MyFolder
        if folder_name in self.special_folders:
            print("The folder {} has a reserved name and will not be modified. "
                  "\nReserved names are: {}. "
                  "\nExit.".format(folder_name, self.special_folders))
            return False

        return True



###############################################################################################################
# Call main() as starting point
if __name__ == '__main__':
    PictureMagic().main()