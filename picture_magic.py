import argparse
from genericpath import isfile
import os
import shutil
import re
import collections

from glob import glob
from pathlib import Path

from exif import Image
import exifread
import exiftool
import fnmatch
import platform


import pytz
import datetime
#import win32com


class PictureMagic(object):

    # reserved folder names
    dirname_aae       = "_EditDataAAE"
    dirname_downloads = "_WhatsApp,Downloads,etc"
    dirname_live_imgs = "_LiveImages"
    dirname_originals = "_Originals"
    dirname_remaining = "Pics_and_Movies"
    special_folders = [dirname_aae, dirname_downloads, dirname_live_imgs, dirname_originals, dirname_remaining]

    # TODO: We also need a mode for moving to subfolders per year (and back).
    #       Maybe the approach using lambdas can be reused in a good way for this!

    def main(self):
        mode_map = {0: 'show folder statistics',
                    1: 'verify year',
                    3: '* move to subfolders',
                    4: '* move back from subfolders'}

        dict_to_str = lambda x: ','.join(" %s:%s" % (str(k), str(v)) for (k, v) in x.items()) if isinstance(x,dict) else x
        parser = argparse.ArgumentParser(description='Organize my pictures')
        parser.add_argument('--mode', dest='mode', required=True, type=int, choices=mode_map.keys(),
                           help='mode of operation. {}'.format(dict_to_str(mode_map)))
        parser.add_argument('--path', dest='path', required=True, type=str,
                           help='path to file or folder')
        parser.add_argument('--year', dest='year', required=False, type=int,
                           help='expected year of capturing for mode 1 (must be 1900 <= x <= 2100)')
        parser.add_argument('--dry_run', dest='dry_run', required=False, action='store_true',
                           help='If true, also the pertaining modes (marked with *) will not make any changes and '
                                'instead only output the planned operations as text.')
        parser.add_argument('--verbose', dest='verbose', required=False, action='store_true',
                           help='If true, print verbose log information.')

        args = parser.parse_args()
        print("Running program in mode: {}".format(args.mode))
        #print("Path is: {}".format(args.path))

        if not os.path.exists(args.path):
            print("The given path {} does not exist. Exit.".format(args.path))
            return

        if args.mode==0:
            self.showStats(args)
        elif args.mode == 1:
            self.verifyYear(args)
        elif args.mode == 3:
            self.moveToSubfolders(args)
        elif args.mode == 4:
            self.moveBackFromSubfolders(args)

    #####################################
    # Mode 0 - show file statistics
    ####################################
    def showStats(self, args):
        if not os.path.exists(args.path):
            print("The given path {} DOESNT exist. Exit.".format(args.path))

        # Shows the number of files for each file type recursively
        files = [path.suffix for path in Path(args.path).glob("**/*") if path.is_file() and path.suffix]
        data = collections.Counter(files)
        for key, val in data.items():
            print("AAA")
            print(f'{key}: {val}')


    #####################################
    # Mode 1 - verify year
    ####################################
    def verifyYear(self, args):
        # Checks all JPG and MOV files in a folder, whether the media timestamps have the expected year.
        # In this mode the expected year is provided in args.year.
        # Prints warning for each file with wrong or unknown year.
        # Subfolders and file types other than JPG or MOV are ignored.

        if not os.path.exists(args.path):
            print(f"The given path {args.path} does not exist. Exit.")
            return
        if args.year < 1900 or args.year > 2100:
            print(f"The given year {args.year} must be between 1900 and 2100. Exit.")
            return

        # this is the function to check the correct year
        def check_dt(dt, expected_year, is_verbose, stats):
            if dt is None:
                if is_verbose: 
                    print(f"UNKNOWN Capture year for {file_path}.")
                stats['UNKNOWN'] = stats['UNKNOWN'] + 1
                return

            parsed_year = str(dt[0:4])
            if str(parsed_year) == str(expected_year):
                if is_verbose: 
                    print(f"CORRECT Capture year ({parsed_year}) of {file_path} -> exact date: {dt}")
                stats['CORRECT'] = stats['CORRECT'] + 1
            else:
                if is_verbose:
                    print(f"WRONG Capture year ({parsed_year}) of {file_path} -> exact date: {dt}")
                stats['WRONG'] = stats['WRONG'] + 1

        statistics = {'UNKNOWN': 0, 'CORRECT': 0, 'WRONG': 0}

        if os.path.isdir(args.path):
            folder_path = Path(args.path.rstrip(os.sep))   # e.g.: C:\Users\Gert\MyFolder\2019
            # check images recursively
            for file_path in Path(folder_path).glob('**/*.JPG'):
                dt = self.getImageCaptureTimestamp(file_path)
                check_dt(dt, args.year, args.verbose, statistics)
            # check movies recursively
            for file_path in Path(folder_path).glob('**/*.MOV'):
                dt = self.getVideoCaptureTimestamp(file_path)
                check_dt(dt, args.year, args.verbose, statistics)

        elif os.path.isfile(args.path):
            file_path = args.path
            dt = self.getVideoCaptureTimestamp(file_path)
            check_dt(dt, args.year, args.verbose, statistics)


        print(f"Processed files: {statistics['UNKNOWN'] + statistics['WRONG'] + statistics['CORRECT']}")
        print(f"Num. files with date UNKNOWN: {statistics['UNKNOWN']}")
        print(f"Num. files with date CORRECT: {statistics['CORRECT']}")
        print(f"Num. files with date WRONG: {statistics['WRONG']} --> this should be 0!")


    #####################################
    # Mode 3 - move to subfolders (changes files)
    ####################################
    def moveToSubfolders(self, args):
        folder_path = Path(args.path.rstrip(os.sep))       # e.g.: C:\Users\Gert\MyFolder
        if not self.checkIfValidDir(folder_path):
            return

        # TODO: Since we don't really move in dry_run, the number of matching files is not equal to normal mode
        #       if a file matches multiple criteriums

        # TODO: For performance reasons it would be way better if we only iterate once over all files.
        #       If we classify each file in one run (criterium after criterium) and store what is going to
        #       happen with each file we could avoid this problem AND the one above.

        # Folder: _EditDataAAE
        # All files with type AAE (*.AAE)
        criterion_aae = lambda curr_file, curr_folder: True if curr_file.endswith('AAE') else False
        self.createSubfolderAndMove(args.dry_run, args.verbose, folder_path, self.dirname_aae, criterion_aae)

        # # Folder: _WhatsApp, Downloads, etc
        # # All files that do not start with IMG_*  OR
        # # that have a type other than JPG or MOV  OR
        # # TODO: JPG files named IMG* that have an empty "Date taken" (WhatsApp clears all EXIF data)
        criterion_downloads = lambda curr_file, curr_folder: True if not re.match(r"^IMG_.*\.(JPG|MOV)$", curr_file) else False
        self.createSubfolderAndMove(args.dry_run, args.verbose, folder_path, self.dirname_downloads, criterion_downloads)


        # # Folder: _LiveImages
        # # All MOV files for which a corresponding JPG file (with the same name) exists.
        # # --> IMG_0433.MOV, IMG_E0910.MOV
        def criterion_live_imgs(curr_file, curr_folder):
            if curr_file.endswith('.MOV'):
                potential_twin_file = '.JPG'.join(curr_file.rsplit('.MOV', 1))    # replace only last occurrence of .MOV with .JPG
                if os.path.isfile(os.path.join(curr_folder, potential_twin_file)):
                    return True
            return False
        self.createSubfolderAndMove(args.dry_run, args.verbose, folder_path, self.dirname_live_imgs, criterion_live_imgs)

        # TODO: This is not so efficient anymore since we check all IMG_ files for matching IMG_E instead the other way round.
        #       Also this approach does not allow for warnings if an IMG_E has no matching Original.
        # Folder: _Originals
        # Contains all files IMG_<num>.JPG (or MOV) for which an edited version IMG_E<num>.JPG (or MOV) exists.
        # --> IMG_0072.JPG (IMG_E0072.JPG exists)
        # --> IMG_E1369.MOV (IMG_E1369.MOV exists)
        # --> WARNING FOR IMG_E0815.JPG --> no corresponding original   # TODO: Not possible this way
        def criterion_originals(curr_file, curr_folder):
            if curr_file.startswith('IMG_'):
                potential_twin_file = 'IMG_E'.join(curr_file.split('IMG_', 1))  # replace only first occurrence of IMG_ with IMG_E
                if os.path.isfile(os.path.join(curr_folder, potential_twin_file)):
                    return True
            return False
        self.createSubfolderAndMove(args.dry_run, args.verbose, folder_path, self.dirname_originals, criterion_originals)

        # # Folder: Pics_and_Movies
        # # All remaining files. Should be only JPG and MOV, where MOV is NOT a live video.
        criterion_remaining = lambda curr_file, curr_folder: True if re.match(r"^IMG_.*\.(JPG|MOV)$", curr_file) else False
        self.createSubfolderAndMove(args.dry_run, args.verbose, folder_path, self.dirname_remaining, criterion_remaining)

        print("Done!")


    #####################################
    # Mode 4 - move back from subfolders to revert mode 3 operations (changes files)
    ####################################
    def moveBackFromSubfolders(self, args):
        folder_path = Path(args.path.rstrip(os.sep))       # e.g.: C:\Users\Gert\MyFolder
        if not self.checkIfValidDir(folder_path):
            return

        self.moveToParentAndDeleteSubfolder(args.dry_run, args.verbose, folder_path, self.dirname_aae)
        self.moveToParentAndDeleteSubfolder(args.dry_run, args.verbose, folder_path, self.dirname_downloads)
        self.moveToParentAndDeleteSubfolder(args.dry_run, args.verbose, folder_path, self.dirname_live_imgs)
        self.moveToParentAndDeleteSubfolder(args.dry_run, args.verbose, folder_path, self.dirname_originals)
        self.moveToParentAndDeleteSubfolder(args.dry_run, args.verbose, folder_path, self.dirname_remaining)
        print("Done!")


    #####################################
    # Helper function for mode 3
    ####################################
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


    #####################################
    # Helper function for mode 4
    ####################################
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


    #####################################
    # Helper function to determine capture date of JPG image from its EXIF data
    # TODO: Also use exiftool instead exifread for this
    ####################################
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


    #####################################
    # Helper function to determine capture date of MOV video from its stored meta-data (platform specific)
    ####################################
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


    #####################################
    # Helper function
    ####################################
    def checkIfValidDir(self, folder_path):

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

#####################################
# Call main() as starting point
####################################
if __name__ == '__main__':
    PictureMagic().main()