# Note: largely based on https://stackoverflow.com/a/48403427

from argparse import ArgumentParser
from glob import glob
import logging
import os
import os.path
import re
import sys

import pythoncom
from win32com.shell import shell, shellcon

logging.basicConfig(level=logging.INFO)
logger = logging.getLogger()

def update_shortcut(filename, target_from, target_to):
    link = pythoncom.CoCreateInstance(
        shell.CLSID_ShellLink,    
        None,
        pythoncom.CLSCTX_INPROC_SERVER,    
        shell.IID_IShellLink
    )
    persist_file = link.QueryInterface(pythoncom.IID_IPersistFile)
    persist_file.Load(filename)

    target, _ = link.GetPath(shell.SLGP_UNCPRIORITY)

    if re.search('(?i)' + re.escape(target_from), target) is not None:
        logger.info(f"Found link target in {filename}: {target}")
        if target_to is not None:
            new_target = re.sub('(?i)' + re.escape(target_from), target_to, target)
            new_dir = os.path.dirname(new_target)
            
            logger.debug(f"Changing link target in {filename} from {target} --> {new_target}...")
            link.SetPath(new_target)
            logger.info(f"Changed link target in {filename} from {target} --> {new_target}")
            
            logger.debug(f"Changing start dir in {filename} to {new_dir}...")
            link.SetWorkingDirectory(new_dir)
            logger.info(f"Changed start dir in {filename} to {new_dir}")

            logger.debug(f"Saving {filename}...")
            persist_file.Save(os.path.realpath(filename), 0)  
  

def search_links(dir, pattern):
  for filename in glob(os.path.join(dir,pattern), recursive=True):
      if filename.endswith(".lnk"):
          yield filename


def main(raw):
    cmd = ArgumentParser(description="Search and replace shortcut targets and working directories")
    cmd.add_argument("--root", default=os.getcwd(), help="root search directory")
    cmd.add_argument("--pattern", default="**\\*.lnk", help="search pattern (glob)")
    cmd.add_argument("-t", "--target-replace", default=None, help="link target replacement text")
    cmd.add_argument("--debug", action="store_true", help="debug logging on")
    cmd.add_argument("--no-debug", action="store_false", help="debug logging off")
    cmd.add_argument("target", help="link target search text") 
    cmd.set_defaults(debug=False)
    args = cmd.parse_args(raw)
    
    if args.debug:
        logger.setLevel(logging.DEBUG)
    
    for f in search_links(args.root, args.pattern):
        logger.debug(f"Found link {f}")
        update_shortcut(f, args.target, args.target_replace)

if __name__ == "__main__":
    main(sys.argv[1:])
    
 