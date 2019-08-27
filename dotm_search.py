#!/usr/bin/env python
# -*- coding: utf-8 -*-

"""
Given a directory path, search all files in the path for a given text string
within the 'word/document.xml' section of a MSWord .dotm file.
"""

import argparse
import os
import zipfile

__author__ = "Jacob Walker"

parser = argparse.ArgumentParser()
parser.add_argument("symbol")
parser.add_argument("--dir", default="./")
args = parser.parse_args()


def get_dotm_files(dir_path):
    return [os.path.join(dir_path, f)
            for f in os.listdir(dir_path)
            if os.path.isfile(os.path.join(dir_path, f))
            and f.endswith('.dotm')
            and zipfile.is_zipfile(os.path.join(dir_path, f))
            ]


def get_symbol(symbol, file):
    dotm = zipfile.ZipFile(file, 'r')
    with dotm.open('word/document.xml', 'r') as f:
        for line in f:
            string = str(line)
            if symbol in string:
                index = string.find(symbol)
                return "..." + string[index - 40:index + 40] + "..."
    return


def get_matches(files):
    number_of_matches = 0
    for file in files:
        found = get_symbol(args.symbol, file)
        if found:
            print("Found '{}' \n In {}".format(found, file))
            number_of_matches += 1
    print("Found {} in {}/{} .dotm files".format(args.symbol, number_of_matches, len(files)))


def main():
    if not os.path.isdir(args.dir):
        print("Directory {} did not exist".format(args.dir))
        return
    print("Searching the directory for .dotm files containing any occurrence of {}".format(args.symbol))
    files = get_dotm_files(args.dir)
    get_matches(files)


if __name__ == '__main__':
    main()
