#!/usr/bin/env python
# -*- coding: utf-8 -*-

"""
Given a directory path, search all files in the path for a given text string
within the 'word/document.xml' section of a MSWord .dotm file.
"""
__author__ = "chewh115"

import argparse
import os
import zipfile


def create_parser():
    parser = argparse.ArgumentParser(
        description="Searches files in a directory for given keyword")

    parser.add_argument('search_term', action="store",
                        help="Keyword you want to search for.")
    parser.add_argument('--directory', action="store", default='./',
                        help="The directory to search in. Default is current directory.")
    search_info = parser.parse_args()
    return search_info


def get_dotm_files(path):
    file_list = []
    for (root, dirs, files) in os.walk(path):
        for file in files:
            if file.endswith('.dotm'):
                file_list.append(os.path.join(root, file)
    return file_list


def unzip_file(files):
    unzipped_files=[]
    for file in files:
        with zipfile.Zipfile(file, 'r') as zip:
            unzipped_files.append(zip.extract('word/document.xml'))
    return unzipped_files


def main():
    search_info=create_parser()
    directory=search_info.directory
    search_term=search_info.search_term
    file_list=get_files(directory)
    unzipped_files=unzip_files(file_list)


if __name__ == '__main__':
    main()
