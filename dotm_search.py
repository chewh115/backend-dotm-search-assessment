#!/usr/bin/env python
# -*- coding: utf-8 -*-

"""
Given a directory path, search all files in the path for a given text string
within the 'word/document.xml' section of a MSWord .dotm file.
"""
__author__ = "chewh115"

import argparse
import os


parser = argparse.ArgumentParser(
    description="Searches files in a directory for given keyword")

parser.add_argument('search-term', action="store",
                    help="Keyword you want to search for.")
parser.add_argument('--directory', 'path', action="store", default='./',
                    help="The directory to search in. Default is current directory.")


def get_files(path):
    for root, dirs, files in (path, topdown=False):
        pass


def main():
    raise NotImplementedError("Your awesome code begins here!")


if __name__ == '__main__':
    main()
