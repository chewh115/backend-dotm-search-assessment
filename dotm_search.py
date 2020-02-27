#!/usr/bin/env python
# -*- coding: utf-8 -*-

"""
Given a directory path, search all files in the path for a given text string
within the 'word/document.xml' section of a MSWord .dotm file.
"""
__author__ = "chewh115, and I talked it out with Janell.Huyck while coding."

import argparse
import os
import zipfile


def create_parser():
    parser = argparse.ArgumentParser(
        description="Searches files in a directory for given keyword")

    parser.add_argument('search_term', action="store",
                        help="Keyword you want to search for.")
    parser.add_argument('--directory', action="store", default='./',
                        help="The directory to search in.")
    search_info = parser.parse_args()
    return search_info


def get_dotm_files(directory):
    file_list = []
    for (root, dirs, files) in os.walk(directory):
        for file in files:
            if file.endswith('.dotm'):
                file_list.append(os.path.join(root, file))
    return file_list


def unzip_file(directory, search_term):
    file_list = get_dotm_files(directory)
    search_term_counter = 0
    for file in file_list:
        with zipfile.ZipFile(file, 'r') as zipped_files:
            zipped_files.extractall('./extracted')
        file_path = './extracted/word/document.xml'
        search_context = search_for_term_and_context(file_path, search_term)
        if search_context != -1:
            print('Match found in: {}'.format(file))
            print('The context is: {}'.format(search_context))
            search_term_counter += 1
    print('Number of files searched for {}: '.format(
        search_term) + str(len(file_list)))
    print('We found {} files containing {}.'.format(
        search_term_counter, search_term))


def search_for_term_and_context(file_path, search_term):
    with open(file_path, 'r') as file_to_search:
        file_words = file_to_search.read()
        term_index = file_words.find(search_term)
        context_slice = min(len(file_words), (term_index + 40))
        if term_index == -1:
            return -1
        elif term_index < 40:
            return file_words[0:context_slice]
        else:
            return file_words[term_index - 40:context_slice]


def main():
    search_info = create_parser()
    directory = search_info.directory
    search_term = search_info.search_term
    unzip_file(directory, search_term)


if __name__ == '__main__':
    main()
