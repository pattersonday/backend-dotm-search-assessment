#!/usr/bin/env python
# -*- coding: utf-8 -*-

"""
Given a directory path, search all files in the path for a given text string
within the 'word/document.xml' section of a MSWord .dotm file.
"""
import zipfile
import os
import sys
import argparse

__author__ = "Patterson Day"

file_name = './dotm_files'

def create_parser():

    """Creating parser for our dotm search"""

    parser = argparse.ArgumentParser(
        description='searches through our text for specific argument')
    parser.add_argument('--dir', help='directory that we are searching in')
    parser.add_argument('text', help='text that we are searching for')

    return parser

def main():
    
    parser = create_parser()
    name_space = parser.parse_args()
    search_text = name_space.text
    search_path = name_space.dir

    print('Searching directory {} for dotm files with text {}...'.format(
        search_path, search_text))

    file_list = os.listdir(search_path)
    match_count = 0
    amount_of_files = 0

    for file in file_list:
        if not file.endswith('.dotm'):
            print('Disregard filename: ' + file)
            continue
        else:
            amount_of_files += 1
        full_path_name = os.path.join(search_path, file)
        if zipfile.is_zipfile(full_path_name):
            with zipfile.ZipFile(full_path_name) as z:
                file_list_inside_zip_archives = z.namelist()
                if 'word/document.xml' in file_list_inside_zip_archives:
                    with z.open('word/document.xml', 'r') as doc:
                        for line in doc:
                            line = line.decode('utf-8')
                            text_location = line.find(search_text)
                            if text_location >= 0:
                                print('Match found in file {}'.format(
                                    full_path_name))
                                print('...' + line[
                                    text_location - 40: text_location + 41:])
                                match_count += 1
    print('Total .dotm files searched: {}'.format(amount_of_files))
    print('Total .dotm files matched: {}'.format(match_count))

if __name__ == '__main__':
    main()
