#!/usr/bin/env python3

import argparse
from collections import OrderedDict
import csv
from datetime import datetime, timezone, timedelta, date
from zoneinfo import ZoneInfo
import io
import json
import os
import re
import shutil
import sys
from zipfile import ZipFile

from PyPDF2 import PdfReader, PdfWriter
from tqdm import tqdm

from score import (
    Binder,
    ShelfMetadata,
)
from sharepoint import SharePointClient, SharePointFile, SharePointFileType
from util import (
    get_user_input,
    is_file_name_sanitized,
    to_utc,
    yellow_text,
    yes_no_validator,
    make_cover_page
)

class Librarian():
    def __init__(self, shelf):
        # type: (str|None) -> None
        self.shelf = Librarian._get_shelf_or_default(shelf)
        self.metadata = self.read_shelf_metadata()
        self.binders = self.read_binder_data()
        print('Initializing SharePoint connection...')
        self.client = SharePointClient()
        print('Connected!')

    @staticmethod
    def init_shelf(shelf, sharepoint_path):
        # type: (str|None, str|None) -> None
        print('Initializing SharePoint connection...')
        client = SharePointClient()
        print('Connected!')

        # process user input
        if shelf is None:
            shelf = get_user_input('Enter shelf name: ', is_file_name_sanitized)
        elif not is_file_name_sanitized(shelf):
            raise ValueError('Please remove special characters in shelf name')
        if sharepoint_path is None:
            sharepoint_path = get_user_input('Enter SharePoint path: ')

        # query
        score_titles = [f.name for f in client.listdir(sharepoint_path) if
            f.file_type == SharePointFileType.FOLDER]
        filtered_score_titles = []
        for title in score_titles:
            response = get_user_input(f'Include "{title}" (y/yes/n/no)? ', yes_no_validator)
            if response.lower() in ['y', 'yes']:
                filtered_score_titles.append(title)

        # create directory structure
        if os.path.exists(f'library/{shelf}'):
            raise ValueError('Shelf already exists')
        os.mkdir(f'library/{shelf}')
        os.mkdir(f'library/{shelf}/binders')
        os.mkdir(f'library/{shelf}/diffs')
        os.mkdir(f'library/{shelf}/scores')
        for title in filtered_score_titles:
            os.mkdir(f'library/{shelf}/scores/{title}')
        # binders.csv
        with open(f'library/{shelf}/binders.csv', 'w') as f:
            writer = csv.writer(f)
            writer.writerow(['Binder','Name(s)','Virtual?'] + filtered_score_titles)
        # metadata.json
        metadata: ShelfMetadata = {
            'sharepoint_path': sharepoint_path,
            'scores': {},
        }
        for title in filtered_score_titles:
            metadata['scores'][title] = OrderedDict()
        with open(f'library/{shelf}/metadata.json', 'w') as f:
            json.dump(metadata, f, indent=4)

        print('\nCreated entries for the following:')
        for title in filtered_score_titles:
            print(f'- {title}')
        print(f'\nGo to library/{shelf}/metadata.json and library/{shelf}/binders.csv to make any changes.\n')

    @staticmethod
    def _get_shelf_or_default(shelf):
        # type: (str|None) -> str
        folders = [p for p in os.listdir('library') if os.path.isdir(f'library/{p}')]
        if shelf is not None:
            if shelf in folders:
                return shelf
            raise ValueError(f'Shelf not found. Please specify one of {folders}')
        if len(folders) == 1:
            return folders[0]

        if len(folders) == 0:
            raise ValueError(f'No shelves found. ' +
                'Use librarian.py init to create one.')
        else:
            raise ValueError(f'Multiple shelves found. ' +
                'Please specify one of {folders}')

    def sync_shelf(self, refresh=False):
        # type: (bool) -> None
        # for each entry in parts, query sharepoint and check for differences
        score_path = f'library/{self.shelf}/scores'
        scores = os.listdir(score_path)
        scores = [i for i in scores if i in self.metadata["scores"]] #Can you make this more compact?
        for score in scores:
            print(f'Syncing {score}...')
            #print([i for i in self.metadata["scores"]])
            score_metadata = self.metadata['scores'][score]
            share_path = os.path.join(self.metadata['sharepoint_path'], score)
            versions = [v for v in self.client.listdir(share_path) if
                v.file_type == SharePointFileType.FOLDER]
            versions.sort(key=lambda f: f.created_at)
            tracked_versions = os.listdir(f'{score_path}/{score}')
            tracked_versions = [p for p in tracked_versions if
                os.path.isdir(f'{score_path}/{score}/{p}')]
            for v in versions:
                if v.name not in tracked_versions:
                    print(f'Downloading {v.name}...')
                    self._add_score_version(score, v)
                elif refresh:
                    version_metadata = score_metadata[v.name]
                    download_date = datetime.fromisoformat(
                        version_metadata['last_downloaded'])
                    last_edit_date = self.client.get_last_activity_date(v)
                    if last_edit_date > download_date:
                        print(f'Redownloading {v.name}...')
                        os.rmdir(f'library/{self.shelf}/scores/{score}/{v.name}')
                        self._add_score_version(score, v)
            print('Done!\n')
        print(f'All scores synced. Go to library/{self.shelf}/metadata.json for future changes.')

    def _add_score_version(self, score, version):
        # type: (str, SharePointFile) -> None
        blob = self.client.download_zip(version)
        with ZipFile(io.BytesIO(blob), 'r') as zf:
            # TODO: check it's well formatted
            zf.extractall(f'library/{self.shelf}/scores/{score}')
        score_metadata = self.metadata['scores'][score]
        if version.name in score_metadata:
            print('Metadata already exists!')
            score_metadata[version.name]['last_downloaded'] = datetime.now().isoformat()
        else:
            regex = self._prompt_regex(score, version.name)
            self.metadata['scores'][score][version.name] = {
                'regex_str': regex,
                'created_at': version.created_at.isoformat(),
                'last_downloaded': datetime.now().isoformat(),
            }
        self.write_shelf_metadata()

    def _prompt_regex(self, score, version):
        # type: (str, str) -> str
        print('Enter a regex to match the version\'s files (examples below):')
        entries = os.listdir(f'library/{self.shelf}/scores/{score}/{version}')
        entries = [e for e in entries if e.endswith('.pdf')][:5]
        for entry in entries:
            print(f'- {entry}')
        prev_metadata = None
        if self.metadata['scores'][score]:
            prev_version = next(reversed(self.metadata['scores'][score]))
            prev_metadata = self.metadata['scores'][score][prev_version]
            print('Previous regex:', prev_metadata['regex_str'])
        while True:
            regex_str = get_user_input('Regex (leave blank to copy previous): ')
            if prev_metadata is not None and regex_str == '':
                regex_str = prev_metadata['regex_str']
            regex = re.compile(regex_str)
            print('Example matches:')
            for entry in entries:
                slug = re.fullmatch(regex, entry)
                if slug is not None:
                    print(f'- "{slug.group(1)}" ({entry})')
                else:
                    print(f'- NO MATCH ({entry})')
            confirm = get_user_input('Looks ok (y/yes/n/no)? ', yes_no_validator)
            if confirm.lower() in ['y', 'yes']:
                break
        return regex_str

    def make_all_binders(self):
        # TODO: add a warning text array so we can print them at the end of a tqdm
        warnings: list[str] = []
        pbar = tqdm(enumerate(self.binders))
        for i, binder in pbar:
            if binder.is_virtual and self.has_update(binder):
                pbar.set_description(f'Generating virtual binder for {binder.title}')
                self._make_binder(i, binder, warnings)
        for w in warnings:
            self._print_warning(w)

    def has_update(self, binder):
        updates = self.get_updated_parts(binder)
        week_start = (datetime.now() - timedelta(days=7)).replace(tzinfo = ZoneInfo("UTC"))
        for s_title, s_meta in self.metadata['scores'].items():
            if s_title in updates:
                part_paths: dict[str, str] = {}
                # start from oldest version, so newer update replace the older ones
                for v_name, v_meta in s_meta.items():
                    created_at = datetime.fromisoformat(v_meta['created_at'])
                    if created_at < week_start:
                        continue
                    part_matches = self._get_binder_version_matches(binder, s_title, v_name)
                    if len(part_matches) > 0:
                        return True
        return False

    def _print_warning(self, text):
        # type: (str) -> None
        print(f'[{yellow_text('WARNING')}] {text}')

    def _make_binder(self, i, binder, warnings):
        # type: (int, Binder, list[str]) -> None
        #{i+1:2d} - {binder.name} - {binder.title} - {date.today()}
        binder_folder_name = f'{binder.title} - {binder.names} - {date.today()}'
        binder_path = f'library/{self.shelf}/binders/{binder_folder_name}'
        updates = self.get_updated_parts(binder)
        week_start = (datetime.now() - timedelta(days=7)).replace(tzinfo = ZoneInfo("UTC"))
        # clear out binder path
        if os.path.isdir(binder_path):
            shutil.rmtree(binder_path)
        os.mkdir(binder_path)
        # add each score's part
        for s_title, s_meta in self.metadata['scores'].items():
            if s_title in updates:
                part_paths: dict[str, str] = {}
                # start from oldest version, so newer update replace the older ones
                for v_name, v_meta in s_meta.items():
                    created_at = datetime.fromisoformat(v_meta['created_at'])
                    if created_at < week_start:
                        continue
                    part_matches = self._get_binder_version_matches(binder, s_title, v_name)
                    for slug, path in part_matches.items():
                        part_paths[slug] = path
                # copy the parts over to the binders
                for slug in binder.parts[s_title]:
                    if slug in part_paths:
                        part_path = part_paths[slug]
                        part_name = os.path.basename(part_path)
                        shutil.copy2(part_path, binder_path)
                        print(f"Successfully copied {part_path} to {binder_path}")
                        #os.symlink(part_path, f'{binder_path}/{part_name}')
                    else:
                        warnings.append(f'{slug} part in {s_title} not found')

    # returns a map of slugs to the file they represent
    def _get_binder_version_matches(self, binder, score, version):
        # type: (Binder, str, str) -> dict[str, str]
        v_meta = self.metadata['scores'][score][version]
        regex = re.compile(v_meta['regex_str'])
        v_path = f'library/{self.shelf}/scores/{score}/{version}'
        v_path = os.path.abspath(v_path)

        part_matches: dict[str, str] = {}
        part_files = os.listdir(v_path)

        for part in part_files:
            slug = re.fullmatch(regex, part)
            if slug is not None:
                slug = slug.group(1)
            if slug in binder.parts[score]:
                part_matches[slug] = f'{v_path}/{part}'
        return part_matches

    def _parse_cutoff_datetime(self, after):
        # type: (datetime|None) -> datetime
        if after is None:
            # check diffs for file names
            diff_dates = os.listdir(f'library/{self.shelf}/diffs')
            if diff_dates:
                diff_dates.sort()
                return to_utc(datetime.fromisoformat(diff_dates[-1]))
            else:
                return datetime(2000, 1, 1, tzinfo=timezone.utc)
        return after

    def make_shelf_diffs(self, after=None, size=100):
        # type: (datetime|None, int) -> None
        after = self._parse_cutoff_datetime(after)
        new_binder_parts: list[list[str]] = []
        # TODO: print out unused paths
        # for each score, check if there are any versions past the indicated
        # date (default of last backup). if so, iterate through them and add
        # to collection. then do the actual diffing

        # print out warnings for unused paths here per version
        # batch parts
        pbar = tqdm(self.binders)
        player_indices = []
        index = 0
        for binder in pbar:
            if not binder.is_virtual:
                pbar.set_description(f'Diffing {binder.title}')
                new_binder_parts.append(self._make_binder_diff(binder, after))
                player_indices.append(index)
            index += 1
        # create version folder
        date_str = datetime.now().strftime('%Y-%m-%d')
        diff_path = f'library/{self.shelf}/diffs/{date_str}'
        if os.path.exists(diff_path):
            shutil.rmtree(diff_path)
        os.mkdir(diff_path)

        # get total number of pages to print
        page_count = 0
        for binder_parts in new_binder_parts:
            binder_page_count = 0
            for part_path in binder_parts:
                part_pdf = PdfReader(part_path)
                binder_page_count += (len(part_pdf.pages) + 1) // 2
            if binder_page_count > 0:
                page_count += binder_page_count + 1
        num_batches = page_count // (size * 1.1) + 1
        target_batch_size = page_count // num_batches
        # go through again and actually batch
        batch_num = 0
        page_count = 0
        batch_pdf = PdfWriter()
        pbar = tqdm(enumerate(new_binder_parts))
        for i, binder_parts in pbar:
            # ignore empty parts
            if not binder_parts:
                continue
            binder_pdf = PdfWriter()
            binder_pdf.append(make_cover_page(self.binders[player_indices[i]], datetime.now()))
            #Keep the cover pages omg they help so much with distribution
            binder_page_count = 1
            for part_path in binder_parts:
                part_pdf = PdfReader(part_path)
                binder_pdf.append(part_pdf)
                if len(part_pdf.pages) % 2 == 1:
                    binder_pdf.add_blank_page()
                binder_page_count += (len(part_pdf.pages) + 1) // 2
            # naive check to make all batches roughly equal sized
            page_delta = target_batch_size*(batch_num + 1) - page_count
            if (batch_num < num_batches - 1
                    and (page_delta < 0 or 2*page_delta < binder_page_count)):
                # flush batch
                batch_pdf.write(f'{diff_path}/batch{batch_num}.pdf')
                batch_pdf.close()
                batch_pdf = PdfWriter()
                batch_num += 1
            # add binder to batch
            buffer = io.BytesIO()
            binder_pdf.write(buffer)
            batch_pdf.append(buffer)
            page_count += binder_page_count
            pbar.set_description(f'Writing {binder_page_count:d} pages for {self.binders[i].title} ({page_count} total)')

        # save last batch
        batch_pdf.write(f'{diff_path}/batch{batch_num}.pdf')
        batch_pdf.close()

    def _make_binder_diff(self, binder, after):
        # type: (Binder, datetime) -> list[str]
        part_list: list[str] = []
        updates = self.get_updated_parts(binder) #Just update this every week depending on updates
        week_start = (datetime.now() - timedelta(days=7)).replace(tzinfo = ZoneInfo("UTC"))
        #There IS a way to automate this if we use the metadata information to determine when updates were
        #added then add updates based on that
        for s_title, s_meta in self.metadata['scores'].items():
            if s_title in updates:
                part_paths: dict[str, str] = {}
                # start from oldest version, so newer update replace the older ones
                for v_name, v_meta in s_meta.items():
                    created_at = datetime.fromisoformat(v_meta['created_at'])
                    if created_at < week_start:
                        continue
                    part_matches = self._get_binder_version_matches(binder, s_title, v_name)
                    for slug, path in part_matches.items():
                        part_paths[slug] = path
                        print(path)
                for path in part_paths.values():
                    part_list.append(path)
        return part_list

    def get_updated_parts(self, binder):
        week_start = (datetime.now() - timedelta(days=7)).replace(tzinfo = ZoneInfo("UTC"))
        updates = set()
        for s_title, s_meta in self.metadata["scores"].items():
            for v_name, v_meta in s_meta.items():
                created_at = datetime.fromisoformat(v_meta["created_at"])
                if created_at > week_start:
                    updates.add(s_title)
        return updates

    def check_all_versions(self, after):
        # type: (datetime|None) -> None
        after = self._parse_cutoff_datetime(after)
        for s_title, s_meta in self.metadata['scores'].items():
            for v_name, v_meta in s_meta.items():
                created_at = datetime.fromisoformat(v_meta['created_at'])
                if created_at < after:
                    continue
                self._check_version(s_title, v_name)

    def _check_version(self, score, version):
        # type: (str, str) -> None
        # get list of all valid slugs
        slugs = []
        for b in self.binders:
            for slug in b.parts[score]:
                slugs.append(slug)
        s_meta = self.metadata['scores'][score]
        v_meta = s_meta[version]
        paths = os.listdir(f'library/{self.shelf}/scores/{score}/{version}')
        paths = [p for p in paths if p.endswith('.pdf')]
        regex = re.compile(v_meta['regex_str'])
        for p in paths:
            slug = re.fullmatch(regex, p)
            if slug is None:
                self._print_warning(f'Unused path {p}')
            else:
                slug = slug.group(1)
                if slug not in slugs:
                    self._print_warning(f'Unused path {p}')

    def update_shelf(self, after, size):
        # type: (datetime|None, int) -> None
        after = self._parse_cutoff_datetime(after)
        self.sync_shelf()
        self.check_all_versions(after)
        self.make_all_binders()
        self.make_shelf_diffs(after, size)

    def read_shelf_metadata(self):
        # type: () -> ShelfMetadata
        with open(f'library/{self.shelf}/metadata.json', 'r') as f:
            return json.load(f, object_pairs_hook=OrderedDict)

    def write_shelf_metadata(self):
        # type: () -> None
        with open(f'library/{self.shelf}/metadata.json', 'w') as f:
            json.dump(self.metadata, f, indent=4)

    def read_binder_data(self):
        # type: () -> list[Binder]
        # TODO: check that all scores are represented?
        binders: list[Binder] = []
        with open(f'library/{self.shelf}/binders.csv', 'r') as f:
            reader = csv.reader(f)
            # consume first row
            score_names = next(reader)[3:]
            print(score_names)
            for row in reader:
                title = row[0]
                names = row[1]
                is_virtual = row[2] == 'Y'
                slugs = row[3:]

                parts = {}
                for score_name, slug in zip(score_names, slugs):
                    if slug == '':
                        parts[score_name] = []
                    else:
                        parts[score_name] = slug.split('/')
                binder = Binder(title, names, parts, is_virtual)
                binders.append(binder)
        return binders

    def compile_parts(self):
        pieces = self.metadata["scores"]
        for i in pieces:
            print(i)
            most_recent_parts = {}
            versions = list(pieces[i])

            all_parts = set()
            #Go to the first version and extract all the parts that should be there
            first_edition_path = f'library/{self.shelf}/scores/{i}/{versions[0]}'
            first_regex = re.compile(pieces[i][versions[0]]["regex_str"])
            first_edition_path = os.path.abspath(first_edition_path)
            first_edition_part_file = os.listdir(first_edition_path)
            for part in first_edition_part_file:
                slug = re.fullmatch(first_regex, part)
                if slug is not None:
                    slug = slug.group(1)
                    all_parts.add(slug)
            
            compiled_parts = set()

            #Go through the most recent versions and add parts until the compiled parts equals all the parts
            #or we run out
            while compiled_parts != all_parts and len(versions) > 0:
                latest_version = versions.pop(-1)
                regex = re.compile(pieces[i][latest_version]["regex_str"])
                version_path = f'library/{self.shelf}/scores/{i}/{latest_version}'
                version_path = os.path.abspath(version_path)
            
                parts = []
                part_files = os.listdir(version_path)

                for part in part_files:
                    slug = re.fullmatch(regex, part)
                    if slug is not None:
                        slug = slug.group(1)
                    if slug is not None and slug not in compiled_parts:
                        compiled_parts.add(slug)
                        parts.append(slug)

                if len(parts) > 0:
                    most_recent_parts[latest_version] = sorted(parts)

            print(most_recent_parts)
            


def cli_parser():
    parser = argparse.ArgumentParser(description='Score management tool.')
    subparsers = parser.add_subparsers(dest='command', help='Commands')

    init_parser = subparsers.add_parser('init', help='Init a new library')
    init_parser.add_argument('-s', '--shelf', required=False, help='Shelf name')
    init_parser.add_argument('-p', '--path', required=False, help='SharePoint directory to track')

    sync_parser = subparsers.add_parser('sync', help='Syncs local files with SharePoint')
    sync_parser.add_argument('shelf', nargs='?', help='Shelf name')

    diff_parser = subparsers.add_parser('diff', help='Generate diffes for printing')
    diff_parser.add_argument('shelf', nargs='?', help='Shelf name')
    diff_parser.add_argument('-a', '--after', required=False, help='Starting time to generate changes from (YYYY-MM-DD)')
    diff_parser.add_argument('-s', '--size', required=False, help='Approximate size of each batch (default: 100 pages)')

    binders_parser = subparsers.add_parser('binders', help='Generate binders')
    binders_parser.add_argument('shelf', nargs='?', help='Shelf name')

    update_parser = subparsers.add_parser('update', help='Generate diffs and binders simultaneously')
    update_parser.add_argument('shelf', nargs='?', help='Shelf name')
    update_parser.add_argument('-a', '--after', required=False, help='Starting time to generate changes from (YYYY-MM-DD)')
    update_parser.add_argument('-s', '--size', required=False, help='Approximate size of each batch (default: 100 pages)')

    check_parser = subparsers.add_parser('check', help='Check for unused parts in recent versions')
    check_parser.add_argument('shelf', nargs='?', help='Shelf name')
    check_parser.add_argument('-a', '--after', required=False, help='Starting time to generate changes from (YYYY-MM-DD)')

    compile_parser = subparsers.add_parser("compile", help="Compile the latest versions of each part in a piece")
    compile_parser.add_argument("shelf", nargs="?", help="Shelf name")

    # add_parser = subparsers.add_parser('add', help='Add score')
    # add_parser.add_argument('shelf', nargs='?', help='Shelf name')

    # rm_parser = subparsers.add_parser('remove', help='Remove score')
    # rm_parser.add_argument('shelf', nargs='?', help='Shelf name')

    return parser

def parse_date_arg(args, arg_name):
    # type: (argparse.Namespace, str) -> datetime|None
    date = getattr(args, arg_name)
    if date is not None:
        date = datetime.fromisoformat(date)
        if date.tzinfo is None:
            date = to_utc(date)
    return date

def parse_diff_args(args):
    # type: (argparse.Namespace) -> tuple[datetime|None, int]
    after = parse_date_arg(args, 'after')
    size = 100
    if args.size is not None:
        size = int(args.size)
    return after, size

if __name__ == '__main__':
    if not os.path.samefile(os.getcwd(), os.path.dirname(os.path.realpath(__file__))):
        print('ERROR: librarian can only be run in its own directory')
        sys.exit(1)
    parser = cli_parser()
    args = parser.parse_args()
    match args.command:
        case 'init':
            Librarian.init_shelf(args.shelf, args.path)
        case 'sync':
            librarian = Librarian(args.shelf)
            librarian.sync_shelf()
        case 'binders':
            librarian = Librarian(args.shelf)
            librarian.make_all_binders()
        case 'diff':
            librarian = Librarian(args.shelf)
            after, size = parse_diff_args(args)
            librarian.make_shelf_diffs(after=after, size=size)
        case 'check':
            librarian = Librarian(args.shelf)
            after = parse_date_arg(args, 'after')
            librarian.check_all_versions(after=after)
        case 'add':
            librarian = Librarian(args.shelf)
            # TODO: remember to update both metadata.json and binders.csv
        case 'remove':
            librarian = Librarian(args.shelf)
        case 'update':
            librarian = Librarian(args.shelf)
            after, size = parse_diff_args(args)
            librarian.update_shelf(after=after, size=size)
        case "compile":
            librarian = Librarian(args.shelf)
            librarian.compile_parts()
        case _:
            parser.print_help()
