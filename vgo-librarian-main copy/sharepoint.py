from dataclasses import dataclass
from datetime import datetime
from enum import IntEnum
import json
import os
import pickle
from time import sleep

from dotenv import dotenv_values
import requests

class SharePointFileType(IntEnum):
    FILE=0
    FOLDER=1

@dataclass
class SharePointFile:
    file_type: SharePointFileType
    name: str
    full_path: str
    item_url: str
    created_at: datetime
    modified_at: datetime

    @staticmethod
    def from_listdir(row):
        item_url = row['.spItemUrl']
        if '?' in item_url:
            item_url = item_url[:item_url.index('?')]

        return SharePointFile(
            file_type=SharePointFileType(int(row['FSObjType'])),
            full_path=row['FileRef'],
            name=row['FileLeafRef'],
            item_url=item_url,
            created_at=datetime.fromisoformat(row['Created_x0020_Date.']),
            modified_at=datetime.fromisoformat(row['Modified.']),
        )

class SharePointClient:
    def __init__(self, cookie_file='cookies.pickle', env_file='example.env'):
        # type: (str, str) -> None
        self.cookie_file = cookie_file
        self.env_file = env_file
        self.session = requests.Session()

        # init environment variables
        env = dotenv_values(self.env_file)
        username = env['MIT_USERNAME']
        if username is None:
            raise ValueError(f'Please set MIT_USERNAME in {self.env_file}')
        if not username.endswith('@mit.edu'):
            username += '@mit.edu'
        password = env['MIT_PASSWORD']
        if password is None:
            raise ValueError(f'Please set MIT_PASSWORD in {self.env_file}')
        sharepoint_url = env['MIT_SHAREPOINT_URL']
        if sharepoint_url is None:
            raise ValueError(f'Please set MIT_SHAREPOINT_URL in {self.env_file}')
        if sharepoint_url and sharepoint_url[-1] == '/':
            sharepoint_url = sharepoint_url[:-1]
        self.username = username
        self.password = password
        self.sharepoint_url = sharepoint_url
        self.site_name = sharepoint_url.rsplit('/', 1)[-1]

        # these are filled out with the test query
        self.access_token = 'access_token='
        self.media_base_url = ''
        # load session cookies
        if os.path.isfile('cookies.pickle'):
            with open('cookies.pickle', 'rb') as f:
                self.session.cookies.update(pickle.load(f))
        else:
            self.reset_session()
        # repeatedly try test query until success
        while True:
            try:
                self.listdir('')
            except ValueError:
                self.session.cookies.clear()
                self.reset_session()
            else:
                break

    def reset_session(self):
        cookies_dict = self._get_auth_cookies()
        for cookie in cookies_dict:
            if 'name' in cookie and 'value' in cookie:
                valid_keys = { 'domain', 'path', 'expires' }
                kwargs = { k: cookie[k] for k in cookie.keys() & valid_keys }
                self.session.cookies.set(
                    name=cookie['name'],
                    value=cookie['value'],
                    **kwargs
                )
        with open('cookies.pickle', 'wb') as f:
            pickle.dump(self.session.cookies, f)

    def download_file(self, file):
        # type: (SharePointFile|str) -> bytes
        if isinstance(file, SharePointFile):
            if file.file_type != SharePointFileType.FILE:
                raise ValueError('Cannot call download_file on a folder')
            file_path = file.full_path
        else:
            file_path = file
        response = self.session.get(f'{self.sharepoint_url}/_layouts/15/download.aspx', 
            params={
                'SourceUrl': file_path,
            })
        if 'Content-Type' in response.headers and response.headers['Content-Type'].startswith('text/html'):
            raise ValueError('File not found')
        return response.content

    def download_zip(self, folder):
        # type: (SharePointFile) -> bytes
        base_name = os.path.basename(folder.full_path)
        doc_id = f'{folder.item_url}?version=Published&{self.access_token}'
        file_data = json.dumps({
            'items': [{
                'name': base_name,
                'size': 0,
                'docId': doc_id,
                'isFolder': folder.file_type == SharePointFileType.FOLDER,
            }]
        })
        # TODO: use https://stackoverflow.com/a/37573701
        response = self.session.post(f'{self.media_base_url}/transform/zip?cs=fFNQTw',
            headers={
            },
            json={
                'zipFileName': f'{base_name}.zip',
                'provider': 'spo',
                'files': file_data,
            })
        return response.content

    # useful for checking activity within a folder
    def get_last_activity_date(self, file):
        # type: (SharePointFile) -> datetime
        response = self.session.get(f'{file.item_url}/activities',
            params={
                '$expand': 'driveItem($select=id,name,webUrl,parentReference,file,folder)',
                '$top': '1'
            },
            headers={
                # generic user-agent
                'User-Agent': 'Mozilla/5.0 (X11; Linux x86_64; rv:139.0) Gecko/20100101 Firefox/139.0',
            },
        )
        response_json = response.json()
        if 'error' in response_json:
            raise ValueError(response_json['error']['message'])
        date_string = response_json['value'][0]['times']['recordedTime']
        return datetime.fromisoformat(date_string)

    def listdir(self, path):
        if path and path[-1] == '/':
            path = path[:-1]
        base_path = f'/sites/{self.site_name}/Shared Documents'
        full_path =  os.path.join(base_path, path)
        response = self.session.post(f'{self.sharepoint_url}/_api/web/GetListUsingPath(DecodedUrl=@a1)/RenderListDataAsStream',
            params={
                '@a1': f"'{base_path}'",
                'RootFolder': f'{full_path}',
            },
            headers={
                'Content-Type': 'application/json;odata=verbose',
            },
            json={
                'parameters':{
                    '__metadata': { 'type':'SP.RenderListDataParameters' },
                    'RenderOptions': 0b1_0000_0000_0111,
                    'AllowMultipleValueFilterForTaxonomyFields':True,
                    'AddRequiredFields':True,
                    'RequireFolderColoringFields':True
                }
            })
        if 'Attempted to perform an unauthorized operation' in response.text:
            raise ValueError('Invalid credentials')
        if 'File Not Found' in response.text:
            raise ValueError('File not found')
        response_json = response.json()
        self.access_token = response_json['ListSchema']['.driveAccessToken']
        self.media_base_url = response_json['ListSchema']['.mediaBaseUrl']
        rows = response_json['ListData']['Row']
        return [SharePointFile.from_listdir(row) for row in rows]

    def _get_auth_cookies(self):
        from playwright.sync_api import sync_playwright

        with sync_playwright() as p:
            print('Automatically logging into SharePoint...')
            browser = p.chromium.launch(headless=False)
            page = browser.new_page()
            # redirect to office365 login
            page.goto(self.sharepoint_url)
            page.wait_for_load_state('networkidle')
            # fill in username
            page.fill('#i0116', self.username)
            page.locator('#idSIButton9').click()
            print('Redirecting to Touchstone...')
            page.wait_for_url('**/okta.mit.edu/**')
            page.wait_for_load_state('networkidle')
            # fill in password
            page.fill('#input28', self.password)
            page.click('input.button.button-primary[type="submit"][value="Verify"][data-type="save"]')
            page.wait_for_url(lambda url: 'duosecurity.com' in url)
            page.wait_for_load_state('networkidle')
            print('Waiting for 2fa...')
            while page.locator('#trust-browser-button').count() == 0:
                sleep(0.2)
            page.locator('#trust-browser-button').click()
            page.wait_for_url(self.sharepoint_url)
            cookies = page.context.cookies()
            browser.close()
            return cookies
