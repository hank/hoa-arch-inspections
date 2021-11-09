from __future__ import print_function
import sys
import pickle
import os
from parse_homeowners import parse_homeowners
from pdf_convert import pdf_convert

from apiclient import discovery
#from oauth2client import client
#from oauth2client import tools
#from oauth2client.file import Storage
import google.auth
from google_auth_oauthlib.helpers import session_from_client_secrets_file, credentials_from_session
from google_auth_oauthlib.flow import Flow

# If modifying these scopes, delete your previously saved credentials
# at ~/.credentials/drive-python-quickstart.json
SCOPE = 'https://www.googleapis.com/auth/drive'
CLIENT_SECRET_FILE = 'client_secret.json'
APPLICATION_NAME = 'Drive API Python Quickstart'


def get_credentials():
    """Gets valid user credentials from storage.

    If nothing has been stored, or if the stored credentials are invalid,
    the OAuth2 flow is completed to obtain the new credentials.

    Returns:
        Credentials, the obtained credential.
    """
    try:
        # Save the credentials for the next run
        with open('token.pickle', 'rb') as token:
            creds = pickle.load(token)
            return creds
    except:
        print("Failed to load cached credentials")
    flow = Flow.from_client_secrets_file(CLIENT_SECRET_FILE, scopes=[SCOPE], redirect_uri='urn:ietf:wg:oauth:2.0:oob')
    auth_url, _ = flow.authorization_url(prompt='consent')

    print('Please go to this URL: {}'.format(auth_url))

    # The user will get an authorization code. This code is used to get the
    # access token.
    code = input('Enter the authorization code: ')
    flow.fetch_token(code=code)
    # Save the credentials for the next run
    with open('token.pickle', 'wb') as token:
         pickle.dump(flow.credentials, token)
    return flow.credentials

def main(homeowners_fpath, outdir, fmat="TXT"):
    """Shows basic usage of the Google Drive API.

    Creates a Google Drive API service object and outputs the names and IDs
    for up to 10 files.
    """
    credentials = get_credentials()
    service = discovery.build('drive', 'v3', credentials=credentials)

    if not os.path.isfile('out.csv'):

        try:
            results = service.files().list(
                pageSize=10,fields="nextPageToken, files(id, name, mimeType)").execute()
        except google.auth.exceptions.RefreshError:
            print("Token expired, getting it again")
            os.unlink('token.pickle')
            credentials = get_credentials()
            results = service.files().list(
                pageSize=10,fields="nextPageToken, files(id, name, mimeType)").execute()
        items = results.get('files', [])
        if not items:
            print('No files found.')
        else:
            print('Files:')
            i = 0
            for item in items:
                print('{0}: {1} ({2})'.format(i, item['name'], item['id']))
                i += 1
        num = int(input("Item number: ").strip())
        results = service.files().export(fileId=items[num]['id'], mimeType='text/csv').execute()
        with open("out.csv", "wb") as f:
            f.write(results)
    else:
        print("Using existing out.csv. Delete to refresh from Google Drive")
    # print(results)
    import csv
    # import io
    # csvIn = io.StringIO(str(results), newline='')
    with open("out.csv", "r", encoding="utf-8") as f:
        reader = csv.reader(f)
        reader = list(reader)
    labels = reader[0]

    with open("letter_template.txt", "r") as template:
        letter_template = template.read()

    try:
        os.makedirs(outdir)
    except:
        pass

    homeowners = parse_homeowners(homeowners_fpath)
    addresses = {x['Street Address'].strip(): x['Last Name'] for x in homeowners}
    for row in reader[1:]:
        # Attempt to find address in homeowners
        house_number = row[labels.index('House Number')].strip()
        street = row[labels.index('Street')].strip()
        if "Latchlift" in street or "Graf" in street or "Jud" in street:
            street += " Court"
        elif "Summer" in street:
            street += " Terrace"
        elif "Bonnie" in street:
            street += " Lane"
        street = street.replace("Grafton's", "Graftons")
        address = f"{house_number} {street}"
        if address not in addresses:
            raise RuntimeError("ERROR, couldn't find {} in addresses".format(address))
        lastname = addresses[address].strip()
        insp_arr = []
        for idx, field in enumerate(row):
            if len(field) == 0:
                field = "None"
            insp_arr.append("{}: {}".format(labels[idx].strip(), field.strip()))
        inspection = "\n".join(insp_arr)
        message = letter_template.format(lastname=lastname, inspection=inspection)
        # Open a file with the address as the name and add fields to it
        fname = (address.replace(r' ', '-').strip() + ".{}".format(fmat))
        if "pdf" == fmat:
            fullfname = os.path.join("inspections", fname)
            print("Creating {}".format(fullfname))
            pdf_convert(fullfname, message)

        else:
            print("Writing {}".format(fname))
            with open(os.path.join("inspections", fname), 'w') as f:
                f.write(message)
    if os.path.exists("out.csv"):
        print("Deleting temporary CSV")
        os.remove("out.csv")

if __name__ == '__main__':
    import argparse
    parser = argparse.ArgumentParser()
    parser.add_argument("homeowners", help="CSV file containing homeowners information. First row needs the column names.")
    parser.add_argument("--outdir", help="Output directory. Files will be overwritten if existing", default="inspections")
    parser.add_argument("--pdf", help="Make PDFs instead of TXT", action="store_true")
    args = parser.parse_args()
    if args.pdf:
        fmat="pdf"
    else:
        fmat="txt"
    main(args.homeowners, args.outdir, fmat=fmat)
