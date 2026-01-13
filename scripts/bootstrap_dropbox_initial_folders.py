import os
from dotenv import load_dotenv

load_dotenv()

from dropbox_uploader import DropboxClient, _norm_path

ROOT = "/CostSeg Team Folder/Mark/Test Client Master"


def main():
    token = os.getenv("DROPBOX_ACCESS_TOKEN", "").strip()
    if not token:
        raise SystemExit("Missing DROPBOX_ACCESS_TOKEN")

    dbx = DropboxClient(access_token=token)

    base = _norm_path(f"{ROOT}/Client Documents")
    dbx.ensure_parents(base)
    dbx.create_folder(base)

    for c in "ABCDEFGHIJKLMNOPQRSTUVWXYZ":
        p = _norm_path(f"{base}/{c}")
        dbx.create_folder(p)
        print("OK:", p)

    print("\nDone. Set:")
    print(f'DROPBOX_ALLOWED_ROOT="{ROOT}"')


if __name__ == "__main__":
    main()
