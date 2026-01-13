import os
from dotenv import load_dotenv

load_dotenv()

from dropbox_uploader import upload_eob_workbook


class Log:
    def info(self, m):
        print("INFO:", m)

    def warning(self, m):
        print("WARN:", m)


def main():
    # Ensure defaults for gating/allowed root
    os.environ["DROPBOX_ENABLE"] = os.getenv("DROPBOX_ENABLE", "1")
    os.environ["DROPBOX_ALLOWED_ROOT"] = os.getenv(
        "DROPBOX_ALLOWED_ROOT",
        "/CostSeg Team Folder/Mark/Test Client Master",
    )

    upload_eob_workbook(
        file_bytes=b"hello",
        filename="__dropbox_smoketest.txt",
        client_name="Patrick Gill",
        year="2025",
        property_address="184 Canyon Creek Trl",
        logger=Log(),
    )


if __name__ == "__main__":
    main()
