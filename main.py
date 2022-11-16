import os
from dotenv import load_dotenv
from pathlib import Path
from asposecellscloud.apis import cells_api


def convert_file():
    """Convert xlsm file chart to png with aspose."""

    # Load api credentials from environment
    load_dotenv()
    api = cells_api.CellsApi(
        os.getenv('CLIENTID'),
        os.getenv('CLIENTSECRET'), "v3.0"
    )

    # Filepath
    name = "test.xlsm"
    in_path = Path(__file__).parent / name

    # Upload file to aspose.cloud
    api.upload_file(name, in_path)

    # Set variables for conversion
    sheet_name = 'G1'
    vertical_resolution = 100
    horizontal_resolution = 90
    file_format = "png"

    # Convert file and write result to root directory
    result = api.cells_worksheets_get_worksheet(
        name, sheet_name, format=file_format,
        vertical_resolution=vertical_resolution,
        horizontal_resolution=horizontal_resolution,
        _preload_content=False)
    with open(Path(__file__).parent / 'result.png', 'wb') as f:
        f.write(result.read())


if __name__ == '__main__':
    convert_file()

