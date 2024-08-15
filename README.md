# KeyTrack

A simple key management system that uses Python, Tkinter, and OpenPyXL to manage and track keys for apartment rooms. This system allows users to collect, return, report lost, and borrow spare keys, all while logging the actions with timestamps.

## Features

- Collect keys with timestamp logging.
- Return keys with timestamp logging and tracking.
- Report lost keys with timestamp logging.
- Borrow spare keys with timestamp logging.
- Graphical User Interface (GUI) for easy interaction.

## Prerequisites

- Python 3.x
- `openpyxl` library
- `tkinter` library (included with Python standard library)

## Installation

1. Clone the repository:

    ```bash
    git clone https://github.com/daviddagyei/KeyTrack.git
    cd KeyTrack
    ```

2. Install the required Python packages:

    ```bash
    pip install openpyxl
    ```

3. Ensure you have an Excel file named `key_distribution.xlsx` with the following columns:

    - Room Number
    - Available Keys
    - Collected By
    - Lost Keys
    - Borrowed Spare Keys
    - Returned Keys

## Usage

Run the script:

```bash
git pull
python keytrack.py
```
## Pushing Updates

After making changes (collecting, returning, reporting lost, borrowing spare keys), commit and push the updated key_distribution.xlsx file to the remote repository to keep the data synced:
1. Add the changes to staging:
   ```bash
   git add key_distribution.xlsx
   ```
2. Commit the changes with a message:
   ```bash
   git commit -m "Updated key distribution data"
   ```
3. Push the changes to the remote repository:
   ```bash
   git push origin main




