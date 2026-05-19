# librarian

NOTE: All credit for this script goes to Justin Yu (fractal161). I only made a few minor edits to it to make it more functional.

The [MIT Video Game Orchestra](https://www.youtube.com/@mitvgo/) performs a large amount of music every semester, which is rehearsed weekly. The arrangements themselves are a core part of the process, as they are frequently modified based on feedback from rehearsals. As a result, updated arrangements must be printed and arranged into binders every week. When done manually, this takes considerable effort (citation: I have done this).

This repository streamlines process by automatically organizing and combining each part into their appropriate binder, batching the outputs for easy printing.

## Setup

This project is written entirely in [Python](https://www.python.org/), which must be installed, alongside [pip](https://pypi.org/project/pip/), its package manager.[^1] Once the repository is downloaded, run `pip install -r requirements.txt` to install the requirements. From here, run the program with `python3 librarian.py`.

[^1]: This should be included with Python by default, but if not, refer to the [installation guide](https://pip.pypa.io/en/stable/installation/).

## Usage

`librarian` organizes its scores into *shelves*, which are collections of arrangements to track over time (e.g. you might make one shelf per semester).

### Initial setup

Once the initial arrangements have been uploaded to SharePoint, run `python librarian.py init` and follow the prompts. The most intensive part of this process is in making the regexes to extract instruments from file names (you can use [https://regex101.com/] for testing).

The results can be found inside the `library/` folder. The last step is manually filling out the `binders.csv` file (which can be opened using e.g. Excel, or directly copied from and external Google Sheet). Make sure that the instrument names you fill out for each box in binders.csv matches the instrument part name in the arrangement (ex. If The Tomorrow With You's Clarinet 1 part is called "Clarinet_in_Bb_1", make sure that's what you enter into the corresponding box in binders.csv).

### Weekly usage

For most cases, you can run `python librarian.py update`, which will do the following:
- Download new versions from sharepoint
- Create updated "virtual binders" for all virtual players
- Generate batched pdf files (diffs) for printing any updates for physical binders

## Considerations

- The input handling isn't very robust, so there may be unexpected effects caused by typos (this is mostly n- and m-dashes)
- The `add` and `remove` commands are unimplemented. This can be done by manually editing the `metadata.json` and `binders.csv`, but this is nice to have.
- It might be nice to sync the `binders.csv` with a Google Sheet
- The binders.csv file must be a csv file and have data from a csv file otherwise the csv-reader package used can't access the data
- A lot of the code is "inefficient" in the sense that certain resources are recomputed several times, particularly with the `binders` and `diff` commands. In practice, the scope of each task should be small enough that this is unnoticeable, but this may be worth revisiting in the future.
- The sharepoint file structure needs to be exactly how it "normally" is
- All subfolders within a score's folder are treated as versions, ordered by the creation date of said subfolders. This means there shouldn't be any extraneous folders
- It is the arrangers' responsibility to insert cover pages
