# LeakQuest

A vibe-coded tool for searching WikiLeaks PlusD. LeakQuest design searches uses keyword search to identify relevant cables, fetches their full text, and then exports everything to a formatted Excel sheet. This Excel sheet also includes summary statistics, so users can understand the corpus they have generated from the keyword searches. 

Some additional notes; 
- LeakQuest is a poor way to explore the broad WikiLeaks PlusD corpus. Users that wish to make use of LeakQuest should first familiarise themselves with what's available in [PlusD](https://wikileaks.org/plusd/) and then use LeakQuest to conduct keyword searches for later analysis. 
- Secondly, LeakQuest is quite slow. WikiLeaks does not have an API, so it works on scraping. In order to not bombard the WikiLeaks servers, it is designed to scrape relatively slowly. Users can speed this up if they so wish. 
- Thirdly, LeakQuest has primarily been tested on the Cablegate collection, rather than the much larger alternative collections in PlusD. While it _should_ work just as well with the other collections, there may be errors. Please let me know if these arise.

## Download

Go to [Releases](https://github.com/E-Hayes-ANU/LeakQuest/releases) and download the version for your platform:

- **Windows:** Download `LeakQuest.exe` and run it directly — no installation needed.
- **Mac:** Download the zip, extract it, and double-click `LeakQuest.command`. On first launch it will set up a Python virtual environment and install dependencies automatically. See [Mac notes](#mac-notes) below.

## Usage

LeakQuest walks you through an interactive prompt:

### 1. Keywords

Enter search terms. Use quotes for exact phrases.

```
Keywords: nuclear proliferation
```

```
Keywords: "arms control" pakistan
```

### 2. Exclude Filter

Optional. Exclude cables containing specific words or phrases. Use quotes for exact phrases.

```
Exclude: classified "nuclear test" weapons
Apply to: both
```

Scope options:
- **both** - exclude from title and body (default). This runs in two passes: first, cables whose titles contain any exclude term are removed before fetching begins. Then, after full text is downloaded, cables whose body text contains any exclude term are removed in a second pass.
- **title** - exclude only by title. Filtered before fetching, so excluded cables are never downloaded.
- **body** - exclude only by full text. All cables are fetched first, then filtered.

### 3. Date Range

Optional. Format: `YYYY-MM-DD`.

```
From date: 1975-01-01
To date: 1979-12-31
```

### 4. Document Sets

Choose which cable collections to search:
- **cablegate** (default) - Cablegate cables only
- **all** - includes Kissinger Cables, Carter Cables, and other sets

### 5. Review and Fetch

LeakQuest shows the result count and a preview of the first 10 cables. Confirm to fetch the full text of every cable.

### 6. Combine Searches (Optional)

After fetching, you can run additional searches and merge results into a single export. Duplicate cables are automatically removed.

### 7. Export

Choose a filename and LeakQuest writes a formatted Excel file with two sheets:

**Cables** sheet columns:
- **WikiLeaks ID** - cable identifier
- **Date** - cable date (YYYY-MM-DD)
- **Title** - cable subject line
- **Full Text** - complete cable text, reflowed for readability
- **URL** - clickable link to the original cable on WikiLeaks

**Statistics** sheet:
- Summary: total cables, date range, and search keywords
- Table 1: distribution of cables by country of origin (sorted by count)
- Table 2: distribution of cables by year (sorted chronologically)

Country names are historically accurate — a 1980 cable from Moscow is attributed to the Soviet Union, not Russia. Embassy post codes are mapped to countries using a built-in lookup table of 250+ diplomatic posts, supplemented by metadata extracted from cable pages at fetch time.

## Examples

**Search for cables about energy in Kazakhstan:**
```
Keywords: energy kazakhstan
```

**Search for an exact phrase within a date range:**
```
Keywords: "nuclear proliferation"
From date: 1970-01-01
To date: 1979-12-31
```

**Combine two searches into one spreadsheet:**
1. Search `"arms control" pakistan`
2. When prompted "Add another search?", say yes
3. Search `"chemical weapons"`
4. Export both result sets to a single file

## Resuming Interrupted Fetches

If LeakQuest is interrupted while fetching cables (network error, closed window, etc.), it saves progress to a checkpoint file. On the next run with the same search, already-fetched cables are loaded from the checkpoint instead of re-downloaded.

## Troubleshooting

**"Failed to fetch search results"** - Make sure `curl` is installed and accessible from your terminal. On Windows 10+, it's included by default.

**Slow fetching** - LeakQuest waits 1.5 seconds between cable fetches to avoid overloading the server. A search with 100 results takes roughly 2.5 minutes to fetch.

**Excel text looks wrong** - Cable text is automatically reflowed to remove the original hard line wraps (~65 character width) while preserving paragraph breaks.

## Mac Notes

**Requirements:** Python 3.8+ (install via `xcode-select --install`, [Homebrew](https://brew.sh), or [python.org](https://www.python.org/downloads/)) and curl (ships with macOS).

**Setup:** Extract the zip, open Terminal, and make the launcher executable:
```
chmod +x LeakQuest-Mac/LeakQuest.command
```
Then double-click `LeakQuest.command` in Finder, or run it from Terminal. On first launch, it creates a virtual environment and installs dependencies automatically.

**Gatekeeper:** macOS may block the script on first run because it is unsigned. If you see a security warning, right-click (or Control-click) the file and choose **Open**, then click **Open** again in the dialog. You only need to do this once.

## Running from Source

If you prefer to run from source instead of using the pre-built downloads:

```
pip install requests beautifulsoup4 openpyxl rich
python leakquest.py
```

Requires Python 3.8+ and curl.
