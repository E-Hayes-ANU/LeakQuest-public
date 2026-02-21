"""LeakQuest - WikiLeaks Cablegate Search & Extract Tool.

Interactive CLI tool to search WikiLeaks Cablegate cables by keyword
and export results to Excel.
"""

import os
import re
import shlex
from datetime import datetime
from pathlib import Path

from rich.console import Console
from rich.panel import Panel
from rich.progress import Progress, SpinnerColumn, TextColumn, BarColumn, TaskProgressColumn, TimeRemainingColumn
from rich.prompt import Prompt, Confirm
from rich.table import Table

from scraper import search_cables, fetch_all_cables
from exporter import export_to_excel
from updater import check_for_update, download_update, apply_windows_update

__version__ = "2.2"

console = Console()


def print_banner():
    console.print(Panel(
        "[bold cyan]LeakQuest[/bold cyan]\n"
        f"[dim]v{__version__}[/dim]\n"
        "[dim]WikiLeaks PlusD Search & Extract Tool[/dim]\n"
        "[dim]Search 2+ million US diplomatic records (1966-2010)[/dim]",
        border_style="cyan",
    ))


def prompt_keywords():
    """Prompt for search keywords."""
    console.print("\n[bold]Step 1: Search Keywords[/bold]")
    console.print("[dim]Enter keywords to search for in cable text. Use quotes for exact phrases.[/dim]")
    keyword = Prompt.ask("[cyan]Keywords[/cyan]", default="").strip()
    return keyword if keyword else None


def prompt_exclude():
    """Prompt for negative search filter.

    Returns (terms, scope) where scope is 'both', 'title', or 'body',
    or (None, None) if no exclude filter.
    """
    console.print("\n[bold]Step 2: Exclude Filter (optional)[/bold]")
    console.print("[dim]Exclude cables containing these words/phrases. Use quotes for exact phrases.[/dim]")
    console.print("[dim]Example: classified \"nuclear test\" weapons[/dim]")
    exclude_input = Prompt.ask("[cyan]Exclude[/cyan]", default="").strip()

    if not exclude_input:
        return None, None

    # Parse space-separated terms, respecting quoted phrases
    try:
        terms = [t.lower() for t in shlex.split(exclude_input) if t.strip()]
    except ValueError:
        # Mismatched quotes — fall back to simple split
        terms = [t.strip().strip('"').lower() for t in exclude_input.split() if t.strip()]

    if not terms:
        return None, None

    scope = Prompt.ask(
        "[cyan]Apply to[/cyan]",
        choices=["both", "title", "body"],
        default="both",
    )
    return terms, scope


def prompt_date_range():
    """Prompt for optional date range."""
    console.print("\n[bold]Step 3: Date Range (optional)[/bold]")
    console.print("[dim]Format: YYYY-MM-DD. Leave empty for no date filter.[/dim]")

    date_from = Prompt.ask("[cyan]From date[/cyan]", default="").strip()
    date_to = Prompt.ask("[cyan]To date[/cyan]", default="").strip()

    # Validate dates are real calendar dates
    for label, val in [("from", date_from), ("to", date_to)]:
        if val:
            try:
                datetime.strptime(val, "%Y-%m-%d")
            except ValueError:
                console.print(f"[yellow]Invalid {label} date '{val}', ignoring.[/yellow]")
                if label == "from":
                    date_from = ""
                else:
                    date_to = ""

    return date_from or None, date_to or None


def prompt_projects():
    """Prompt for document sets."""
    console.print("\n[bold]Step 4: Document Sets[/bold]")
    console.print("[dim]Default: Cablegate only. Enter 'all' to include Kissinger/Carter cables too.[/dim]")

    choice = Prompt.ask("[cyan]Document set[/cyan]", choices=["cablegate", "all"], default="cablegate")
    if choice == "all":
        return ["cg", "cc", "fp", "ee", "ps"]
    return ["cg"]


def prompt_filename(keyword):
    """Prompt for output filename."""
    # Build default name from search keywords
    if keyword:
        name_part = re.sub(r"[^\w\s-]", "", keyword).strip().replace(" ", "_")[:30]
    else:
        name_part = "results"
    default_name = f"leakquest_{name_part}.xlsx"

    console.print(f"\n[bold]Step 5: Output File[/bold]")
    filename = Prompt.ask("[cyan]Output filename[/cyan]", default=default_name).strip()

    if not filename.endswith(".xlsx"):
        filename += ".xlsx"

    return filename


def do_search_and_fetch(search_num=1):
    """Run a single search-and-fetch cycle.

    Returns (cable_list, keyword) or (None, None) if cancelled/empty.
    """
    label = f" (search #{search_num})" if search_num > 1 else ""
    console.print(f"\n[bold]{'=' * 40}{label}[/bold]")

    keyword = prompt_keywords()

    if not keyword:
        console.print("[red]Error: You must provide search keywords.[/red]")
        return None, None

    exclude, exclude_scope = prompt_exclude()
    date_from, date_to = prompt_date_range()
    projects = prompt_projects()

    # Summary
    console.print("\n")
    summary = Table(title=f"Search Parameters{label}", show_header=False, border_style="cyan")
    summary.add_column("Parameter", style="bold")
    summary.add_column("Value")
    summary.add_row("Keywords", keyword)
    if exclude:
        exclude_display = f"{', '.join(exclude)} [dim]({exclude_scope})[/dim]"
    else:
        exclude_display = "[dim]none[/dim]"
    summary.add_row("Exclude", exclude_display)
    summary.add_row("Date range", f"{date_from or '...'} to {date_to or '...'}")
    summary.add_row("Document sets", ", ".join(projects))
    console.print(summary)

    if not Confirm.ask("\n[cyan]Start search?[/cyan]", default=True):
        console.print("[yellow]Cancelled.[/yellow]")
        return None, None

    # Search
    console.print()
    with console.status("[bold cyan]Searching WikiLeaks PlusD...", spinner="dots"):
        cable_list = search_cables(
            keyword=keyword,
            date_from=date_from,
            date_to=date_to,
            projects=projects,
            progress_callback=lambda msg: console.log(msg),
        )

    if not cable_list:
        console.print("[yellow]No cables found matching your search criteria.[/yellow]")
        return None, None

    # Apply exclude filter to titles
    if exclude and exclude_scope in ("both", "title"):
        before_count = len(cable_list)
        cable_list = [
            c for c in cable_list
            if not any(term in c.get("title", "").lower() for term in exclude)
        ]
        excluded = before_count - len(cable_list)
        if excluded:
            console.print(f"[yellow]Excluded {excluded} cables by title filter.[/yellow]")

    if not cable_list:
        console.print("[yellow]All cables were excluded by the filter.[/yellow]")
        return None, None

    # Confirm
    console.print(f"\n[bold green]Found {len(cable_list)} cables.[/bold green]")

    preview = Table(title="Preview (first 10)", show_header=True, header_style="bold")
    preview.add_column("Cable ID", style="cyan", width=22)
    preview.add_column("Date", width=12)
    preview.add_column("Title", max_width=60)
    for cable in cable_list[:10]:
        preview.add_row(cable["cable_id"], cable.get("date", ""), cable.get("title", ""))
    console.print(preview)

    # Estimate fetch time (1.5s rate-limit sleep + ~5.5s avg for the HTTP request)
    est_seconds = int(len(cable_list) * 7)
    if est_seconds >= 3600:
        hours = est_seconds // 3600
        minutes = (est_seconds % 3600) // 60
        est_display = f"{hours}h {minutes}m" if minutes else f"{hours}h"
    elif est_seconds >= 60:
        est_display = f"{est_seconds // 60}m {est_seconds % 60}s"
    else:
        est_display = f"{est_seconds}s"

    if est_seconds >= 3600:
        console.print(f"\n[bold yellow]Estimated fetch time: {est_display}[/bold yellow]")
        if not Confirm.ask(
            f"[cyan]This will take a while. Fetch all {len(cable_list)} cables?[/cyan]",
            default=False,
        ):
            console.print("[yellow]Cancelled.[/yellow]")
            return None, None
    else:
        if not Confirm.ask(
            f"\n[cyan]Fetch full text for all {len(cable_list)} cables? (est. {est_display})[/cyan]",
            default=True,
        ):
            console.print("[yellow]Cancelled.[/yellow]")
            return None, None

    # Fetch cables with progress bar
    checkpoint_file = f"checkpoint_{search_num}.json"
    fetched_cables = []

    console.print()
    with Progress(
        SpinnerColumn(),
        TextColumn("[progress.description]{task.description}"),
        BarColumn(),
        TaskProgressColumn(),
        TimeRemainingColumn(),
        console=console,
    ) as progress:
        task = progress.add_task("Fetching cables...", total=len(cable_list))

        for i, total, cable_data in fetch_all_cables(cable_list, checkpoint_file):
            fetched_cables.append(cable_data)
            progress.update(task, completed=i, description=f"Fetching {cable_data['cable_id']}...")

    # Report any fetch failures
    failed = [c for c in fetched_cables if "_fetch_error" in c]
    if failed:
        console.print(f"\n[bold yellow]Warning: {len(failed)} cable(s) failed to fetch:[/bold yellow]")
        for c in failed[:20]:
            console.print(f"  [yellow]- {c['cable_id']}: {c['_fetch_error']}[/yellow]")
        if len(failed) > 20:
            console.print(f"  [yellow]... and {len(failed) - 20} more[/yellow]")

    # Filter by full text
    if exclude and exclude_scope in ("both", "body"):
        before_count = len(fetched_cables)
        fetched_cables = [
            c for c in fetched_cables
            if not any(term in c.get("full_text", "").lower() for term in exclude)
        ]
        excluded = before_count - len(fetched_cables)
        if excluded:
            console.print(f"[yellow]Excluded {excluded} more cables by full text filter.[/yellow]")

    # Client-side date range filter (safety net — server may return out-of-range results)
    if date_from or date_to:
        before_count = len(fetched_cables)
        filtered = []
        for c in fetched_cables:
            cable_date = c.get("date", "")
            if not cable_date:
                filtered.append(c)  # keep cables with no date
                continue
            # Dates are already normalized to YYYY-MM-DD by scraper
            if date_from and cable_date < date_from:
                continue
            if date_to and cable_date > date_to:
                continue
            filtered.append(c)
        excluded_count = before_count - len(filtered)
        if excluded_count:
            console.print(f"[yellow]Excluded {excluded_count} cable(s) outside date range.[/yellow]")
        fetched_cables = filtered

    if not fetched_cables:
        console.print("[yellow]All cables were excluded by the filter.[/yellow]")
        return None, None

    console.print(f"[green]Collected {len(fetched_cables)} cables from this search.[/green]")
    return fetched_cables, keyword


def run_session():
    """Run one or more searches and export to a single Excel file."""
    all_cables = []
    all_keywords = []
    search_num = 0

    while True:
        search_num += 1
        result, kw = do_search_and_fetch(search_num)

        if result:
            if kw:
                all_keywords.append(kw)
            # Deduplicate against cables already collected
            existing_ids = {c["cable_id"] for c in all_cables}
            new_cables = [c for c in result if c["cable_id"] not in existing_ids]
            dupes = len(result) - len(new_cables)
            all_cables.extend(new_cables)
            if dupes:
                console.print(f"[yellow]Skipped {dupes} duplicate cables already in results.[/yellow]")
            console.print(
                f"[bold cyan]Total cables collected so far: {len(all_cables)}[/bold cyan]"
            )

        if not all_cables:
            console.print("[yellow]No cables collected yet.[/yellow]")
            if not Confirm.ask("[cyan]Try another search?[/cyan]", default=True):
                return
            continue

        if not Confirm.ask("[cyan]Add another search to this export?[/cyan]", default=False):
            break
        console.print()

    # Prompt for filename
    first_kw = all_keywords[0] if all_keywords else None
    filename = prompt_filename(first_kw)

    # Export
    console.print(f"\n[bold cyan]Exporting to {filename}...[/bold cyan]")
    count = export_to_excel(all_cables, filename, keywords=all_keywords)

    # Clean up checkpoint files only after successful export
    for i in range(1, search_num + 1):
        cp = Path(f"checkpoint_{i}.json")
        if cp.exists():
            cp.unlink()

    console.print()
    console.print(Panel(
        f"[bold green]Done![/bold green]\n\n"
        f"Exported [bold]{count}[/bold] cables to [bold cyan]{filename}[/bold cyan]\n"
        f"File size: {os.path.getsize(filename) / 1024:.1f} KB",
        title="Export Complete",
        border_style="green",
    ))


def _check_and_apply_update():
    """Check for updates and offer to download. Returns True if app should exit."""
    import sys

    update = check_for_update(__version__)
    if not update:
        return False

    new_version, download_url, asset_name = update
    console.print(
        f"\n[bold yellow]A new version is available: v{new_version}[/bold yellow] "
        f"[dim](current: v{__version__})[/dim]"
    )

    if not Confirm.ask("[cyan]Download update?[/cyan]", default=True):
        return False

    dest_path = download_update(download_url, asset_name, console)
    if not dest_path:
        return False

    if sys.platform == "win32":
        apply_windows_update(dest_path, console)
        return True  # signal main() to exit
    else:
        console.print(
            f"\n[bold green]Update downloaded to:[/bold green] {dest_path}\n"
            "[dim]Replace the current files to complete the update.[/dim]"
        )
        return False


def main():
    print_banner()

    if _check_and_apply_update():
        return

    while True:
        try:
            run_session()
        except KeyboardInterrupt:
            console.print("\n[yellow]Interrupted.[/yellow]")
        except Exception:
            console.print_exception()

        console.print()
        if not Confirm.ask("[cyan]Start a new session?[/cyan]", default=True):
            console.print("[bold cyan]Goodbye![/bold cyan]")
            break
        console.print("\n" + "=" * 60 + "\n")


if __name__ == "__main__":
    try:
        main()
    except Exception:
        console.print_exception()
    finally:
        input("\nPress Enter to exit...")
